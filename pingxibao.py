import os
import sys  # 导入sys库，用于定位打包后的资源路径
import time
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import configparser
import docx
import fitz  # PyMuPDF
from openai import OpenAI
import threading
import queue

# ==============================================================================
# --- 配置区 ---
# ==============================================================================
CONFIG_FILE = 'config.ini'
EXAMPLE_RUBRIC_FILE = 'example_rubric.txt'  # 定义示例文件名
OUTPUT_FOLDER_NAME = "graded_feedback"
MODEL_NAME = "deepseek-chat"

# ==============================================================================
# --- Prompt框架 (已包含防截断指令) ---
# ==============================================================================
PROMPT_FRAME = """
# 角色
你是一名严谨、经验丰富的大学实验课程助教。

# 任务
请严格根据我接下来提供的【评分标准】，批改学生提交的实验报告。你需要：
1.  对报告的各个部分进行详细评价。
2.  总结报告的主要优点和待改进之处。
3.  根据各项表现和【评分标准】中的分值，给出一个建议的百分制分数。

# 【评分标准】
---
{user_rubric}
---

# 输出格式与规则
请严格遵守以下所有规则：
1.  **【最重要】必须只输出纯文本。绝对禁止使用任何Markdown、HTML、XML或任何其他标记语言（例如，不要使用 #, *, **, ``, <sup> 等符号）。**
2.  对于公式，请使用普通字符来表示，例如 "2^-delta_delta_Ct"。
3.  严格按照以下格式进行输出，使用等号长线作为分隔符：

====================================
综合评价:
[在这里用一句话总结报告的整体水平]

分项评语:
  - [评分标准中的第一项名称]: [在此处填写对该项的评价...]
  - [评分标准中的第二项名称]: [在此处填写对该项的评价...]
  - [以此类推，根据评分标准列出所有项...]

主要优点:
  - [在此处分点列出报告的优点]

主要待改进点:
  - [在此处分点列出具体的、可操作的修改建议]

建议分数:
[在此处给出一个具体的百分制分数，例如：88/100]
====================================

现在，请开始批改这份报告：
"""


# ==============================================================================
# --- 核心功能函数 ---
# ==============================================================================

def load_api_key():
    """从config.ini文件中加载API Key。如果文件或Key不存在，则创建模板并提示用户。"""
    config = configparser.ConfigParser()
    if not os.path.exists(CONFIG_FILE):
        config['API'] = {'DEEPSEEK_API_KEY': 'YOUR_KEY_GOES_HERE'}
        try:
            with open(CONFIG_FILE, 'w') as configfile:
                config.write(configfile)
            messagebox.showinfo("首次运行提示",
                                f"已为您创建配置文件 '{CONFIG_FILE}'。\n\n请打开该文件，填入您的DeepSeek API Key后，重新运行程序。")
        except Exception as e:
            messagebox.showerror("错误", f"创建配置文件失败: {e}")
        return None
    try:
        config.read(CONFIG_FILE)
        api_key = config.get('API', 'DEEPSEEK_API_KEY', fallback=None)
    except configparser.Error as e:
        messagebox.showerror("配置错误", f"读取 '{CONFIG_FILE}' 文件时出错: {e}\n\n请检查文件格式是否正确。")
        return None
    if not api_key or api_key == 'YOUR_KEY_GOES_HERE':
        messagebox.showwarning("API Key未设置", f"请在 '{CONFIG_FILE}' 文件中设置您的API Key。")
        return None
    return api_key


def get_user_input_with_gui(parent):
    """在一个Toplevel窗口中让用户输入评分标准。"""
    rubric_text = ""
    dialog = tk.Toplevel(parent)
    dialog.title("步骤1: 输入本次实验的评分标准")
    dialog.geometry("800x600")
    dialog.attributes('-topmost', True)

    def on_submit():
        nonlocal rubric_text
        rubric_text = text_area.get("1.0", "end-1c").strip()
        if rubric_text:
            dialog.destroy()
        else:
            messagebox.showwarning("输入错误", "评分标准不能为空！", parent=dialog)

    tk.Label(dialog, text="请在下方文本框中输入或粘贴本次批阅的评分标准 (Rubric):", font=("Arial", 12)).pack(pady=10)
    text_area = scrolledtext.ScrolledText(dialog, wrap=tk.WORD, font=("Arial", 11))
    text_area.pack(expand=True, fill='both', padx=10, pady=5)

    try:
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(".")
        example_file_path = os.path.join(base_path, EXAMPLE_RUBRIC_FILE)
        with open(example_file_path, 'r', encoding='utf-8') as f:
            example_rubric = f.read()
        text_area.insert(tk.INSERT, example_rubric)
    except Exception as e:
        print(f"加载示例评分标准失败: {e}")
        text_area.insert(tk.INSERT, "# 在此粘贴您的评分标准...")

    tk.Button(dialog, text="确认评分标准并继续", command=on_submit, font=("Arial", 12), bg="#4CAF50", fg="white").pack(
        pady=10)
    parent.wait_window(dialog)
    return rubric_text


def select_input_folder():
    """弹出文件对话框让用户选择文件夹。"""
    folder_path = filedialog.askdirectory(title="步骤2: 请选择包含实验报告的文件夹")
    return folder_path


def clean_ai_response(text):
    """一个备用的清理函数，以防AI偶尔不完全遵守格式。"""
    if not isinstance(text, str): return ""
    text = text.replace('### ', '').replace('**', '').replace('* ', '  - ')
    return text.strip()


def extract_text_from_file(file_path):
    """根据文件扩展名，从.docx或.pdf文件中提取纯文本。"""
    if file_path.lower().endswith(".docx"):
        try:
            doc = docx.Document(file_path)
            return "\n".join(para.text for para in doc.paragraphs)
        except Exception as e:
            return f"Error: 读取docx文件'{os.path.basename(file_path)}'时出错: {e}"
    elif file_path.lower().endswith(".pdf"):
        try:
            with fitz.open(file_path) as doc:
                return "".join(page.get_text() for page in doc)
        except Exception as e:
            return f"Error: 读取pdf文件'{os.path.basename(file_path)}'时出错: {e}"
    else:
        return f"Error: 不支持的文件格式: {os.path.basename(file_path)}"


def grade_lab_report(report_text, client, model_name, user_rubric):
    """调用DeepSeek API来批改实验报告，并增加超时设置。"""
    full_prompt = PROMPT_FRAME.format(user_rubric=user_rubric) + "\n\n" + report_text
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "user", "content": full_prompt}],
            temperature=0.2,
            max_tokens=2000,
            stream=False,
            # === 关键修改：增加超时参数 ===
            # 设置一个较长的超时时间，例如60秒。默认值通常较短。
            timeout=60.0
        )
        return response.choices[0].message.content
    except Exception as e:
        # 捕获并返回更详细的错误信息，便于调试
        return f"Error: 调用API时发生错误: {str(e)}"


def batch_grading_worker(input_folder, user_rubric, client, progress_queue):
    """这是在后台线程中运行的实际批阅工作函数。"""
    try:
        files_to_grade = [f for f in os.listdir(input_folder) if f.lower().endswith(('.docx', '.pdf'))]
        if not files_to_grade:
            progress_queue.put("FINISH:在指定文件夹中没有找到任何.docx或.pdf文件。")
            return

        total_files = len(files_to_grade)
        success_count, error_count = 0, 0
        output_folder = os.path.join(os.path.dirname(input_folder) or '.', OUTPUT_FOLDER_NAME)
        os.makedirs(output_folder, exist_ok=True)

        for i, filename in enumerate(files_to_grade):
            progress_queue.put(f"正在处理: {i + 1}/{total_files} - {filename}")
            file_path = os.path.join(input_folder, filename)
            report_content = extract_text_from_file(file_path)
            if report_content.startswith("Error:") or not report_content.strip():
                print(f"跳过文件 {filename}: {report_content}")
                error_count += 1
                continue

            ai_raw_feedback = grade_lab_report(report_content, client, MODEL_NAME, user_rubric)
            if ai_raw_feedback.startswith("Error:"):
                print(f"文件 {filename} 的AI处理失败: {ai_raw_feedback}")
                error_count += 1
                continue

            cleaned_feedback = clean_ai_response(ai_raw_feedback)
            base_name = os.path.splitext(filename)[0]
            output_filename = os.path.join(output_folder, f"评语_{base_name}.txt")
            try:
                with open(output_filename, 'w', encoding='utf-8') as f:
                    f.write(cleaned_feedback)
                success_count += 1
            except Exception as e:
                print(f"保存文件 {output_filename} 失败: {e}")
                error_count += 1

            time.sleep(1)

        final_message = f"所有报告批阅完成！\n\n成功: {success_count} 份\n失败: {error_count} 份"
        progress_queue.put(f"FINISH:{final_message}")
    except Exception as e:
        progress_queue.put(f"FINISH:处理过程中发生意外错误: {e}")


def main():
    """主函数，负责GUI和启动工作线程。"""
    root = tk.Tk()
    root.withdraw()
    print("--- 实验报告评析宝 启动 ---")

    api_key = load_api_key()
    if not api_key:
        print("API Key未配置，程序退出。")
        return
    print("API Key加载成功。")

    user_rubric = get_user_input_with_gui(root)
    if not user_rubric:
        print("未输入评分标准，程序退出。")
        return
    print("评分标准已确认。")

    input_folder = select_input_folder()
    if not input_folder:
        print("您没有选择任何文件夹，程序已退出。")
        return
    print(f"已选择报告文件夹: {input_folder}")

    try:
        client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
    except Exception as e:
        messagebox.showerror("API客户端错误", f"初始化API客户端失败: {e}")
        return

    loading_window = tk.Toplevel(root)
    loading_window.title("正在处理中...")
    loading_window.geometry("400x150")
    loading_window.resizable(False, False)
    loading_window.attributes('-topmost', True)

    progress_text = tk.StringVar(value="准备开始处理...")
    tk.Label(loading_window, text="请稍候，AI正在努力工作中...", font=("Arial", 14)).pack(pady=20)
    tk.Label(loading_window, textvariable=progress_text, font=("Arial", 10), wraplength=380).pack(pady=10)

    progress_queue = queue.Queue()
    worker_thread = threading.Thread(target=batch_grading_worker,
                                     args=(input_folder, user_rubric, client, progress_queue))
    worker_thread.start()

    def update_progress():
        try:
            message = progress_queue.get_nowait()
            if message.startswith("FINISH:"):
                loading_window.destroy()
                final_summary = message.split(":", 1)[1]
                messagebox.showinfo("任务完成", final_summary)
                root.quit()
            else:
                progress_text.set(message)
                root.after(100, update_progress)
        except queue.Empty:
            root.after(100, update_progress)

    root.after(100, update_progress)
    root.mainloop()

    print("--- 评析宝运行结束 ---")


if __name__ == "__main__":
    main()
