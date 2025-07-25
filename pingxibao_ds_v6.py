import os
import sys
import time
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
# configparser 库已不再需要
import docx
import fitz
from openai import OpenAI
import threading
import queue  # <-- 新增导入

# ==============================================================================
# --- 配置区 ---
# ==============================================================================

DEEPSEEK_API_KEY = "sk-d76e22b5011342c4a99be7a5e88dc3be"

EXAMPLE_RUBRIC_FILE = 'example_rubric.txt'
OUTPUT_FOLDER_NAME = "graded_feedback"
MODEL_NAME = "deepseek-chat"

# ==============================================================================
# --- Prompt框架---
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

def get_user_input_with_gui(parent):
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
    folder_path = filedialog.askdirectory(title="步骤2: 请选择包含实验报告的文件夹")
    return folder_path


def clean_ai_response(text):
    if not isinstance(text, str): return ""
    text = text.replace('### ', '').replace('**', '').replace('* ', '  - ')
    return text.strip()


def extract_text_from_file(file_path):
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
    full_prompt = PROMPT_FRAME.format(user_rubric=user_rubric) + "\n\n" + report_text
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "user", "content": full_prompt}],
            temperature=0.2, max_tokens=2000, stream=False, timeout=60.0
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error: 调用API时发生错误: {str(e)}"

# ==============================================================================
# --- (修改后) 工作线程函数 ---
# ==============================================================================

def batch_grading_worker(input_folder, user_rubric, client, progress_queue):
    """这个函数在后台线程中运行，负责所有耗时操作"""
    try:
        files_to_grade = [f for f in os.listdir(input_folder) if f.lower().endswith(('.docx', '.pdf'))]
        if not files_to_grade:
            # 使用约定的 "FINISH:" 前缀来表示任务结束
            progress_queue.put("FINISH:在指定文件夹中没有找到任何.docx或.pdf文件。")
            return

        total_files = len(files_to_grade)
        success_count, error_count = 0, 0

        # 创建输出文件夹
        # os.path.dirname(input_folder) or '.' 确保即使选择根目录也能正常工作
        output_folder_parent = os.path.dirname(input_folder) or '.'
        output_folder = os.path.join(output_folder_parent, OUTPUT_FOLDER_NAME)
        os.makedirs(output_folder, exist_ok=True)

        for i, filename in enumerate(files_to_grade):
            # 将当前进度放入队列
            progress_queue.put(f"正在处理: {i + 1}/{total_files} - {filename}")
            
            file_path = os.path.join(input_folder, filename)
            report_content = extract_text_from_file(file_path)

            if report_content.startswith("Error:") or not report_content.strip():
                progress_queue.put(f"跳过: {filename} (文件读取失败或内容为空)")
                error_count += 1
                continue
            
            ai_raw_feedback = grade_lab_report(report_content, client, MODEL_NAME, user_rubric)

            if ai_raw_feedback.startswith("Error:"):
                progress_queue.put(f"失败: {filename} ({ai_raw_feedback})")
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
                progress_queue.put(f"保存失败: {filename} ({e})")
                error_count += 1
            
            # 可以在这里稍微暂停，避免API调用过于频繁（如果需要）
            # time.sleep(1) 

        final_message = f"所有报告批阅完成！\n\n输出文件夹: {output_folder}\n成功: {success_count} 份\n失败: {error_count} 份"
        progress_queue.put(f"FINISH:{final_message}")

    except Exception as e:
        progress_queue.put(f"FINISH:处理过程中发生意外错误: {e}")

# ==============================================================================
# --- (修改后) 主函数 ---
# ==============================================================================

def main():
    """主函数，负责GUI和启动工作线程。"""
    root = tk.Tk()
    root.withdraw()
    print("--- 实验报告评析宝 启动 ---")

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
        client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url="https://api.deepseek.com")
        print("API客户端初始化成功。")
    except Exception as e:
        messagebox.showerror("API客户端错误", f"初始化API客户端失败: {e}", parent=root)
        return

    # === 创建进度窗口和通信队列 ===
    loading_window = tk.Toplevel(root)
    loading_window.title("正在处理中...")
    loading_window.geometry("450x200")
    loading_window.resizable(False, False)
    loading_window.attributes('-topmost', True)
    loading_window.protocol("WM_DELETE_WINDOW", lambda: None) # 禁用关闭按钮

    tk.Label(loading_window, text="AI正在全力工作中，请稍候...", font=("Arial", 16, "bold")).pack(pady=(20, 10))
    tk.Label(loading_window, text="我正在快马加鞭地为您评阅报告，马上就好！", font=("Arial", 11), fg="gray").pack(pady=5)
    
    progress_text = tk.StringVar(value="准备开始处理...")
    tk.Label(loading_window, textvariable=progress_text, font=("Arial", 10), wraplength=420).pack(pady=10)

    progress_queue = queue.Queue()

    # === 启动后台工作线程 ===
    # 将需要传递给工作函数的参数打包成元组
    worker_args = (input_folder, user_rubric, client, progress_queue)
    worker_thread = threading.Thread(target=batch_grading_worker, args=worker_args, daemon=True)
    worker_thread.start()

    # === 新增：定义一个函数来检查队列并更新GUI ===
    def check_progress_queue():
        try:
            # get_nowait() 不会阻塞，如果队列为空会抛出异常
            message = progress_queue.get_nowait()
            
            if message.startswith("FINISH:"):
                # 任务完成
                final_message = message.split(":", 1)[1]
                loading_window.destroy()
                messagebox.showinfo("任务完成", final_message, parent=root)
                # 任务完成后可以关闭主程序
                root.destroy()
            else:
                # 更新进度文本
                progress_text.set(message)
                # 100毫秒后再次检查队列
                root.after(100, check_progress_queue)
        except queue.Empty:
            # 队列为空，继续等待
            root.after(100, check_progress_queue)

    # === 开始轮询队列 ===
    root.after(100, check_progress_queue)

    # 启动tkinter的主事件循环（对于没有主窗口的程序，这种模式下需要它）
    root.mainloop()


if __name__ == "__main__":
    main()