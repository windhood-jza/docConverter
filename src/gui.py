#!/usr/bin/env python
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
from . import converter  # 相对导入
from . import utils  # 相对导入
import logging
import webbrowser

# 尝试获取已配置的 logger, 如果没有则基本配置
try:
    logger = logging.getLogger("docConverterApp")
    if not logger.handlers:
        logging.basicConfig(
            level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
        )
        logger = logging.getLogger(__name__)
except:
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
    )
    logger = logging.getLogger(__name__)


class App:
    def __init__(self, master):
        self.master = master
        master.title("Word 表格转 Excel 工具")

        # 增大窗口尺寸
        window_width = 650
        window_height = 400  # 增加高度以容纳 Text 区域
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        master.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        master.resizable(True, True)  # 允许调整窗口大小

        # 样式
        style = ttk.Style()
        style.theme_use("clam")  # 或者 'alt', 'default', 'classic'

        # 变量
        self.word_path_var = tk.StringVar()
        self.excel_path_var = tk.StringVar()
        self.log_path = None
        self.output_excel_path = None

        self._create_widgets()
        # 初始状态显示在 Text 区域
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.insert(tk.END, "请选择 Word 文件和 Excel 保存路径")
        self.status_text.config(state=tk.DISABLED)

    def _create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        # 让主框架随窗口缩放
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)

        # --- 输入区域 --- (使用新的 frame)
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        input_frame.columnconfigure(1, weight=1)  # 让输入框随宽度变化

        ttk.Label(input_frame, text="Word 文件:").grid(
            row=0, column=0, sticky=tk.W, pady=3, padx=5
        )
        word_entry = ttk.Entry(input_frame, textvariable=self.word_path_var)
        word_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=3)
        word_button = ttk.Button(
            input_frame, text="选择...", command=self._select_word_file
        )
        word_button.grid(row=0, column=2, sticky=tk.E, padx=5, pady=3)

        ttk.Label(input_frame, text="Excel 保存为:").grid(
            row=1, column=0, sticky=tk.W, pady=3, padx=5
        )
        excel_entry = ttk.Entry(input_frame, textvariable=self.excel_path_var)
        excel_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=3)
        excel_button = ttk.Button(
            input_frame, text="选择...", command=self._select_excel_file
        )
        excel_button.grid(row=1, column=2, sticky=tk.E, padx=5, pady=3)

        # --- 转换按钮 --- (单独一行，居中)
        self.convert_button = ttk.Button(
            main_frame, text="开始转换", command=self._start_conversion
        )
        self.convert_button.grid(row=1, column=0, pady=10)

        # --- 状态/结果显示区域 (使用 ScrolledText) ---
        ttk.Label(main_frame, text="状态与结果:").grid(
            row=2, column=0, sticky=tk.W, padx=5, pady=(10, 0)
        )
        self.status_text = scrolledtext.ScrolledText(
            main_frame, height=8, wrap=tk.WORD, state=tk.DISABLED
        )
        self.status_text.grid(
            row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5
        )
        # 让 Text 区域随窗口缩放
        main_frame.rowconfigure(3, weight=1)
        main_frame.columnconfigure(0, weight=1)

        # --- 底部按钮区域 --- (使用新的 frame)
        bottom_button_frame = ttk.Frame(main_frame)
        bottom_button_frame.grid(row=4, column=0, pady=10)

        self.open_excel_button = ttk.Button(
            bottom_button_frame,
            text="打开 Excel",
            command=self._open_excel,
            state=tk.DISABLED,
        )
        self.open_excel_button.pack(side=tk.LEFT, padx=10)

        self.open_folder_button = ttk.Button(
            bottom_button_frame,
            text="打开所在文件夹",
            command=self._open_folder,
            state=tk.DISABLED,
        )
        self.open_folder_button.pack(side=tk.LEFT, padx=10)

        self.open_log_button = ttk.Button(
            bottom_button_frame,
            text="打开日志",
            command=self._open_log,
            state=tk.DISABLED,
        )
        self.open_log_button.pack(side=tk.LEFT, padx=10)

    def _select_word_file(self):
        file_path = filedialog.askopenfilename(
            title="选择 Word 文档",
            filetypes=(("Word 文档", "*.docx"), ("所有文件", "*.*")),
        )
        if file_path:
            self.word_path_var.set(file_path)
            logger.info(f"Word file selected: {file_path}")
            self.status_text.config(state=tk.NORMAL)
            self.status_text.delete(1.0, tk.END)
            self.status_text.insert(tk.END, "Word 文件已选择")
            self.status_text.config(state=tk.DISABLED)

    def _select_excel_file(self):
        file_path = filedialog.asksaveasfilename(
            title="选择 Excel 保存路径",
            defaultextension=".xlsx",
            filetypes=(("Excel 文件", "*.xlsx"), ("所有文件", "*.*")),
        )
        if file_path:
            self.excel_path_var.set(file_path)
            logger.info(f"Excel save path selected: {file_path}")
            self.status_text.config(state=tk.NORMAL)
            self.status_text.delete(1.0, tk.END)
            self.status_text.insert(tk.END, "Excel 保存路径已选择")
            self.status_text.config(state=tk.DISABLED)

    def _start_conversion(self):
        word_path = self.word_path_var.get()
        excel_path = self.excel_path_var.get()

        if not word_path or not os.path.exists(word_path):
            # 直接显示错误到状态区，不弹窗
            self.status_text.config(state=tk.NORMAL)
            self.status_text.delete(1.0, tk.END)
            self.status_text.insert(tk.END, "错误: 请选择有效的 Word 文件路径！")
            self.status_text.config(state=tk.DISABLED)
            logger.error("Start conversion attempt failed: Invalid Word path.")
            return
        if not excel_path:
            self.status_text.config(state=tk.NORMAL)
            self.status_text.delete(1.0, tk.END)
            self.status_text.insert(tk.END, "错误: 请选择 Excel 文件保存路径！")
            self.status_text.config(state=tk.DISABLED)
            logger.error("Start conversion attempt failed: Invalid Excel path.")
            return

        # 禁用按钮
        self.convert_button.config(state=tk.DISABLED)
        self.open_excel_button.config(state=tk.DISABLED)
        self.open_folder_button.config(state=tk.DISABLED)
        self.open_log_button.config(state=tk.DISABLED)
        self.log_path = None

        # 更新状态区
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.insert(tk.END, "正在处理中，请稍候...")
        self.status_text.config(state=tk.DISABLED)
        logger.info("Starting conversion process in a new thread...")

        thread = threading.Thread(
            target=self._run_conversion_thread,
            args=(word_path, excel_path),
            daemon=True,
        )
        thread.start()

    def _run_conversion_thread(self, word_path, excel_path):
        """在后台线程中执行转换逻辑。"""
        result = None
        converter_instance = None  # 初始化
        try:
            converter_instance = converter.DocConverter(word_path, excel_path)
            result = converter_instance.convert()
            logger.info("Conversion thread finished.")
        except Exception as e:
            error_message = f"转换过程中发生意外错误: {e}"
            logger.error(error_message, exc_info=True)
            result = {
                "status": "error",
                "message": error_message,
                "success": 0,
                "errors": 0,
                "skipped_empty": 0,
                "excel_path": excel_path,
                # 尝试获取日志路径，即使出错也可能需要
                "log_path": (
                    getattr(converter_instance, "log_path", None)
                    if converter_instance
                    else None
                ),
            }
        if self.master.winfo_exists():
            self.master.after(0, lambda: self._update_gui_post_conversion(result))

    def _update_gui_post_conversion(self, result):
        """在主线程中根据转换结果更新 GUI。"""
        if not self.master.winfo_exists():
            logger.warning("GUI window closed before conversion result could be displayed.")
            return

        self.convert_button.config(state=tk.NORMAL)

        # 准备状态消息
        status_message = result.get("message", "发生未知错误")
        success_count = result.get("success", 0)
        error_count = result.get("errors", 0)
        # 获取总的跳过行数
        total_skipped = result.get("total_skipped_rows", 0)

        # 简化最终状态信息
        final_status = (
            f"状态: {status_message}\n"
            f"(成功: {success_count}, 失败: {error_count}, 跳过空行: {total_skipped})"
        )

        # 将状态消息写入 Text 区域
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.insert(tk.END, final_status)
        self.status_text.config(state=tk.DISABLED)

        # 更新 internal state
        self.log_path = result.get("log_path")
        self.output_excel_path = result.get("excel_path")

        # 更新按钮状态
        # 移除所有 messagebox 调用
        if result.get("status") == "success" or (
            result.get("status") == "warning" and success_count > 0
        ):
            # messagebox.showinfo("完成", final_status)
            self.open_excel_button.config(state=tk.NORMAL)
            self.open_folder_button.config(state=tk.NORMAL)
            logger.info(f"Conversion finished. Status: {result.get('status')}, Success: {success_count}, Errors: {error_count}, Total Skipped: {total_skipped}")
        elif result.get("status") == "warning":
            # messagebox.showwarning("警告", final_status)
            self.open_folder_button.config(state=tk.NORMAL)  # 允许打开文件夹看日志
            logger.warning(f"Conversion finished with warning. Status: {result.get('status')}, Success: {success_count}, Errors: {error_count}, Total Skipped: {total_skipped}")
        else:  # error
            # messagebox.showerror("错误", final_status)
            self.open_folder_button.config(state=tk.NORMAL)  # 允许打开文件夹看日志
            logger.error(f"Conversion failed. Status: {result.get('status')}, Success: {success_count}, Errors: {error_count}, Total Skipped: {total_skipped}, Message: {status_message}")

        # 总是尝试启用日志按钮 (如果日志文件存在)
        if self.log_path and os.path.exists(self.log_path):
            self.open_log_button.config(state=tk.NORMAL)
        else:
            self.open_log_button.config(state=tk.DISABLED)

    def _open_generic(self, path, file_type="文件"):
        """通用打开文件或文件夹的方法。"""
        if path and os.path.exists(path):
            logger.info(f"Attempting to open {file_type}: {path}")
            try:
                os.startfile(path)
            except AttributeError:
                try:
                    # 对于文件，使用 file:/// 协议；对于目录，直接打开
                    uri = (
                        f"file:///{os.path.abspath(path)}"
                        if os.path.isfile(path)
                        else os.path.abspath(path)
                    )
                    webbrowser.open(uri)
                except Exception as e:
                    messagebox.showerror(
                        "错误", f"无法打开{file_type}: {e}\n路径: {path}"
                    )  # 这里仍然用弹窗提示打开失败
            except Exception as e:
                messagebox.showerror(
                    "错误", f"打开{file_type}时出错: {e}\n路径: {path}"
                )
        else:
            messagebox.showwarning("警告", f"{file_type}路径无效或不存在。")
            logger.warning(
                f"Attempted to open {file_type}, but path is invalid or file does not exist: {path}"
            )

    def _open_excel(self):
        self._open_generic(self.output_excel_path, "Excel 文件")

    def _open_folder(self):
        folder_path = None
        if self.output_excel_path:
            folder_path = os.path.dirname(self.output_excel_path)
        # elif self.log_path: # 如果 Excel 路径无效，可以尝试用日志路径的目录
        #     folder_path = os.path.dirname(self.log_path)
        self._open_generic(folder_path, "文件夹")

    def _open_log(self):
        self._open_generic(self.log_path, "日志文件")

    def run(self):
        self.master.mainloop()


# 主程序块 (如果需要直接运行 GUI 测试)
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    app.run()
