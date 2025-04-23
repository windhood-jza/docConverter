import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from . import converter  # 相对导入
from . import utils  # 相对导入
import logging

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
        master.title("Word 表格转 CSV 工具")

        # 窗口大小和居中
        window_width = 500
        window_height = 250
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        master.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        master.resizable(False, False)  # 禁止调整窗口大小

        # 样式
        style = ttk.Style()
        style.theme_use("clam")  # 或者 'alt', 'default', 'classic'

        # 变量
        self.word_path_var = tk.StringVar()
        self.csv_path_var = tk.StringVar()
        self.status_var = tk.StringVar()
        self.log_path = None
        self.output_csv_path = None  # 存储实际的输出 CSV 路径

        self._create_widgets()
        self.status_var.set("请选择 Word 文件和 CSV 保存路径")

    def _create_widgets(self):
        frame = ttk.Frame(self.master, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)

        # Word 文件选择
        ttk.Label(frame, text="Word 文件:").grid(row=0, column=0, sticky=tk.W, pady=2)
        word_entry = ttk.Entry(frame, textvariable=self.word_path_var, width=40)
        word_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2)
        word_button = ttk.Button(frame, text="选择...", command=self._select_word_file)
        word_button.grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)

        # CSV 文件保存路径
        ttk.Label(frame, text="CSV 保存为:").grid(row=1, column=0, sticky=tk.W, pady=2)
        csv_entry = ttk.Entry(frame, textvariable=self.csv_path_var, width=40)
        csv_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2)
        csv_button = ttk.Button(frame, text="选择...", command=self._select_csv_file)
        csv_button.grid(row=1, column=2, sticky=tk.W, padx=5, pady=2)

        # 转换按钮
        self.convert_button = ttk.Button(
            frame, text="开始转换", command=self._start_conversion
        )
        self.convert_button.grid(row=2, column=1, pady=10)

        # 状态标签
        status_label = ttk.Label(frame, textvariable=self.status_var, wraplength=480)
        status_label.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)

        # 打开文件/文件夹按钮 (初始禁用/隐藏)
        self.open_csv_button = ttk.Button(
            frame, text="打开 CSV", command=self._open_csv, state=tk.DISABLED
        )
        self.open_csv_button.grid(row=4, column=0, sticky=tk.W, pady=5, padx=5)

        self.open_folder_button = ttk.Button(
            frame, text="打开所在文件夹", command=self._open_folder, state=tk.DISABLED
        )
        self.open_folder_button.grid(row=4, column=1, pady=5, padx=5)

        # 打开日志按钮初始不创建，转换后根据是否有 log_path 再创建和显示
        self.open_log_button = None

        frame.columnconfigure(1, weight=1)

    def _select_word_file(self):
        file_path = filedialog.askopenfilename(
            title="选择 Word 文档",
            filetypes=(("Word 文档", "*.docx"), ("所有文件", "*.*")),
        )
        if file_path:
            self.word_path_var.set(file_path)
            logger.info(f"Word file selected: {file_path}")
            self.status_var.set("Word 文件已选择")

    def _select_csv_file(self):
        file_path = filedialog.asksaveasfilename(
            title="选择 CSV 保存路径",
            defaultextension=".csv",
            filetypes=(("CSV 文件", "*.csv"), ("所有文件", "*.*")),
        )
        if file_path:
            self.csv_path_var.set(file_path)
            logger.info(f"CSV save path selected: {file_path}")
            self.status_var.set("CSV 保存路径已选择")

    def _start_conversion(self):
        word_path = self.word_path_var.get()
        csv_path = self.csv_path_var.get()

        if not word_path or not os.path.exists(word_path):
            messagebox.showerror("错误", "请选择有效的 Word 文件路径！")
            logger.error("Start conversion attempt failed: Invalid Word path.")
            return
        if not csv_path:
            messagebox.showerror("错误", "请选择 CSV 文件保存路径！")
            logger.error("Start conversion attempt failed: Invalid CSV path.")
            return

        self.convert_button.config(state=tk.DISABLED)
        self.open_csv_button.config(state=tk.DISABLED)
        self.open_folder_button.config(state=tk.DISABLED)
        if self.open_log_button:
            self.open_log_button.grid_remove()  # 隐藏日志按钮

        self.status_var.set("正在处理中，请稍候...")
        logger.info("Starting conversion process in a new thread...")

        # 使用线程避免 GUI 阻塞
        thread = threading.Thread(
            target=self._run_conversion_thread,
            args=(word_path, csv_path),
            daemon=True,  # 主窗口关闭时线程也退出
        )
        thread.start()

    def _run_conversion_thread(self, word_path, csv_path):
        """在后台线程中执行转换逻辑。"""
        result = None
        try:
            converter_instance = converter.DocConverter(word_path, csv_path)
            result = converter_instance.convert()
            logger.info("Conversion thread finished.")
        except Exception as e:
            error_message = f"转换过程中发生意外错误: {e}"
            logger.error(error_message, exc_info=True)
            # 创建一个错误结果字典以便在 GUI 中显示
            result = {
                "status": "error",
                "message": error_message,
                "success": 0,
                "errors": 0,
                "skipped_empty": 0,
                "csv_path": csv_path,
                "log_path": getattr(
                    converter_instance, "log_path", None
                ),  # 尝试获取日志路径
            }

        # 使用 after 将结果传递回主线程更新 GUI
        if self.master.winfo_exists():  # 检查主窗口是否还存在
            self.master.after(0, lambda: self._update_gui_post_conversion(result))

    def _update_gui_post_conversion(self, result):
        """在主线程中根据转换结果更新 GUI。"""
        if not self.master.winfo_exists():  # 再次检查，以防窗口在等待时关闭
            logger.warning(
                "GUI window closed before conversion result could be displayed."
            )
            return

        self.convert_button.config(state=tk.NORMAL)  # 恢复转换按钮
        self.status_var.set(result.get("message", "发生未知错误"))
        self.log_path = result.get("log_path")
        self.output_csv_path = result.get("csv_path")  # 更新实际输出路径

        if result.get("status") == "success" or (
            result.get("status") == "warning" and result.get("success", 0) > 0
        ):
            messagebox.showinfo("完成", result.get("message", "转换完成"))
            self.open_csv_button.config(state=tk.NORMAL)
            self.open_folder_button.config(state=tk.NORMAL)
            logger.info("Conversion successful or partially successful.")
        elif result.get("status") == "warning":  # 成功数为0的警告
            messagebox.showwarning(
                "警告", result.get("message", "转换完成，但没有数据写入")
            )
            self.open_folder_button.config(state=tk.NORMAL)  # 允许打开文件夹看日志
            logger.warning(f"Conversion resulted in a warning: {result.get('message')}")
        else:  # error
            messagebox.showerror(
                "错误", result.get("message", "转换失败，请检查日志文件获取详细信息。")
            )
            self.open_folder_button.config(state=tk.NORMAL)  # 允许打开文件夹看日志
            logger.error(f"Conversion failed: {result.get('message')}")

        # 处理日志按钮
        if self.log_path and os.path.exists(self.log_path):
            if not self.open_log_button:
                # 如果按钮不存在则创建
                self.open_log_button = ttk.Button(
                    self.master.winfo_children()[0],
                    text="打开日志",
                    command=self._open_log,
                )
                self.open_log_button.grid(row=4, column=2, sticky=tk.E, pady=5, padx=5)
            else:
                # 如果已存在则重新显示并启用
                self.open_log_button.config(state=tk.NORMAL)
                self.open_log_button.grid()  # 确保它可见
        elif self.open_log_button:  # 如果日志路径无效或不存在，但按钮存在，则隐藏
            self.open_log_button.grid_remove()

    def _open_csv(self):
        if self.output_csv_path and os.path.exists(self.output_csv_path):
            logger.info(f"Attempting to open CSV file: {self.output_csv_path}")
            if not utils.open_file_or_folder(self.output_csv_path):
                messagebox.showerror(
                    "错误", f"无法打开 CSV 文件：\n{self.output_csv_path}"
                )
        else:
            messagebox.showwarning("警告", "CSV 文件路径无效或文件不存在。")
            logger.warning(
                "Attempted to open CSV, but path is invalid or file does not exist."
            )

    def _open_folder(self):
        folder_path = None
        if self.output_csv_path:
            folder_path = os.path.dirname(self.output_csv_path)
        elif self.log_path:
            folder_path = os.path.dirname(self.log_path)

        if folder_path and os.path.exists(folder_path):
            logger.info(f"Attempting to open folder: {folder_path}")
            if not utils.open_file_or_folder(folder_path):
                messagebox.showerror("错误", f"无法打开文件夹：\n{folder_path}")
        else:
            messagebox.showwarning("警告", "无法确定要打开的文件夹路径。")
            logger.warning(
                "Attempted to open folder, but could not determine a valid path."
            )

    def _open_log(self):
        if self.log_path and os.path.exists(self.log_path):
            logger.info(f"Attempting to open log file: {self.log_path}")
            if not utils.open_file_or_folder(self.log_path):
                messagebox.showerror("错误", f"无法打开日志文件：\n{self.log_path}")
        else:
            messagebox.showwarning("警告", "日志文件路径无效或文件不存在。")
            logger.warning(
                "Attempted to open log file, but path is invalid or file does not exist."
            )

    def run(self):
        self.master.mainloop()


# 主程序块 (如果需要直接运行 GUI 测试)
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    app.run()
