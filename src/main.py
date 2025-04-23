import tkinter as tk

# 确保 gui 在 src 目录下，并且 Python 可以找到它
# 如果从 docConverter 根目录运行，需要确保 src 在 Python 路径中
# 或者使用相对导入 (如果作为包运行)
# 为了简单起见，假设从根目录运行或者 src 在 PYTHONPATH
from src.gui import App

if __name__ == "__main__":
    # 创建 Tkinter 根窗口
    root = tk.Tk()
    # 创建应用程序实例
    app = App(root)
    # 运行应用程序主循环
    app.run()
