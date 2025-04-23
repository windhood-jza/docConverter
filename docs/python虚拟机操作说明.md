# Python 虚拟环境操作说明 (docConverter 项目)

本文档说明如何在 `docConverter` 项目中使用 Python 虚拟环境。

## 1. 什么是 Python 虚拟环境？为什么使用它？

Python 虚拟环境是为特定项目创建一个独立的 Python 运行环境的方法。

**主要好处：**

*   **依赖隔离:** 每个项目可以拥有自己独立的库及其版本，避免不同项目间的库版本冲突。例如，`docConverter` 需要 `python-docx`，这个库只会被安装到它的虚拟环境中，不影响电脑上其他 Python 项目。
*   **清晰的依赖管理:** 结合 `requirements.txt` 文件，可以精确地管理和复现项目所需的依赖库。
*   **保持全局环境清洁:** 避免将大量项目特定的库安装到系统全局 Python 中。

简单来说，虚拟环境为 `docConverter` 项目提供了一个干净、隔离的"工作空间"。

## 2. 创建虚拟环境 (只需执行一次)

1.  打开你的终端 (例如 Windows PowerShell)。
2.  使用 `cd` 命令导航到 `docConverter` 项目的根目录 (即包含 `src`, `docs`, `requirements.txt` 的目录)。
    ```powershell
    cd path\to\docConverter 
    ```
3.  运行以下命令来创建名为 `venv` 的虚拟环境目录：
    ```powershell
    python -m venv venv 
    ```
    *   (如果 `python` 命令无效或指向 Python 2, 尝试 `py -3 -m venv venv` 或 `python3 -m venv venv`)
4.  你会在 `docConverter` 目录下看到一个新创建的 `venv` 文件夹。

## 3. 激活虚拟环境 (每次开始工作时)

在你开始为 `docConverter` 项目工作（例如运行脚本、安装库）之前，需要先激活虚拟环境。

1.  确保你的终端在 `docConverter` 项目根目录下。
2.  运行以下命令 (Windows PowerShell): 
    ```powershell
    .\venv\Scripts\Activate.ps1
    ```
3.  **执行策略问题:** 如果遇到关于执行策略 (Execution Policy) 的错误，提示脚本被禁用，请尝试运行以下命令**临时允许**当前会话执行脚本，然后再重新运行激活命令：
    ```powershell
    Set-ExecutionPolicy RemoteSigned -Scope Process
    ```
4.  成功激活后，你的终端提示符前面会出现 `(venv)` 字样，例如：
    ```powershell
    (venv) PS D:\DEV\docConverter>
    ```
    这表示你已进入虚拟环境。

## 4. 安装项目依赖 (激活环境后)

激活虚拟环境后，你需要安装项目所需的库。

1.  确保你处于激活的虚拟环境中 (提示符前有 `(venv)` )。
2.  运行以下命令，它会读取 `requirements.txt` 文件并安装其中列出的库 (`python-docx`):
    ```powershell
    pip install -r requirements.txt
    ```

## 5. 使用虚拟环境

一旦虚拟环境被激活，你在此终端中运行的所有 `python` 和 `pip` 命令都将使用虚拟环境中的解释器和库。例如，将来运行程序时：

```powershell
(venv) PS D:\DEV\docConverter> python src/main.py 
```

## 6. 退出虚拟环境 (完成工作时)

当你完成工作，想退出虚拟环境时，只需在终端中运行：

```powershell
deactivate
```

提示符前面的 `(venv)` 会消失。 