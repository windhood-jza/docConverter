import datetime
import os
import platform
import subprocess
import logging

# 配置日志记录器 (与 logger_config.py 保持一致)
# 如果此模块作为独立脚本运行, 则基本配置; 否则期望主程序已配置
try:
    # 尝试获取已配置的 logger
    logger = logging.getLogger("docConverterApp")
    # 如果没有 handler, 说明可能未被主程序配置, 做基本配置
    if not logger.handlers:
        logging.basicConfig(
            level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
        )
        logger = logging.getLogger(__name__)  # 使用当前模块名，避免干扰主 logger
except:
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
    )
    logger = logging.getLogger(__name__)


# 支持的日期格式列表
SUPPORTED_DATE_FORMATS = [
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y.%m.%d",
    "%y-%m-%d",
    "%y/%m/%d",
    "%y.%m.%d",
    "%Y年%m月%d日",
    "%Y年%m月",  # 支持年月格式
    "%m/%d/%Y",
    "%m-%d-%Y",
    "%m.%d.%Y",
    "%m/%d/%y",
    "%m-%d-%y",
    "%m.%d.%y",
    # 添加更多可能的格式...
]


def parse_date(date_str):
    """
    尝试使用多种格式解析日期字符串。
    支持仅年月格式，解析为该月第一天。
    返回 date 对象或 None。
    """
    if not isinstance(date_str, str):
        logger.warning(f"parse_date received non-string input: {date_str}")
        return None
    date_str = date_str.strip()
    if not date_str:
        return None

    for fmt in SUPPORTED_DATE_FORMATS:
        try:
            # 尝试解析完整日期
            dt = datetime.datetime.strptime(date_str, fmt)
            return dt.date()
        except ValueError:
            continue  # 格式不匹配，尝试下一个

    # 如果所有格式都失败，返回 None
    logger.warning(
        f"Could not parse date string: '{date_str}' with any supported format."
    )
    return None


def normalize_header(cell_text):
    """
    标准化表头文本：去除首尾空格并转为小写。
    """
    if not isinstance(cell_text, str):
        # 对于非字符串（可能是数字或其他类型），尝试转换为字符串
        try:
            cell_text = str(cell_text)
        except:
            logger.warning(
                f"normalize_header received non-string input that couldn't be converted: {type(cell_text)}"
            )
            return ""  # 或者返回 None，根据需要处理
    return cell_text.strip().lower()


def open_file_or_folder(path):
    """
    根据操作系统打开文件或文件夹。
    """
    if not path or not os.path.exists(path):
        logger.error(f"Path does not exist or is None: {path}")
        return False
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(path)  # 在 Windows 上更可靠
        elif system == "Darwin":  # macOS
            subprocess.Popen(["open", path])
        else:  # Linux and other Unix-like
            subprocess.Popen(["xdg-open", path])
        logger.info(f"Attempted to open: {path}")
        return True
    except Exception as e:
        logger.error(f"Failed to open path '{path}': {e}")
        return False


# 测试块
if __name__ == "__main__":
    # 配置基本日志以便测试输出
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
    )

    print("--- Testing parse_date ---")
    test_dates = [
        "2023-10-26",
        "2023/10/27",
        "2023.10.28",
        "23-11-01",
        "23/11/02",
        "23.11.03",
        "2024年5月20日",
        "2024年5月",
        "12/31/2023",
        "12-31-23",
        "Invalid Date",
        "",
        None,
        123,
    ]
    for d in test_dates:
        parsed = parse_date(d)
        print(f"Original: '{d}', Parsed: {parsed}")

    print("\n--- Testing normalize_header ---")
    test_headers = [
        "  Header 1 ",
        "HEADER 2",
        " header3 ",
        " 带空格的标题 ",
        "",
        None,
        12345,
    ]
    for h in test_headers:
        normalized = normalize_header(h)
        print(f"Original: '{h}', Normalized: '{normalized}'")

    print("\n--- Testing open_file_or_folder ---")
    # 创建临时文件和目录用于测试
    temp_dir = "temp_test_open"
    temp_file = os.path.join(temp_dir, "test_file.txt")
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    with open(temp_file, "w") as f:
        f.write("Test content.")

    print(f"Attempting to open existing file: {temp_file}")
    open_file_or_folder(temp_file)
    print(f"Attempting to open existing folder: {temp_dir}")
    open_file_or_folder(temp_dir)
    print(f"Attempting to open non-existing path: non_existent_file.txt")
    open_file_or_folder("non_existent_file.txt")

    # 清理临时文件和目录 (可选)
    # import time
    # time.sleep(2) # 等待文件浏览器打开
    # os.remove(temp_file)
    # os.rmdir(temp_dir)
    # print("Cleaned up temp file and directory.")
