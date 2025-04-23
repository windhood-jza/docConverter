import logging
import os

DEFAULT_LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
DEFAULT_LOG_LEVEL = logging.INFO


def setup_logging(log_file_path):
    """配置日志记录器

    Args:
        log_file_path (str): 日志文件的完整路径。
    """
    log_dir = os.path.dirname(log_file_path)
    if log_dir and not os.path.exists(log_dir):
        os.makedirs(log_dir)  # 确保日志目录存在

    # 获取或创建 logger
    # 使用特定的名字，避免直接修改 root logger，除非确实需要
    logger = logging.getLogger("docConverterApp")
    logger.setLevel(DEFAULT_LOG_LEVEL)

    # 防止重复添加 handler (如果此函数可能被多次调用)
    if not logger.handlers:
        # 创建 FileHandler
        # 使用追加模式 'a'，编码为 utf-8
        file_handler = logging.FileHandler(log_file_path, mode="a", encoding="utf-8")
        file_handler.setLevel(DEFAULT_LOG_LEVEL)

        # 创建 Formatter
        formatter = logging.Formatter(DEFAULT_LOG_FORMAT)
        file_handler.setFormatter(formatter)

        # 将 Handler 添加到 Logger
        logger.addHandler(file_handler)

        # (可选) 如果也想在控制台看到日志输出，可以添加 StreamHandler
        # console_handler = logging.StreamHandler()
        # console_handler.setLevel(logging.DEBUG) # 控制台可以显示更详细的信息
        # console_handler.setFormatter(formatter)
        # logger.addHandler(console_handler)

    return logger


if __name__ == "__main__":
    # 测试日志配置
    test_log_file = "test_app.log"
    logger = setup_logging(test_log_file)

    logger.info("这是 INFO 级别的测试日志。")
    logger.warning("这是 WARNING 级别的测试日志。")
    logger.error("这是 ERROR 级别的测试日志。")
    print(f"测试日志已写入 {test_log_file}")

    # 清理测试文件
    # import time
    # time.sleep(1) # 确保文件写入完成
    # if os.path.exists(test_log_file):
    #     os.remove(test_log_file)
    #     print(f"测试日志文件 {test_log_file} 已删除。")
