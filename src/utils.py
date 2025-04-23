#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re
from datetime import datetime


def normalize_header(header_text):
    """标准化表头文本：去除所有空格，并将全角标点替换为半角。"""
    if header_text is None:
        return ""
    text = str(header_text).strip()
    # 替换常见全角符号为对应半角
    text = (
        text.replace("（", "(")
        .replace("）", ")")
        .replace("：", ":")
        .replace("，", ",")
        .replace("。", ".")
    )
    # 移除所有空白字符 (包括全角空格 \u3000 和其他如换行符、制表符等)
    text = re.sub(r"\s+", "", text)
    return text  # strip() 不再需要，因为所有空格已被移除


def parse_date(date_str):
    """尝试解析多种常见格式的日期字符串。"""
    if not date_str:
        return None
    # 常见的日期格式列表
    formats = [
        "%Y-%m-%d",  # 2023-10-26
        "%Y/%m/%d",  # 2023/10/26
        "%Y.%m.%d",  # 2023.10.26
        "%y-%m-%d",  # 23-10-26
        "%y/%m/%d",  # 23/10/26
        "%y.%m.%d",  # 23.10.26
        # 可以根据需要添加更多格式
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue  # 尝试下一种格式
    # 如果所有格式都失败
    return None
