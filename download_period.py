"""下载文件的业务月份目录规则。"""

import os
from datetime import date


_MONTH_NAMES = (
    "一月", "二月", "三月", "四月", "五月", "六月",
    "七月", "八月", "九月", "十月", "十一月", "十二月",
)


def previous_month_data_folder_name(reference_date=None):
    """返回上一个自然月的目录名，例如 ``2026年六月数据``。"""
    current = reference_date or date.today()
    year = current.year
    month = current.month - 1
    if month == 0:
        year -= 1
        month = 12
    return f"{year}年{_MONTH_NAMES[month - 1]}数据"


def monthly_download_dir(downloads_root, reference_date=None):
    """返回 downloads 下当前业务月份的根目录。"""
    return os.path.join(
        str(downloads_root), previous_month_data_folder_name(reference_date)
    )
