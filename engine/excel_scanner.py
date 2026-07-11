"""
Excel 扫描器 —— 递归扫描文件夹中的 .xlsx 文件。
"""
import os
from typing import Optional


def scan_xlsx_files(root_dir: str) -> list[str]:
    """递归扫描目录下所有 .xlsx 文件，返回绝对路径列表。"""
    if not os.path.isdir(root_dir):
        return []

    result = []
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for fname in filenames:
            if fname.lower().endswith('.xlsx') and not fname.startswith('~$'):
                full_path = os.path.join(dirpath, fname)
                result.append(full_path)

    # 按文件名排序
    result.sort(key=lambda p: os.path.basename(p).lower())
    return result


def filter_by_name(files: list[str], keyword: str) -> list[str]:
    """按文件名过滤（大小写不敏感）。"""
    if not keyword.strip():
        return files
    kw = keyword.strip().lower()
    return [f for f in files if kw in os.path.basename(f).lower()]
