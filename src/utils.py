#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工具函数库

作者: gba-mep
版本: v2.1
日期: 2026-06-02
"""

import os
import re
from pathlib import Path
from typing import List, Optional
from datetime import datetime


def clean_sheet_name(name: str) -> str:
    """
    清理工作表名称（移除非法字符）
    
    Args:
        name: 原始名称
    
    Returns:
        清理后的名称
    """
    illegal = ['\\', '/', '*', '?', ':', '[', ']']
    for ch in illegal:
        name = name.replace(ch, '_')
    return name[:31] if len(name) > 31 else name


def clean_filename(name: str) -> str:
    """
    清理文件名（移除非法字符）
    
    Args:
        name: 原始名称
    
    Returns:
        清理后的名称
    """
    illegal = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for ch in illegal:
        name = name.replace(ch, '_')
    return name[:200] if len(name) > 200 else name


def get_relative_path(base_path: str, full_path: str) -> str:
    """
    计算相对路径
    
    Args:
        base_path: 基础路径
        full_path: 完整路径
    
    Returns:
        相对路径
    """
    base = Path(base_path).resolve()
    full = Path(full_path).resolve()
    
    try:
        rel = full.relative_to(base)
        return f"./{rel}"
    except ValueError:
        return str(full)


def col_letter(col_num: int) -> str:
    """
    列号转字母（1 = A, 27 = AA）
    
    Args:
        col_num: 列号（1-based）
    
    Returns:
        列字母
    """
    result = ""
    while col_num > 0:
        col_num -= 1
        result = chr(65 + (col_num % 26)) + result
        col_num //= 26
    return result


def col_number(col_letter: str) -> int:
    """
    列字母转号（A = 1, AA = 27）
    
    Args:
        col_letter: 列字母
    
    Returns:
        列号（1-based）
    """
    result = 0
    for ch in col_letter.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result


def find_files_recursively(
    folder: str,
    pattern: str = "*",
    include_subfolders: bool = True
) -> List[Path]:
    """
    递归查找文件
    
    Args:
        folder: 搜索文件夹
        pattern: 文件模式（如 "*.pdf"）
        include_subfolders: 是否包含子文件夹
    
    Returns:
        文件路径列表
    """
    folder_path = Path(folder)
    if not folder_path.exists():
        return []
    
    if include_subfolders:
        return list(folder_path.glob(f"**/{pattern}"))
    else:
        return list(folder_path.glob(pattern))


def get_latest_file(files: List[Path]) -> Optional[Path]:
    """
    获取最新修改的文件
    
    Args:
        files: 文件列表
    
    Returns:
        最新文件路径，空列表返回 None
    """
    if not files:
        return None
    
    latest = files[0]
    latest_time = latest.stat().st_mtime
    
    for file_path in files[1:]:
        mtime = file_path.stat().st_mtime
        if mtime > latest_time:
            latest = file_path
            latest_time = mtime
    
    return latest


class Logger:
    """
    日志记录器
    
    使用示例：
        logger = Logger("./logs", "fill_log")
        logger.log("开始处理")
        logger.close()
    """
    
    def __init__(self, folder: str, prefix: str):
        """
        初始化日志记录器
        
        Args:
            folder: 日志文件夹
            prefix: 日志文件前缀
        """
        log_folder = Path(folder)
        log_folder.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_path = log_folder / f"{prefix}_{timestamp}.txt"
        self.log_file = open(self.log_path, 'w', encoding='utf-8')
        
        self.log(f"开始时间: {datetime.now()}")
    
    def log(self, message: str):
        """
        写入日志
        
        Args:
            message: 日志消息
        """
        if self.log_file:
            self.log_file.write(f"{message}\n")
            self.log_file.flush()
    
    def close(self):
        """关闭日志"""
        if self.log_file:
            self.log(f"结束时间: {datetime.now()}")
            self.log_file.close()
            self.log_file = None


def setup_chinese_font():
    """
    设置中文字体（用于终端输出）
    
    确保 Python 能正确输出中文到 Windows 终端
    """
    import sys
    import io

    if sys.platform == 'win32' and hasattr(sys.stdout, 'buffer') and not sys.stdout.closed:
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer,
            encoding='utf-8',
            errors='replace'
        )
    if sys.platform == 'win32' and hasattr(sys.stderr, 'buffer') and not sys.stderr.closed:
        sys.stderr = io.TextIOWrapper(
            sys.stderr.buffer,
            encoding='utf-8',
            errors='replace'
        )
