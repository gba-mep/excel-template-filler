#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
示例：自动链接文件

展示如何根据单元格内容自动匹配文件并创建超链接
"""

import sys
from pathlib import Path

# 添加 src 目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from auto_linker import AutoLinker


def main():
    """自动链接示例"""
    
    # 示例文件路径（相对于 examples/ 目录）
    workbook = str(Path(__file__).parent / "data" / "sample_list.xlsx")
    search_root = str(Path(__file__).parent / "files")
    
    # 检查文件
    if not Path(workbook).exists():
        print(f"[ERROR] 工作簿不存在: {workbook}")
        return
    
    if not Path(search_root).exists():
        print(f"[ERROR] 搜索目录不存在: {search_root}")
        return
    
    # 初始化
    linker = AutoLinker(
        workbook=workbook,
        target_column=1,  # A 列
        search_root=search_root
    )
    
    # 执行自动链接
    print("=== 自动链接文件 ===")
    print(f"工作簿: {workbook}")
    print(f"搜索目录: {search_root}")
    print(f"目标列: A 列")
    print()
    
    result = linker.convert_to_hyperlinks(
        extensions=[".pdf", ".docx", ".xlsx"],  # 支持的扩展名
        include_subfolders=True,  # 包含子文件夹
        use_wildcard=True  # 使用通配符匹配
    )
    
    print(f"\n=== 结果 ===")
    print(f"成功: {result['success']}")
    print(f"跳过: {result['skipped']}")
    
    # 关闭
    linker.close()
    print("\n[OK] 完成")


if __name__ == "__main__":
    main()
