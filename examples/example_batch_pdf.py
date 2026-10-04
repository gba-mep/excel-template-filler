#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
示例：批量生成 PDF

展示如何批量填充模板并导出为 PDF 文件
"""

import sys
from pathlib import Path

# 添加 src 目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from template_filler import TemplateFiller


def main():
    """批量 PDF 示例"""
    
    # 示例文件路径（相对于 examples/ 目录）
    data_source = str(Path(__file__).parent / "data" / "sample_data.xlsx")
    template = str(Path(__file__).parent / "templates" / "sample_template.xlsx")
    output_dir = str(Path(__file__).parent / "output" / "filled_forms")
    
    # 检查文件
    if not Path(data_source).exists():
        print(f"[ERROR] 数据源不存在: {data_source}")
        return
    
    if not Path(template).exists():
        print(f"[ERROR] 模板不存在: {template}")
        return
    
    # 初始化
    filler = TemplateFiller(data_source, template)
    
    # 扫描占位符
    placeholders = filler.scan_placeholders()
    print(f"发现占位符: {placeholders}")
    
    # 验证数据
    validation = filler.validate_data()
    if not validation['valid']:
        print(f"[WARNING] 数据验证失败: {validation['missing_fields']}")
        print("[INFO] 将继续处理，但部分占位符可能无法替换")
    
    # 批量导出 PDF
    print(f"\n=== 批量导出 PDF ===")
    
    try:
        output_files = filler.fill_and_export(
            start_row=2,
            end_row=11,  # 生成 10 份
            export_format="pdf",
            output_dir=output_dir,
            empty_handling="clear",
            naming_mode="auto"  # 使用 A 列值作为文件名
        )
        
        print(f"\n已生成 {len(output_files)} 个 PDF 文件:")
        for file_path in output_files:
            print(f"  - {file_path}")
        
    except Exception as e:
        print(f"[ERROR] 导出失败: {e}")
        print("[INFO] PDF 导出需要安装 pywin32 和 Microsoft Excel")
        print("[INFO] 或使用 export_format='excel'")
    
    # 关闭
    filler.close()


if __name__ == "__main__":
    main()
