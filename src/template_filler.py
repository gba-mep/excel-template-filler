#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
统一模板填充器 - 自动选择引擎

根据模板特徵自动选择 ZIP 引擎或 openpyxl 引擎

作者: gba-mep
版本: v2.1（融合版）
日期: 2026-06-02
"""

import os
import sys
import io
import json
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple

from engines import create_engine, BaseEngine, ZIPEngine, OpenPYXLEngine


class TemplateFiller:
    """
    统一模板填充器
    
    自动检测模板类型并选择最优引擎：
    - 含图片/打印设置 → ZIP 引擎
    - 无图片 → openpyxl 引擎
    
    使用示例：
        filler = TemplateFiller("数据源.xlsx", "模板.xlsx")
        placeholders = filler.scan_placeholders()
        output_files = filler.fill_and_export(2, 11, "pdf", "./output")
    """
    
    def __init__(
        self,
        data_source: str = None,
        template: str = None,
        engine_type: str = "auto"
    ):
        """
        初始化模板填充器
        
        Args:
            data_source: 数据源 Excel 路径
            template: 模板 Excel 路径
            engine_type: 引擎类型
                "auto" - 自动检测
                "openpyxl" - 强制使用 openpyxl
                "zip" - 强制使用 ZIP
        """
        self.data_source_path = Path(data_source) if data_source else None
        self.template_path = Path(template) if template else None
        self.engine_type = engine_type
        
        # 引擎实例
        self.engine: Optional[BaseEngine] = None
        
        # 数据
        self.data_list: List[Dict[str, Any]] = []
        self.field_map: Dict[str, str] = {}
        
        # 载入模板
        if self.template_path:
            self._load_template()
    
    def _load_template(self):
        """载入模板"""
        if not self.template_path.exists():
            raise FileNotFoundError(f"模板不存在: {self.template_path}")
        
        # 创建引擎
        self.engine = create_engine(
            template_path=str(self.template_path),
            engine_type=self.engine_type
        )
        
        # 载入模板
        if not self.engine.load_template(str(self.template_path)):
            raise RuntimeError(f"载入模板失败: {self.template_path}")
        
        # 显示引擎信息
        engine_name = "ZIP" if isinstance(self.engine, ZIPEngine) else "openpyxl"
        print(f"[INFO] 使用引擎: {engine_name}")
    
    def has_images(self) -> bool:
        """检测模板是否含图片"""
        return isinstance(self.engine, ZIPEngine)
    
    def scan_placeholders(self) -> List[str]:
        """
        扫描模板中的占位符
        
        Returns:
            占位符列表
        """
        if not self.engine:
            raise RuntimeError("请先载入模板")
        
        return self.engine.scan_placeholders()
    
    def load_data(
        self,
        data_source: str = None,
        field_map: Dict[str, str] = None
    ):
        """
        载入数据源
        
        Args:
            data_source: 数据源 Excel 路径
            field_map: 字段映射 {占位符: 字段名}
        """
        if data_source:
            self.data_source_path = Path(data_source)
        
        if field_map:
            self.field_map = field_map
        
        if not self.data_source_path:
            raise RuntimeError("请提供数据源路径")
        
        if not self.data_source_path.exists():
            raise FileNotFoundError(f"数据源不存在: {self.data_source_path}")
        
        # 使用 openpyxl 读取数据
        from openpyxl import load_workbook
        
        wb = load_workbook(self.data_source_path, data_only=True)
        ws = wb.active
        
        # 读取表头
        headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]
        
        # 读取数据
        self.data_list = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if any(row):  # 跳过空行
                data_row = {}
                for i, header in enumerate(headers):
                    if header and i < len(row):
                        data_row[header] = row[i]
                self.data_list.append(data_row)
        
        wb.close()
        
        print(f"[OK] 已载入数据: {len(self.data_list)} 条记录")
    
    def validate_data(self) -> Dict[str, Any]:
        """
        验证数据完整性
        
        Returns:
            验证结果
        """
        if not self.engine:
            raise RuntimeError("请先载入模板")
        
        # 获取占位符
        placeholders = self.scan_placeholders()
        
        # 获取字段名
        field_names = set(self.field_map.values()) if self.field_map else set(placeholders)
        
        # 获取数据源字段
        if self.data_list:
            data_fields = set()
            for row in self.data_list:
                data_fields.update(row.keys())
        else:
            data_fields = set()
        
        # 检查缺失字段
        missing = field_names - data_fields
        
        return {
            "valid": len(missing) == 0,
            "total_rows": len(self.data_list),
            "missing_fields": list(missing),
            "placeholders": placeholders,
            "field_names": list(field_names),
            "data_fields": list(data_fields)
        }
    
    def fill_and_export(
        self,
        start_row: int = None,
        end_row: int = None,
        export_format: str = "excel",
        output_dir: str = "./output",
        field_map: Dict[str, str] = None,
        **kwargs
    ) -> List[str]:
        """
        填充模板并导出
        
        Args:
            start_row: 开始行（数据源，1-based，不包含表头）
            end_row: 结束行
            export_format: 导出格式
                "excel" - 合并为单个 Excel
                "pdf" - 导出多个 PDF（需要 win32com）
                "both" - 两者都
            output_dir: 输出目录
            field_map: 字段映射 {占位符: 字段名}
        
        Returns:
            导出的文件路径列表
        """
        if not self.engine:
            raise RuntimeError("请先载入模板")
        
        if not self.data_list:
            raise RuntimeError("请先载入数据")
        
        if field_map:
            self.field_map = field_map
        
        if not self.field_map:
            raise RuntimeError("请提供字段映射")
        
        # 创建输出目录
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        # 选择数据行
        if start_row is None:
            start_row = 1
        if end_row is None:
            end_row = len(self.data_list)
        
        selected_data = self.data_list[start_row - 1:end_row]
        
        output_files = []
        
        if export_format == "excel":
            # 合并为单个 Excel
            timestamp = self._get_timestamp()
            output_file = output_path / f"Merged_{timestamp}.xlsx"
            
            result = self.engine.fill_and_export(
                data_list=selected_data,
                field_map=self.field_map,
                output_path=str(output_file)
            )
            
            output_files.append(result)
        
        elif export_format == "pdf":
            # 导出多个 PDF（需要 win32com）
            try:
                import win32com.client
                
                # 先生成 Excel
                temp_excel = output_path / "temp.xlsx"
                self.engine.fill_and_export(
                    data_list=selected_data,
                    field_map=self.field_map,
                    output_path=str(temp_excel)
                )
                
                # 导出 PDF
                excel_app = win32com.client.Dispatch("Excel.Application")
                excel_app.Visible = False
                
                wb = excel_app.Workbooks.Open(str(temp_excel.absolute()))
                
                for i, sheet in enumerate(wb.Sheets):
                    pdf_name = selected_data[i].get(
                        list(self.field_map.values())[0],
                        f"Sheet_{i + 1}"
                    )
                    pdf_path = output_path / f"{pdf_name}.pdf"
                    sheet.ExportAsFixedFormat(0, str(pdf_path.absolute()))  # 0 = xlTypePDF
                    output_files.append(str(pdf_path))
                
                wb.Close(False)
                excel_app.Quit()
                
                # 删除临时文件
                temp_excel.unlink()
                
                print(f"[OK] 已生成 {len(output_files)} 个 PDF 文件")
                
            except ImportError:
                print("[WARNING] win32com 未安装，无法导出 PDF")
                print("[INFO] 请安装: pip install pywin32")
        
        elif export_format == "both":
            # 两者都导出
            excel_files = self.fill_and_export(
                start_row, end_row, "excel", output_dir, field_map
            )
            pdf_files = self.fill_and_export(
                start_row, end_row, "pdf", output_dir, field_map
            )
            output_files = excel_files + pdf_files
        
        return output_files
    
    def _get_timestamp(self) -> str:
        """获取时间戳"""
        from datetime import datetime
        return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
    def close(self):
        """关闭引擎"""
        if self.engine and hasattr(self.engine, 'close'):
            self.engine.close()


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description="统一模板填充器")
    parser.add_argument("--template", required=True, help="模板 Excel 路径")
    parser.add_argument("--data", help="数据源 Excel 路径")
    parser.add_argument("--config", help="配置文件 JSON（包含 data 和 fields）")
    parser.add_argument("--engine", choices=["auto", "openpyxl", "zip"], default="auto", help="引擎类型")
    parser.add_argument("--output", default="./output.xlsx", help="输出路径")
    parser.add_argument("--scan", action="store_true", help="仅扫描占位符")
    
    args = parser.parse_args()
    
    # 创建填充器
    filler = TemplateFiller(
        template=args.template,
        engine_type=args.engine
    )
    
    if args.scan:
        # 仅扫描
        placeholders = filler.scan_placeholders()
        print(f"\n发现占位符: {placeholders}")
        
        engine_name = "ZIP" if isinstance(filler.engine, ZIPEngine) else "openpyxl"
        print(f"引擎类型: {engine_name}")
    
    elif args.config:
        # 从配置文件读取
        with open(args.config, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
        
        filler.field_map = cfg.get('fields', {})
        filler.data_list = cfg.get('data', [])
        
        output_path = filler.fill_and_export(
            output_dir=str(Path(args.output).parent),
            field_map=filler.field_map
        )
        
        print(f"\n✅ 完成！请在 Excel 中打开验证: {output_path}")
    
    elif args.data:
        # 从数据源读取
        filler.load_data(args.data)
        
        placeholders = filler.scan_placeholders()
        
        # 自动生成字段映射
        field_map = {ph: ph for ph in placeholders}
        
        output_files = filler.fill_and_export(
            field_map=field_map,
            output_dir=str(Path(args.output).parent)
        )
        
        print(f"\n✅ 完成！已生成 {len(output_files)} 个文件")
    
    else:
        print("[ERROR] 请提供 --config 或 --data 参数")
        sys.exit(1)
    
    filler.close()


if __name__ == "__main__":
    main()
