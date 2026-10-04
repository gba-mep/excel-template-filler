#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ZIP 引擎 - 纯 ZIP 方案处理含图片模板

完美保留图片、打印设置、二进制资源
通过直接操作 XLSX (ZIP) 内部 XML 实现

作者: gba-mep
版本: v2.1
日期: 2026-06-02
"""

import os
import re
import sys
import io
import json
import zipfile
import shutil
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple


class ZIPEngine:
    """
    ZIP 引擎 - 纯 ZIP 方案处理 Excel 模板
    
    核心能力：
    1. 完美保留图片、打印设置、二进制资源
    2. 支持占位符替换
    3. 批量生成多个 Sheet
    
    使用示例：
        engine = ZIPEngine()
        engine.load_template("模板.xlsx")
        placeholders = engine.scan_placeholders()
        output_files = engine.fill_and_export(data_list, field_map, "output.xlsx")
    """
    
    def __init__(self):
        """初始化 ZIP 引擎"""
        self.template_files: Dict[str, bytes] = {}
        self.shared_strings: List[str] = []
        self.sheet1_xml: str = ""
        self.has_images: bool = False
        self.has_printer_settings: bool = False
        
    def load_template(self, template_path: str) -> bool:
        """
        载入模板文件
        
        Args:
            template_path: 模板 Excel 路径
        
        Returns:
            是否成功载入
        """
        try:
            with zipfile.ZipFile(template_path, 'r') as zf:
                # 读取所有文件到内存
                self.template_files = {name: zf.read(name) for name in zf.namelist()}
            
            # 检测模板特徵
            self._detect_template_features()
            
            # 解析共享字符串
            self._parse_shared_strings()
            
            # 读取 sheet1 XML
            if 'xl/worksheets/sheet1.xml' in self.template_files:
                self.sheet1_xml = self.template_files['xl/worksheets/sheet1.xml'].decode('utf-8')
            
            print(f"[OK] 已载入模板: {template_path}")
            print(f"  - 图片: {'是' if self.has_images else '否'}")
            print(f"  - 打印设置: {'是' if self.has_printer_settings else '否'}")
            print(f"  - 共享字符串: {len(self.shared_strings)} 条")
            
            return True
            
        except Exception as e:
            print(f"[ERROR] 载入模板失败: {e}")
            return False
    
    def _detect_template_features(self):
        """检测模板特徵"""
        # 检测图片
        self.has_images = any(
            name.startswith('xl/media/') 
            for name in self.template_files.keys()
        )
        
        # 检测打印设置
        self.has_printer_settings = any(
            'printerSettings' in name 
            for name in self.template_files.keys()
        )
    
    def _parse_shared_strings(self):
        """解析共享字符串"""
        self.shared_strings.clear()
        
        if 'xl/sharedStrings.xml' in self.template_files:
            ss_raw = self.template_files['xl/sharedStrings.xml'].decode('utf-8')
            # 提取所有 <si>...</si> 内容
            self.shared_strings = re.findall(r'<si>(.*?)</si>', ss_raw, re.DOTALL)
    
    def scan_placeholders(self) -> List[str]:
        """
        扫描模板中的占位符
        
        Returns:
            占位符列表
        """
        placeholders = []
        
        for content in self.shared_strings:
            # 提取文本
            texts = re.findall(r'<t[^>]*>([^<]*)</t>', content)
            text = ''.join(texts).strip()
            
            # 检测占位符格式（支持 {} 和 {{}}）
            if (text.startswith('{') and text.endswith('}')) or \
               (text.startswith('{{') and text.endswith('}}')):
                placeholders.append(text)
        
        return placeholders
    
    def fill_template(
        self,
        data: Dict[str, Any],
        field_map: Dict[str, str]
    ) -> str:
        """
        填充单个模板
        
        Args:
            data: 数据字典
            field_map: 字段映射 {占位符: 字段名}
        
        Returns:
            填充后的 sheet XML
        """
        return self._build_inline_sheet(data, field_map)
    
    def fill_and_export(
        self,
        data_list: List[Dict[str, Any]],
        field_map: Dict[str, str],
        output_path: str
    ) -> str:
        """
        批量填充并导出
        
        Args:
            data_list: 数据列表
            field_map: 字段映射 {占位符: 字段名}
            output_path: 输出路径
        
        Returns:
            输出文件路径
        """
        if not self.template_files:
            raise RuntimeError("请先载入模板")
        
        print(f"[INFO] 开始生成 {len(data_list)} 个 Sheet...")
        
        # 复制模板文件
        new_files = {k: v for k, v in self.template_files.items()}
        
        # 生成每个 Sheet
        sheet_names = []
        for i, data in enumerate(data_list):
            # 获取名称（使用第一个字段的值）
            first_field = list(field_map.values())[0] if field_map else None
            sheet_name = str(data.get(first_field, f'Sheet{i + 1}'))[:31] if first_field else f'Sheet{i + 1}'
            sheet_names.append(sheet_name)
            
            # 填充模板
            xml_content = self._build_inline_sheet(data, field_map)
            
            if i == 0:
                # 第一项替换 sheet1
                new_files['xl/worksheets/sheet1.xml'] = xml_content.encode('utf-8')
            else:
                # 新建后续 Sheet
                sn = i + 1
                new_files[f'xl/worksheets/sheet{sn}.xml'] = xml_content.encode('utf-8')
                
                # 复制 sheet1 的关系文件（含 printerSettings + drawing）
                if 'xl/worksheets/_rels/sheet1.xml.rels' in self.template_files:
                    new_files[f'xl/worksheets/_rels/sheet{sn}.xml.rels'] = \
                        self.template_files['xl/worksheets/_rels/sheet1.xml.rels']
        
        print(f"[OK] 已生成 {len(sheet_names)} 个 Sheet")
        
        # 更新各种 XML 文件
        self._update_workbook_xml(new_files, sheet_names)
        self._update_workbook_rels(new_files, len(data_list))
        self._update_content_types(new_files, len(data_list))
        
        # 删除 sharedStrings.xml（已转为 inlineStr）
        new_files.pop('xl/sharedStrings.xml', None)
        
        # 写入输出文件
        self._write_output(new_files, output_path)
        
        return output_path
    
    def _build_inline_sheet(
        self,
        data: Dict[str, Any],
        field_map: Dict[str, str]
    ) -> str:
        """
        构建 inlineStr Sheet（核心算法）
        
        将 t="s"（共享字符串引用）转换为 t="inlineStr"（内联字符串）
        同时保留样式索引（s 属性）
        """
        # 占位符 → 实际值
        ph_to_val = {ph: self._esc(data.get(fk, '')) for ph, fk in field_map.items()}
        
        # 构建替换后的共享字符串纯文本列表
        new_ss = []
        for content in self.shared_strings:
            result = content
            for ph, val in ph_to_val.items():
                result = result.replace(ph, val)
            texts = re.findall(r'<t[^>]*>([^<]*)</t>', result)
            new_ss.append(''.join(texts).strip())
        
        # 转换 t="s" → t="inlineStr"，关键：保留原单元格 s 属性
        def cell_replacer(m):
            cell_xml = m.group(0)
            ref_m = re.search(r'r="([^"]+)"', cell_xml)
            ref = ref_m.group(1) if ref_m else ''
            
            # 提取并保留样式索引
            s_m = re.search(r'\bs="(\d+)"', cell_xml)
            s_attr = f' s="{s_m.group(1)}"' if s_m else ''
            
            v_m = re.search(r'<v>(\d+)</v>', cell_xml)
            if v_m:
                idx = int(v_m.group(1))
                if idx < len(new_ss):
                    val = new_ss[idx]
                    if val:
                        return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t>{val}</t></is></c>'
            return cell_xml
        
        return re.sub(r'<c [^>]*t="s"[^>]*>.*?</c>', cell_replacer, self.sheet1_xml, flags=re.DOTALL)
    
    def _update_workbook_xml(self, new_files: Dict[str, bytes], sheet_names: List[str]):
        """更新 workbook.xml"""
        wb_xml = self.template_files['xl/workbook.xml'].decode('utf-8')
        
        def sheet_rid(i):
            # i=0 → rId1（模板原有 sheet1），i≥1 → rId{i+4}（避免覆盖 theme/styles）
            return 'rId1' if i == 0 else f'rId{i + 4}'
        
        sheets_parts = [
            f'<sheet name="{self._esc(name)}" sheetId="{i + 1}" state="visible" r:id="{sheet_rid(i)}"/>'
            for i, name in enumerate(sheet_names)
        ]
        
        wb_xml = re.sub(
            r'<sheets>.*?</sheets>',
            '<sheets>' + ''.join(sheets_parts) + '</sheets>',
            wb_xml,
            flags=re.DOTALL
        )
        
        new_files['xl/workbook.xml'] = wb_xml.encode('utf-8')
    
    def _update_workbook_rels(self, new_files: Dict[str, bytes], num_sheets: int):
        """更新 workbook.xml.rels"""
        wb_rels = self.template_files['xl/_rels/workbook.xml.rels'].decode('utf-8')
        
        # 移除 sharedStrings 引用（已转为 inlineStr，不再需要）
        wb_rels = re.sub(r'<Relationship[^>]*sharedStrings[^>]*/>', '', wb_rels)
        
        # 追加新 Sheet（i=1..N → rId{i+4}）
        new_rel_lines = []
        for i in range(1, num_sheets):
            rid = i + 4
            new_rel_lines.append(
                f'  <Relationship Id="rId{rid}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                f'Target="worksheets/sheet{rid}.xml"/>'
            )
        
        # 在 </Relationships> 之前插入
        close_tag = '</Relationships>'
        idx = wb_rels.rfind(close_tag)
        if idx != -1:
            wb_rels = wb_rels[:idx] + '\n'.join(new_rel_lines) + '\n' + wb_rels[idx:]
        
        new_files['xl/_rels/workbook.xml.rels'] = wb_rels.encode('utf-8')
    
    def _update_content_types(self, new_files: Dict[str, bytes], num_sheets: int):
        """更新 Content_Types.xml"""
        ct = self.template_files['[Content_Types].xml'].decode('utf-8')
        
        # 移除 sharedStrings 引用
        ct = re.sub(r'\s*<Override[^>]*sharedStrings[^>]*/>', '', ct)
        
        # 追加新 Sheet 的 Override
        new_override_lines = [
            f'  <Override PartName="/xl/worksheets/sheet{i + 1}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(1, num_sheets)
        ]
        
        # 在 </Types> 之前插入
        close_tag = '</Types>'
        idx = ct.rfind(close_tag)
        if idx != -1:
            ct = ct[:idx] + '\n'.join(new_override_lines) + '\n' + ct[idx:]
        
        new_files['[Content_Types].xml'] = ct.encode('utf-8')
    
    def _write_output(self, new_files: Dict[str, bytes], output_path: str):
        """写入输出文件"""
        # 使用临时文件
        temp = os.path.join(os.environ.get('TEMP', '.'), '_zip_engine_temp.xlsx')
        
        with zipfile.ZipFile(temp, 'w', zipfile.ZIP_DEFLATED) as zf:
            for name, data in new_files.items():
                zf.writestr(name, data)
        
        # 复制到目标位置
        shutil.copy(temp, output_path)
        
        # 清理临时文件
        if os.path.exists(temp):
            os.remove(temp)
        
        # 验证输出
        size = os.path.getsize(output_path)
        print(f"[OK] 已生成: {output_path}")
        print(f"  大小: {size:,} bytes ({size / 1024:.1f} KB)")
    
    @staticmethod
    def _esc(s: str) -> str:
        """XML 特殊字符转义"""
        if s is None:
            return ''
        return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    
    @staticmethod
    def detect_template_type(template_path: str) -> str:
        """
        检测模板类型
        
        Args:
            template_path: 模板路径
        
        Returns:
            "zip" 或 "openpyxl"
        """
        try:
            with zipfile.ZipFile(template_path, 'r') as zf:
                names = zf.namelist()
                
                # 检查是否包含图片
                has_images = any(name.startswith('xl/media/') for name in names)
                
                # 检查是否包含打印设置
                has_printer = any('printerSettings' in name for name in names)
                
                return "zip" if (has_images or has_printer) else "openpyxl"
        
        except Exception:
            return "openpyxl"


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description="ZIP 引擎 - 纯 ZIP 方案处理 Excel 模板")
    parser.add_argument("--template", required=True, help="模板 Excel 路径")
    parser.add_argument("--config", help="配置文件 JSON（包含 data 和 fields）")
    parser.add_argument("--output", default="./output.xlsx", help="输出路径")
    parser.add_argument("--scan", action="store_true", help="仅扫描占位符")
    
    args = parser.parse_args()
    
    engine = ZIPEngine()
    
    # 载入模板
    if not engine.load_template(args.template):
        sys.exit(1)
    
    if args.scan:
        # 仅扫描
        placeholders = engine.scan_placeholders()
        print(f"\n发现占位符: {placeholders}")
    
    elif args.config:
        # 填充并导出
        with open(args.config, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
        
        output_path = engine.fill_and_export(
            data_list=cfg['data'],
            field_map=cfg['fields'],
            output_path=args.output
        )
        
        print(f"\n✅ 完成！请在 Excel 中打开验证: {output_path}")
    
    else:
        print("[ERROR] 请提供 --config 或 --scan 参数")
        sys.exit(1)


if __name__ == "__main__":
    main()
