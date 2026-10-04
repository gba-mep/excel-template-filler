#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BQ 页合并器 - 将申请表 PDF 与 BQ 页面合并

作者: gba-mep
版本: v2.1
日期: 2026-06-02
"""

import os
import re
import sys
import io
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    import fitz  # PyMuPDF
except ImportError:
    print("[ERROR] 请安装 PyMuPDF: pip install PyMuPDF")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("[ERROR] 请安装 openpyxl: pip install openpyxl")
    sys.exit(1)


class BQMerger:
    """
    BQ 页合并器
    
    功能：
    1. 读取总表获取 EL 编号 + 材料名 + BQ 编号
    2. 在 BQ PDF 中搜索每个编号所在的页码（正则匹配）
    3. 将报批表 PDF 与对应 BQ 整页合并
    4. 输出文件名格式：EL-XXX 材料名.pdf
    
    使用示例：
        merger = BQMerger()
        merger.load_bq_pdf("BQ.pdf")
        merger.load_zongbiao("总表.xlsx")
        merger.merge_pdfs("./input", "./output")
    """
    
    def __init__(self):
        """初始化 BQ 合并器"""
        self.bq_doc: Optional[fitz.Document] = None
        self.bq_page_index: Dict[str, int] = {}  # {BQ编号: 页码}
        self.zongbiao_data: List[Dict[str, str]] = []
    
    def load_bq_pdf(self, bq_path: str) -> bool:
        """
        载入 BQ PDF 并建立页码索引
        
        Args:
            bq_path: BQ PDF 路径
        
        Returns:
            是否成功载入
        """
        try:
            self.bq_doc = fitz.open(bq_path)
            
            print(f"[OK] 已载入 BQ PDF: {bq_path}")
            print(f"  总页数: {self.bq_doc.page_count}")
            
            # 建立页码索引
            self._build_page_index()
            
            return True
            
        except Exception as e:
            print(f"[ERROR] 载入 BQ PDF 失败: {e}")
            return False
    
    def _build_page_index(self):
        """建立 BQ 编号 → 页码索引"""
        self.bq_page_index.clear()
        
        for page_num in range(self.bq_doc.page_count):
            text = self.bq_doc[page_num].get_text()
            
            # 匹配 BQ 编号格式：1.1, 2.1.3, 1.1-a 等
            matches = re.findall(r'\b(\d+\.\d+(?:[-.]\d+)?)\b', text)
            
            for m in matches:
                if m not in self.bq_page_index:
                    self.bq_page_index[m] = page_num + 1  # 转为 1-based 页码
        
        print(f"[OK] 已建立页码索引: {len(self.bq_page_index)} 个 BQ 编号")
    
    def load_zongbiao(
        self,
        zongbiao_path: str,
        sheet_index: int = 0,
        start_row: int = 7,
        col_bq: int = 1,
        col_el: int = 2,
        col_name: int = 3
    ) -> int:
        """
        读取数据源中的条目列表
        
        Args:
            zongbiao_path: 总表 Excel 路径
            sheet_index: 工作表索引
            start_row: 数据起始行
            col_bq: BQ 编号列（1-based）
            col_el: EL 编号列
            col_name: 材料名列
        
        Returns:
            读取的材料数量
        """
        try:
            wb = openpyxl.load_workbook(zongbiao_path, data_only=True)
            ws = wb.worksheets[sheet_index]
            
            self.zongbiao_data.clear()
            
            for row_idx in range(start_row, ws.max_row + 1):
                bq = ws.cell(row_idx, col_bq).value
                el = ws.cell(row_idx, col_el).value
                name = ws.cell(row_idx, col_name).value
                
                # 跳过空行
                if bq is None and el is None:
                    break
                
                self.zongbiao_data.append({
                    'bq': str(bq).strip() if bq else '',
                    'el': str(el).strip() if el else '',
                    'name': str(name).strip() if name else ''
                })
            
            wb.close()
            
            print(f"[OK] 已读取总表: {zongbiao_path}")
            print(f"  材料数: {len(self.zongbiao_data)}")
            
            return len(self.zongbiao_data)
            
        except Exception as e:
            print(f"[ERROR] 读取总表失败: {e}")
            return 0
    
    def match_bq_pages(self) -> Tuple[int, int]:
        """
        匹配 BQ 页码
        
        Returns:
            (成功数, 失败数)
        """
        if not self.bq_doc:
            raise RuntimeError("请先载入 BQ PDF")
        
        if not self.zongbiao_data:
            raise RuntimeError("请先载入总表")
        
        success = 0
        failed = 0
        
        for item in self.zongbiao_data:
            bq = item['bq']
            
            if bq in self.bq_page_index:
                # 直接匹配
                item['bq_page'] = self.bq_page_index[bq]
                success += 1
            else:
                # 尝试模糊匹配（去掉后缀）
                found = False
                parts = bq.rsplit('.', 1)
                
                for n in range(len(parts), 0, -1):
                    base = '.'.join(parts[:n])
                    if base in self.bq_page_index:
                        item['bq_page'] = self.bq_page_index[base]
                        found = True
                        success += 1
                        break
                
                if not found:
                    item['bq_page'] = None
                    failed += 1
                    print(f"  [WARN] 未找到 BQ 页: {item['el']}  BQ={bq}")
        
        print(f"[OK] BQ 页匹配完成: {success} 成功, {failed} 失败")
        
        return success, failed
    
    def merge_pdfs(
        self,
        input_dir: str,
        output_dir: str
    ) -> Tuple[int, int]:
        """
        合并 PDF
        
        Args:
            input_dir: 报批表 PDF 输入目录
            output_dir: 合并后 PDF 输出目录
        
        Returns:
            (成功数, 跳过数)
        """
        if not self.bq_doc:
            raise RuntimeError("请先载入 BQ PDF")
        
        # 创建输出目录
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        input_path = Path(input_dir)
        
        success = 0
        skipped = 0
        
        print(f"\n[INFO] 开始合并 PDF...")
        
        for item in self.zongbiao_data:
            el = item['el']
            bq_page = item.get('bq_page')
            
            # 查找报批表 PDF
            baopiao_file = None
            for f in input_path.iterdir():
                if f.suffix.lower() == '.pdf' and f.name.lower().startswith(el.lower()):
                    baopiao_file = f
                    break
            
            if not baopiao_file:
                print(f"  [SKIP] 无报批表 PDF: {el}")
                skipped += 1
                continue
            
            if bq_page is None:
                print(f"  [SKIP] 无 BQ 页: {el}")
                skipped += 1
                continue
            
            # 合并
            try:
                output = fitz.open()
                
                # 插入报批表
                baopiao_doc = fitz.open(str(baopiao_file))
                output.insert_pdf(baopiao_doc)
                baopiao_doc.close()
                
                # 插入 BQ 页
                output.insert_pdf(
                    self.bq_doc,
                    from_page=bq_page - 1,
                    to_page=bq_page - 1
                )
                
                # 生成输出文件名
                safe_name = re.sub(r'[\\/:*?"<>|]', '-', item['name'])
                output_name = f"{el} {safe_name}.pdf"
                output_file = output_path / output_name
                
                output.save(str(output_file))
                output.close()
                
                print(f"  [OK] {el}  {item['name'][:18]}  → BQ p{bq_page}")
                success += 1
                
            except Exception as e:
                print(f"  [ERROR] 合并失败: {el}  {e}")
                skipped += 1
        
        print(f"\n[OK] 合并完成: {success} 成功, {skipped} 跳过")
        print(f"  输出目录: {output_dir}")
        
        return success, skipped
    
    def close(self):
        """关闭 BQ PDF"""
        if self.bq_doc:
            self.bq_doc.close()
            self.bq_doc = None


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description="BQ 页合并器")
    parser.add_argument("--zongbiao", required=True, help="总表 Excel 路径")
    parser.add_argument("--bq", required=True, help="BQ PDF 路径")
    parser.add_argument("--input", required=True, help="报批表 PDF 输入目录")
    parser.add_argument("--output", required=True, help="合并后 PDF 输出目录")
    parser.add_argument("--sheet", type=int, default=0, help="工作表索引（默认 0）")
    parser.add_argument("--start-row", type=int, default=7, help="数据起始行（默认 7）")
    
    args = parser.parse_args()
    
    merger = BQMerger()
    
    # 载入 BQ PDF
    if not merger.load_bq_pdf(args.bq):
        sys.exit(1)
    
    # 载入总表
    if not merger.load_zongbiao(args.zongbiao, args.sheet, args.start_row):
        sys.exit(1)
    
    # 匹配 BQ 页
    merger.match_bq_pages()
    
    # 合并 PDF
    merger.merge_pdfs(args.input, args.output)
    
    merger.close()
    
    print("\n✅ 完成！")


if __name__ == "__main__":
    main()
