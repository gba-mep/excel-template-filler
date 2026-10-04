#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
引擎基类 - 定义统一的模板填充接口

所有引擎（openpyxl, zip）都继承此基类

作者: gba-mep
版本: v2.1
日期: 2026-06-02
"""

from abc import ABC, abstractmethod
from typing import Dict, List, Any, Optional
from pathlib import Path


class BaseEngine(ABC):
    """
    模板填充引擎基类
    
    定义所有引擎必须实现的接口
    """
    
    def __init__(self):
        """初始化引擎"""
        self.template_path: Optional[str] = None
        self.is_loaded: bool = False
    
    @abstractmethod
    def load_template(self, template_path: str) -> bool:
        """
        载入模板文件
        
        Args:
            template_path: 模板 Excel 路径
        
        Returns:
            是否成功载入
        """
        pass
    
    @abstractmethod
    def scan_placeholders(self) -> List[str]:
        """
        扫描模板中的占位符
        
        Returns:
            占位符列表
        """
        pass
    
    @abstractmethod
    def fill_template(
        self,
        data: Dict[str, Any],
        field_map: Dict[str, str]
    ) -> Any:
        """
        填充单个模板
        
        Args:
            data: 数据字典
            field_map: 字段映射 {占位符: 字段名}
        
        Returns:
            填充后的对象（具体类型由子类定义）
        """
        pass
    
    @abstractmethod
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
        pass
    
    @staticmethod
    def detect_template_type(template_path: str) -> str:
        """
        检测模板类型（静态方法）
        
        Args:
            template_path: 模板路径
        
        Returns:
            "zip" 或 "openpyxl"
        """
        import zipfile
        
        try:
            with zipfile.ZipFile(template_path, 'r') as zf:
                names = zf.namelist()
                
                # 检查是否包含图片
                has_images = any(name.startswith('xl/media/') for name in names)
                
                # 检查是否包含打印设置
                has_printer = any('printerSettings' in name for name in names)
                
                # 检查是否包含 DrawingML
                has_drawings = any('drawings' in name for name in names)
                
                # 有图片、打印设置或绘图 → 使用 ZIP 引擎
                return "zip" if (has_images or has_printer or has_drawings) else "openpyxl"
        
        except Exception:
            return "openpyxl"
    
    def get_template_info(self) -> Dict[str, Any]:
        """
        获取模板信息
        
        Returns:
            模板信息字典
        """
        if not self.template_path:
            return {}
        
        path = Path(self.template_path)
        
        return {
            "path": str(path),
            "name": path.name,
            "size": path.stat().st_size if path.exists() else 0,
            "type": self.detect_template_type(self.template_path),
            "is_loaded": self.is_loaded
        }
