# 引擎选择与使用

---

## 自动引擎选择机制

技能会自动检测模板类型并选择最优引擎：

```python
def detect_template_type(template_path: str) -> str:
    """
    检测模板类型
    
    Returns:
        "zip" - 含图片/打印设置/绘图
        "openpyxl" - 无图片
    """
    import zipfile
    
    with zipfile.ZipFile(template_path) as zf:
        names = zf.namelist()
        
        # 检查是否包含图片
        has_images = any(name.startswith('xl/media/') for name in names)
        
        # 检查是否包含打印设置
        has_printer = any('printerSettings' in name for name in names)
        
        # 检查是否包含 DrawingML
        has_drawings = any('drawings' in name for name in names)
        
        return "zip" if (has_images or has_printer or has_drawings) else "openpyxl"
```

---

## openpyxl 引擎

**适用场景：** 无图片模板，纯数据填充

**优势：**
- 代码简洁、易于维护
- 纯 Python 实现，无需额外依赖
- 适合快速开发

**限制：**
- 会丢失图片
- 会丢失打印设置
- 会丢失 DrawingML 绘图

**使用示例：**

```python
from scripts.engines.openpyxl_engine import OpenPYXLEngine

engine = OpenPYXLEngine(template="模板.xlsx")
engine.load_data("数据源.xlsx")
engine.fill_placeholders(field_map={"{编号}": "编号", "{材料}": "材料"})
engine.save("输出.xlsx")
```

---

## ZIP 引擎 ★核心

**适用场景：** 含图片模板、含打印设置模板

**优势：**
- ✅ 完美保留图片
- ✅ 完美保留打印设置
- ✅ 完美保留二进制资源
- ✅ 无 openpyxl 依赖，纯标准库

**核心技术：**

绕过 openpyxl，直接操作 XLSX 内部 XML：

```
模板.xlsx（ZIP）
  → 读取全部文件到内存
  → 修改 worksheet XML（t="s" → inlineStr，保留 s 属性）
  → 生成新的 sheet2~sheetN
  → 更新 workbook.xml / rels / Content_Types
  → 直接写入新 ZIP（保留所有原始资源）
```

**使用示例：**

```python
from scripts.engines.zip_engine import ZIPEngine

engine = ZIPEngine(template="报批表模板.xlsx")
engine.load_data("材料审批总表.xlsx")
engine.fill_and_generate_sheets(
    start_row=2,
    end_row=11,
    field_map={"{编号}": "编号", "{材料}": "材料"}
)
engine.save_all("./报批表")
```

---

## 引擎对比

| 特性 | openpyxl 引擎 | ZIP 引擎 |
|:---|:---:|:---:|
| 图片保留 | ❌ | ✅ |
| 打印设置保留 | ❌ | ✅ |
| DrawingML 保留 | ❌ | ✅ |
| 代码复杂度 | 低 | 高 |
| 依赖 | openpyxl | 无（标准库）|
| 适用模板 | 纯数据 | 含图片/打印设置 |

---

## 执行建议

1. **默认使用自动检测** — 让技能自动选择引擎
2. **含图片模板优先 ZIP 引擎** — 确保图片不丢失
3. **测试验证** — 生成后在 Excel 中打开验证

---

_最后更新：2026-06-02_
