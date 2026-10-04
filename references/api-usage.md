# API 用法详解

## 基本用法

```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from template_filler import TemplateFiller

# 初始化（自动选择引擎）
filler = TemplateFiller(
    data_source="data/sample_data.xlsx",
    template="templates/sample_template.xlsx"
)

# 检查引擎类型
print(f"Engine: {'ZIP' if filler.has_images() else 'openpyxl'}")

# 加载数据
filler.load_data()

# 扫描占位符
placeholders = filler.scan_placeholders()

# 填充并导出
output_files = filler.fill_and_export(
    field_map={
        "{ID}": "ID",
        "{Name}": "Name",
        "{Brand}": "Brand",
        "{Qty}": "Qty"
    },
    output_dir="./output"
)
```

---

## 双引擎架构

| 引擎 | 适用场景 | 优势 |
|:---|:---|:---|
| **openpyxl** | 无图片模板 | API 乾净，易维护 |
| **ZIP** | 含图片模板 | 完美保留图片、打印设置、二进制资源 |

`TemplateFiller` 根据模板内容自动选择引擎，无需手动配置。

---

## 完整流程（模板填充 + BQ 合并）

```python
# Step 1: 生成填充 PDF
filler = TemplateFiller("data.xlsx", "template.xlsx")
filler.load_data()
filler.fill_and_export(
    export_format="pdf", output_dir="./pdfs"
)

# Step 2: BQ 页合并
from exporters.bq_merger import BQMerger
merger = BQMerger()
merger.load_bq_pdf("bq_reference.pdf")
merger.load_zongbiao("data.xlsx")
merger.match_bq_pages()
merger.merge_pdfs("./pdfs", "./final_output")
```

---

## 自动链接

```python
from auto_linker import AutoLinker

linker = AutoLinker("summary.xlsx", 1, "./files")
linker.link_all()
```

---

## 限制

1. 占位符必须在 shared strings 中（普通单元格文本）
2. 单元格中占位符+文字混合不支持
3. BQ 合并需要带文本层的 PDF（不支持扫描图片）
