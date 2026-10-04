# BQ 页合并器使用指南

---

## 功能概述

BQ 页合并器用于将报批表 PDF 与 BQ 标书页自动合并，生成最终递交文件。

**典型流程：**
```
报批表 PDF + BQ 页 = 最终递交文件
```

---

## 使用步骤

### Step 1: 初始化合并器

```python
from scripts.exporters.bq_merger import BQMerger

merger = BQMerger()
```

### Step 2: 载入 BQ PDF

```python
merger.load_bq_pdf("BQ标书.pdf")
```

**要求：**
- BQ PDF 必须有文字层（从 Word/Excel 转换来的 PDF）
- 如果是纯扫描图（无文字），需改用 OCR 或手动提供页码映射

### Step 3: 载入总表

```python
merger.load_zongbiao(
    zongbiao_path="材料审批总表.xlsx",
    sheet_index=0,
    start_row=7,
    col_bq=1,    # A 列：BQ 编号
    col_el=2,    # B 列：EL 编号
    col_name=3   # C 列：材料名
)
```

**参数说明：**
- `zongbiao_path`: 总表 Excel 路径
- `sheet_index`: Sheet 索引（默认 0）
- `start_row`: 数据起始行（默认 7）
- `col_bq`: BQ 编号所在列（默认 1 = A 列）
- `col_el`: EL 编号所在列（默认 2 = B 列）
- `col_name`: 材料名所在列（默认 3 = C 列）

### Step 4: 匹配 BQ 页码

```python
success, failed = merger.match_bq_pages()

print(f"成功匹配: {len(success)} 个")
print(f"匹配失败: {len(failed)} 个")
```

**匹配逻辑：**
1. 从总表读取 EL 编号和 BQ 编号
2. 在 BQ PDF 中搜索每个 BQ 编号
3. 记录每个编号所在的页码

### Step 5: 合并 PDF

```python
merger.merge_pdfs("./报批表PDF", "./最终输出")
```

**输入目录结构：**
```
./报批表PDF/
├── EL-001.pdf
├── EL-002.pdf
├── EL-003.pdf
└── ...
```

**输出文件名格式：**
```
EL-XXX 材料名.pdf
```

### Step 6: 关闭合并器

```python
merger.close()
```

---

## 完整示例

```python
from scripts.exporters.bq_merger import BQMerger

# 初始化
merger = BQMerger()

# 载入 BQ PDF
merger.load_bq_pdf("BQ标书.pdf")

# 载入总表
merger.load_zongbiao("材料审批总表.xlsx")

# 匹配 BQ 页码
success, failed = merger.match_bq_pages()

# 合并 PDF
merger.merge_pdfs("./报批表PDF", "./最终输出")

# 关闭
merger.close()

print("✅ 合并完成")
```

---

## 限制与注意

1. **BQ PDF 必须有文字层** — 扫描件无效
2. **总表格式必须规范** — EL 编号、BQ 编号、材料名列必须对齐
3. **输出目录会自动创建** — 无需手动创建
4. **匹配失败的项会跳过** — 不影响其他项合并

---

## 常见问题

### Q1: BQ PDF 是扫描件怎么办？

**方案 A：** 使用 OCR 工具提取文字层
**方案 B：** 手动提供页码映射（Excel 表格）

### Q2: 总表格式不规范怎么办？

调整 `load_zongbiao()` 参数：
- `col_bq`: 指定 BQ 编号实际所在列
- `col_el`: 指定 EL 编号实际所在列
- `col_name`: 指定材料名实际所在列

### Q3: 合并后文件太大怎么办？

使用 PDF 压缩工具：
```python
import fitz  # PyMuPDF

doc = fitz.open("大文件.pdf")
doc.save("压缩后.pdf", deflate=True, garbage=4)
doc.close()
```

---

_最后更新：2026-06-02_
