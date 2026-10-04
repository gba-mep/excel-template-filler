# 目录结构与模块说明

```
excel-template-filler/
├── src/
│   ├── template_filler.py       # 统一入口（自动引擎选择）
│   ├── engines/
│   │   ├── base_engine.py       # 引擎接口基类
│   │   ├── openpyxl_engine.py   # openpyxl 引擎（无图片）
│   │   └── zip_engine.py        # ZIP 引擎（图片保留）
│   ├── exporters/
│   │   └── bq_merger.py         # BQ 页合并器
│   ├── auto_linker.py           # 自动超链接创建器
│   ├── file_grabber.py          # 文件名抓取器
│   └── utils.py                 # 工具函数库
├── examples/
│   ├── data/                    # 示例数据文件
│   ├── templates/               # 示例模板
│   ├── example_basic.py         # 基本填充示例
│   ├── example_batch_pdf.py    # 批量 PDF 示例
│   └── example_auto_link.py    # 自动链接示例
├── references/                  # 模块文档
├── pyproject.toml
├── requirements.txt
└── README.md
```

## 依赖

| 依赖 | 版本 | 用途 |
|:---|:---|:---|
| openpyxl | >= 3.1.0 | 无图片模板引擎 |
| pywin32 | 可选 | PDF 导出（Windows） |
