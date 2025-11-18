# PDF 转 DOCX Demo

基于 pdf2docx 开源库实现批量 PDF 转 DOCX 的 demo 脚本。

## 功能特点

- 批量转换 PDF 到 DOCX 格式
- 为每个文件创建独立输出目录，包含转换结果和调试信息
- 自动错误处理和 fallback 机制（当标准配置失败时自动尝试关闭 lattice 表格解析）
- 详细的转换日志记录
- 支持通过配置文件调整转换参数

## 安装依赖

建议使用虚拟环境：

```bash
# 创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
source venv/bin/activate  # Linux/Mac
# 或 venv\Scripts\activate  # Windows

# 安装依赖
pip install -r requirements.txt
```

## 使用方法

### 基础用法

将 PDF 文件放入 `pdf_data/` 目录，然后运行：

```bash
python convert.py
```

转换结果会输出到 `output/` 目录，每个 PDF 对应一个独立的子目录。

### 命令行参数

```bash
# 指定输入目录
python convert.py --input-dir /path/to/pdfs

# 指定输出目录
python convert.py --output-dir /path/to/output

# 启用调试模式（生成调试文件）
python convert.py --debug

# 转换单个文件
python convert.py --single /path/to/file.pdf
```

## 输出结构

每个 PDF 转换后的输出结构：

```
output/
  └── 文件名/
      ├── 文件名.docx          # 转换后的 Word 文档
      ├── conversion.log       # 转换日志
      └── debug/              # 调试信息（仅在启用调试模式时生成）
          ├── layout_page_0.json
          └── debug_page_0.pdf
```

## 配置文件

`config.yaml` 包含转换参数配置：

- `parse_lattice_table`: 是否启用网格线驱动的表格检测（默认 true）
- `enable_debug`: 是否生成调试文件（默认 false）
- 其他 pdf2docx 支持的参数

## 技术说明

本项目使用 [pdf2docx](https://github.com/dothinking/pdf2docx) 库作为核心转换引擎。该库基于 PyMuPDF 提取 PDF 元素，通过规则进行版面分析，并使用 python-docx 生成 DOCX 文档。

特别适合处理：
- WPS/Word 导出的电子 PDF（非扫描件）
- 包含复杂表格的合同、法律文件
- 需要保持格式一致性的文档转换

## 常见问题

**Q: 转换失败或表格识别错误？**

A: 脚本会自动尝试 fallback 模式（关闭 parse_lattice_table）。如果仍有问题，可以在 config.yaml 中调整参数。

**Q: 如何查看详细的转换过程？**

A: 使用 `--debug` 参数运行，会在每个文件的 debug/ 目录下生成布局分析文件。

## 项目结构

```
pdf_to_word_V2/
├── pdf_data/           # 输入 PDF 文件目录
├── output/            # 输出目录（自动创建）
├── config.yaml        # 配置文件
├── convert.py         # 核心转换脚本
├── requirements.txt   # Python 依赖
└── README.md         # 本文件
```

