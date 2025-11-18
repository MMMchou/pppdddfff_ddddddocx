# PaddleOCR PDF 转 Word/Markdown 工具

基于 PaddleOCR 实现的 PDF 文档转换工具，支持自动识别文字、表格、版面，并输出 Word 和 Markdown 格式。

## ✨ 主要特性

- 📄 **自动转换**: PDF → Word/Markdown，一键完成
- 🔍 **智能识别**: 文字识别 + 表格识别 + 版面分析
- 📑 **自动合并**: 多页 PDF 自动合并成单个文档（保持原页面结构）
- 📁 **智能整理**: 输出文件自动分类存放（最终文档、分页、图片、调试信息）
- 🚀 **批量处理**: 支持批量转换整个目录的 PDF 文件

## 🚀 快速开始

### 1. 激活环境

```bash
cd /home/zhouzhihan/pdf_to_word_V1
source paddleocr_env/bin/activate
```

### 2. 转换单个文件

```bash
python convert.py "你的PDF文件.pdf"
```

**示例**:
```bash
python convert.py "pdf_sample_data/picture_type/常规2.pdf"
```

### 3. 批量转换

```bash
python convert.py pdf_sample_data/picture_type --batch
```

## 📁 输出结构

转换完成后，输出会自动整理成以下结构：

```
output/文件名/
├── final/                          ← 🎯 最终结果（推荐使用）
│   ├── 文件名.docx                 ← 完整合并的 Word 文档
│   ├── 文件名.md                   ← 完整合并的 Markdown 文档
│   └── README.txt                  ← 说明文件
│
├── pages/                          ← 分页文档（每页独立）
│   ├── page_0.docx                 ← 第 1 页
│   ├── page_1.docx                 ← 第 2 页
│   └── ...
│
├── images/                         ← 可视化结果
│   ├── *_layout_det_res.png        ← 版面检测结果
│   ├── *_overall_ocr_res.png       ← OCR 识别结果
│   └── extracted/                  ← 从 PDF 提取的图片
│
└── debug/                          ← 调试信息
    ├── json/                       ← 原始 JSON 数据
    └── tex/                        ← LaTeX 公式
```

**推荐直接使用**: `output/文件名/final/文件名.docx`

## ⚙️ 配置说明

编辑 `config.yaml` 自定义配置：

```yaml
# 输出目录
output_dir: output

# 是否使用 GPU（需要 CUDA 支持）
use_gpu: false

# 是否启用表格识别
enable_table: true

# 是否启用 OCR 文本识别
enable_ocr: true

# 是否启用版面分析
enable_layout: true

# 识别语言：'ch'(中文) 或 'en'(英文)
lang: ch

# 输出格式：'word', 'markdown', 'both'
output_format: both
```

## 🛠️ 高级功能

### 1. 手动整理输出

如果输出目录未自动整理，可以手动运行：

```bash
# 整理单个目录
python organize_output.py output/文件名

# 批量整理所有目录
python organize_output.py output --batch
```

### 2. 对比文档内容

查看不同合并模式的文档统计信息：

```bash
python check_pages.py output/文件名/final
```

## 📝 使用示例

### 示例 1: 转换单个合同文件

```bash
python convert.py "合同.pdf"

# 结果位置：
# output/合同/final/合同.docx    ← 直接使用
# output/合同/final/合同.md      ← Markdown 版本
```

### 示例 2: 批量转换所有 PDF

```bash
python convert.py pdf_sample_data/picture_type --batch

# 每个 PDF 都会生成独立的输出目录
# output/文件1/final/文件1.docx
# output/文件2/final/文件2.docx
# ...
```

### 示例 3: 只需要 Word 格式

修改 `config.yaml`:
```yaml
output_format: word  # 只输出 Word，不生成 Markdown
```

## ❓ 常见问题

### Q1: 合并后的文档页数正确吗？
**A**: 默认使用带分页符的合并模式，每个原始 PDF 页面之间添加一个分页符，保持原文档的页面结构。

### Q2: 首次运行很慢？
**A**: 首次运行需要下载模型（约 200MB），需要几分钟。后续转换会快很多。

### Q3: 如何提高转换速度？
**A**: 
- 使用 GPU 加速（修改 `config.yaml` 中的 `use_gpu: true`，需要 CUDA 支持）
- 关闭不需要的功能（如 `enable_table: false`）

### Q4: 识别效果不好怎么办？
**A**:
- 确保 PDF 质量良好（清晰、无倾斜）
- 查看 `images/` 目录中的可视化结果，了解识别情况
- 中文文档确保 `lang: ch`，英文文档使用 `lang: en`

### Q5: 如何保持原 PDF 的分页结构？
**A**: 默认已启用带分页符模式，自动在每个原始 PDF 页面之间添加分页符，保持原文档结构。

### Q6: 支持哪些 PDF 类型？
**A**:
- ✅ 扫描版 PDF（图片型）- 通过 OCR 识别
- ✅ 文字版 PDF - 提取文字 + 版面分析
- ✅ 混合型 PDF - 智能处理

## 📂 项目结构

```
pdf_to_word_V1/
├── convert.py              ← 主转换脚本
├── organize_output.py      ← 输出整理脚本（支持单个/批量）
├── check_pages.py          ← 文档检查工具
├── config.yaml             ← 配置文件
├── README.md               ← 使用说明文档
├── PaddleOCR/              ← PaddleOCR 源码
├── paddleocr_env/          ← Python 虚拟环境
├── pdf_sample_data/        ← 示例 PDF 文件
└── output/                 ← 转换结果输出目录
```

## 🔧 依赖环境

- Python 3.8+
- PaddlePaddle
- PaddleOCR
- python-docx
- PyYAML
- tqdm

所有依赖已安装在 `paddleocr_env` 虚拟环境中。

## 📄 命令速查

```bash
# 激活环境（每次使用前）
source paddleocr_env/bin/activate

# 转换单个文件
python convert.py "文件.pdf"

# 批量转换
python convert.py 目录路径 --batch

# 整理单个输出（如需手动整理）
python organize_output.py output/文件名

# 批量整理所有输出
python organize_output.py output --batch

# 检查文档内容
python check_pages.py output/文件名/final
```

## ⚠️ 注意事项

1. **首次运行**: 会自动下载模型，需要网络连接和几分钟时间
2. **环境激活**: 每次使用前必须激活虚拟环境
3. **文件路径**: 支持中文路径和文件名
4. **输出覆盖**: 相同文件名会覆盖之前的输出
5. **最终文档**: 推荐直接使用 `final/` 目录中的合并文档

## 📊 性能参考

| PDF 类型 | 页数 | 转换时间 (CPU) | 文件大小 |
|---------|------|---------------|---------|
| 扫描件  | 4页  | ~2-3分钟      | ~40KB   |
| 文字版  | 4页  | ~1-2分钟      | ~35KB   |
| 含表格  | 10页 | ~5-6分钟      | ~80KB   |

*首次运行需额外增加模型下载时间*

## 📮 技术支持

遇到问题？
1. 查看 `output/文件名/images/` 中的可视化结果
2. 检查 `output/文件名/debug/json/` 中的原始数据
3. 确认虚拟环境已正确激活
4. 确保 PDF 文件质量良好

## 🎯 最佳实践

1. **推荐工作流程**:
   ```bash
   source paddleocr_env/bin/activate
   python convert.py "你的文件.pdf"
   # 直接使用 output/你的文件/final/你的文件.docx
   ```

2. **批量处理**:
   ```bash
   python convert.py pdf_directory --batch
   # 所有结果在 output/ 目录下
   ```

3. **质量检查**: 查看 `images/` 目录中的可视化结果，确认识别质量

4. **格式调整**: 如需调整，使用 Word 编辑 `final/*.docx` 文件

---

**项目位置**: `/home/zhouzhihan/pdf_to_word_V1`

**版本**: v1.1

**最后更新**: 2025-11-17
