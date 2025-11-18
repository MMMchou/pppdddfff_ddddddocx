## 一、场景再确认 & 总体策略
你的约束：

+ PDF 来源：**WPS / Word 导出的电子 PDF**（不是扫描件）
+ 文档类型：**合同 / 法律文件**（复杂表格、多级条款、强格式约束）
+ 目标：**格式尽量 1:1 还原到可编辑的 DOCX**，特别是表格布局
+ 技术栈：**Python 3 优先**，先脚本，后面可能做 Web 服务
+ 要求：**优先开源，避免重复造轮子 & 无意义工作**

在这种前提下：

✅ **推荐唯一主线方案：以 **`**pdf2docx**`** 为核心，做“标准 PDF 解析”路径。**

PaddleOCR 官方的排版还原模块，在处理“标准 PDF”时，也是基于 pdf2docx 做优化的，并且明确写了：  
**优点：对电子 PDF（非纸质扫描）还原效果最好，每页分页一致**；  
缺点：部分中英文混排可能乱码、有些内容被整体当成表格、图片恢复效果一般。([Gitee](https://gitee.com/yang_le_gittee/PaddleOCR/blob/release/2.6/ppstructure/recovery/README.md?utm_source=chatgpt.com))

你现在的输入就是标准电子 PDF，所以完全命中它的“优势”场景。

---

## 二、pdf2docx 能力概览（跟你项目最相关的点）
官方描述：

+ 用 **PyMuPDF**（fitz）读取 PDF 文本/图像/绘图等元素
+ 用规则在文档级做版面分析（section、header/footer、margin）
+ 将页面解析为段落、表格、图片，再用 `python-docx` 生成 DOCX ([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))
+ 支持表格：边框、底纹、合并单元格、竖排文本、部分边框隐藏、嵌套表格等 ([PyPI](https://pypi.org/project/pdf2docx/?utm_source=chatgpt.com))

**你最关心的：**

+ ✅ 文本：合并为段落（不是逐行文本框）
+ ✅ 表格：**真正的 Word 表格对象**，可编辑，而不是线条+文本框的假表格 ([PyPI](https://pypi.org/project/pdf2docx/?utm_source=chatgpt.com))
+ ✅ 图片：行内图片/浮动图片都能迁移
+ ✅ 多页：支持多页、多进程解析 ([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))

缺点（你需要在 overview 里写清楚、并设计补救策略）：

+ 没有真正的“Word 页眉/页脚”概念，PDF 的页眉内容会当普通文本处理 ([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))
+ 极少数页面可能因为把线误识别为表格而抛错（`making page error`），需要调整表格解析参数，例如关闭 `parse_lattice_table` ([Stack Overflow](https://stackoverflow.com/questions/77657464/while-running-pdf2docx-im-getting-making-page-error))
+ 某些怪异布局会被整体当成“大表格”，需要 QA 筛查 ([Gitee](https://gitee.com/yang_le_gittee/PaddleOCR/blob/release/2.6/ppstructure/recovery/README.md?utm_source=chatgpt.com))

---

## 三、整体架构设计（脚本阶段）
可以把整个项目抽象为 4 层：

1. **输入层（Sources）**
    - WPS / Word 导出 PDF
    - 文件夹 + 文件命名规范（如：合同号_版本号.pdf）
2. **转换核心层（Core Conversion）**
    - 核心库：`pdf2docx.Converter` ([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))
    - 关键 API：`convert()`, `debug_page()`, `extract_tables()`, `default_settings`
3. **质量控制层（QA & 再加工）**
    - 定义“金样本 PDF 集”（若干代表性合同）
    - 自动/半自动检查：页数、段落数量、表格数量、关键字段文本是否一致
    - 必要时用 `python-docx` 做二次修正（比如页眉/页脚）
4. **批处理 & 部署层（Batch & Deployment）**
    - 脚本版：命令行批量转换
    - 未来 Web：FastAPI + 队列（Celery/RQ）+ 存储（本地/对象存储）

你现在写 overview，可以按这个 4 层结构展开。

---

## 四、核心使用模式 & 参数调优（这是重点）
### 4.1 基本用法模式
**1）单文件基础转换（推荐）**

```python
from pdf2docx import Converter

def convert_pdf_to_docx(pdf_path: str, docx_path: str, **kwargs):
    cv = Converter(pdf_path)
    try:
        cv.convert(docx_path, start=0, end=None, **kwargs)
    finally:
        cv.close()
```

+ `start=0, end=None`：默认全文件，支持分页（后面讲批量/并行）([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))
+ `**kwargs`：传入所有布局/表格相关的调优参数（来源于 `default_settings`）

**2）一行式 **`**parse()**`** 简化接口**

还有 `pdf2docx.parse()` 这种快捷函数，但官方建议如果要调整参数、做调优，还是用 `Converter` 类。([GeeksforGeeks](https://www.geeksforgeeks.org/python/how-to-convert-a-pdf-to-document-using-python/?utm_source=chatgpt.com))

---

### 4.2 了解 & 利用 `default_settings`
`Converter.default_settings` 是一个 property，会返回默认解析配置（`dict`）。官方文档写明：传给 `convert()` 的 `**kwargs` 会覆盖这些默认配置。([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))

**最佳实践：在代码里打一次 log，把默认配置 dump 出来，自己看一遍。**

```python
from pdf2docx import Converter
import json

cv = Converter("sample.pdf")
settings = cv.default_settings
print(json.dumps(settings, indent=2, ensure_ascii=False))
cv.close()
```

你可以把这个结果保存到项目文档里，作为“当前版本 pdf2docx 的配置基线”，以后如果升级库版本，可以重跑对比差异。

常见/重要的配置键（名字可能略有出入，建议以实际 dump 为准）：

+ **表格相关**
    - `parse_lattice_table`: 是否使用“网格线驱动”的表格检测（lattice 模式）
        * 复杂线条/装饰线比较多的 PDF，有时会误判出“伪表格”，甚至触发 `making page error`，StackOverflow 的解决方案就是把这个参数关掉：  
`cv.convert(..., parse_lattice_table=False)`([Stack Overflow](https://stackoverflow.com/questions/77657464/while-running-pdf2docx-im-getting-making-page-error))
    - 某些设置是控制表格线条/间距阈值的（比如 line / gap 阈值），决定是不是把一组文字归为同一行/列 → 用来调优复杂合同表格效果
+ **文本/段落相关**
    - 行/段落合并阈值（如：行间距阈值、同段落的水平对齐容差）
    - 段落缩进和对齐识别（影响 `左对齐/居中/两端对齐`）([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))
+ **多进程相关**
    - 是否启用多进程解析（在 `start/end` 连续时才生效）([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))

你不需要记住每个键的名字，但要在 overview 里说明：

✅ 我们会在项目初始化阶段 dump 一份 `default_settings`，  
✅ 然后针对“合同文档 + 复杂表格”场景做少量修改（例如关闭 lattice 模式，微调表格检测阈值），  
✅ 并把调整过的配置固化为 YAML/JSON 配置文件，便于版本管理和回滚。

---

### 4.3 复杂表格场景的具体调优策略
针对合同常见的复杂表格（合并单元格、多级表头、跨页等），你可以这样设计策略：

1. **先跑一轮“默认配置 + 样本合同”**
    - 选 10–20 份复杂度不同的合同 PDF（多页、长表格、有合并单元格、有没有边框的表格）
    - 用默认配置转换成 DOCX
    - 人工检查：
        * 表格个数是否一致
        * 行列结构是否正确
        * 合并单元格是否还原
        * 有没有整页被当成一个巨型表格的问题（Paddle 文档提到过这种现象）([Gitee](https://gitee.com/yang_le_gittee/PaddleOCR/blob/release/2.6/ppstructure/recovery/README.md?utm_source=chatgpt.com))
2. **遇到报错或误判：优先调整 **`**parse_lattice_table**`
    - 如果日志出现 `Ignore page X due to making page error: list index out of range`，StackOverflow 经验表明是“把线误判成表格结构失败”导致。
    - 此时可以对该类文档用：

```python
cv.convert(docx_path, parse_lattice_table=False)
```

让 pdf2docx 把表格检测从“网格线模式”切换到更宽松的模式，通常能避免报错，并在复杂线条背景下得到更合理的文字布局。([Stack Overflow](https://stackoverflow.com/questions/77657464/while-running-pdf2docx-im-getting-making-page-error))

3. **对少数特例可按“文档级配置模板”处理**
    - 比如你发现：
        * 某一批合同模板 A：表格线非常规，很容易误判 → 给模板 A 配一个专用配置（关闭 lattice + 放宽行距阈值）
        * 模板 B：表格边界很规整 → 用默认配置即可
    - 你可以设计一个“模板识别 + 配置路由”的架构：
        * 按文件名 / 目录 / PDF 元数据，决定用哪一套配置参数

---

### 4.4 调试与排版问题定位：`debug_page` / `serialize`
pdf2docx 提供了**调试 API**，非常适合你做“为什么这一页表格炸了”的分析：([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))

```python
cv = Converter(pdf_path)
cv.debug_page(
    i=page_index,
    docx_filename="debug_page_5.docx",
    debug_pdf="debug_layout_page_5.pdf",
    layout_file="layout_page_5.json",
    # 这里同样可以传解析参数
    parse_lattice_table=False
)
cv.close()
```

+ `debug_pdf`：会生成一个带布局标记的 PDF（例如标出文本块、表格区域）
+ `layout_file`：是解析后的中间 JSON，包含每个块的坐标、类型、层级结构

你可以在项目文档里写：

复杂页面/表格如果出现结构问题，将使用 `debug_page` 生成布局调试文件，辅助分析 pdf2docx 的识别逻辑，并据此调整参数。

同时，`serialize()` / `deserialize()` 允许你把整个解析结果存为 JSON，再恢复成 `pages` 数据结构，方便做后处理或对比不同配置的结果。([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))

---

### 4.5 页眉/页脚 & 分页策略
pdf2docx 在“文档级分析”阶段，会尝试识别 header/footer，但**不会自动转成 Word 真页眉/页脚**，只是用来帮助判断正文区域。([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))

**推荐策略：**

1. 转换后，如果你非常在意“Word 真页眉/页脚”：
    - 用脚本/人工检查每页顶部重复的几行内容（例如“XX 公司合同编号：xxx”）
    - 将其移动到 Word 页眉中，并删除正文中的重复文本
    - 这个步骤可以半自动：脚本先标记可能的重复片段，人来确认
2. 保持 PDF 分页与 Word 分页一致：
    - 默认 `convert()` 会让每个 PDF 页在 DOCX 中对应 1 页，便于比对（Paddle 文档也提到“每页保持不变”的优点）。([Gitee](https://gitee.com/yang_le_gittee/PaddleOCR/blob/release/2.6/ppstructure/recovery/README.md?utm_source=chatgpt.com))
    - 如果后面需要对 Word 做大量编辑，可以再统一移除多余的分页符。

---

### 4.6 批量转换 & 异常隔离
一个简单稳妥的批处理脚本骨架（脚本版本）：

```python
from pdf2docx import Converter
from pathlib import Path
import json
import logging

logging.basicConfig(level=logging.INFO)

def convert_single(pdf_path: Path, out_dir: Path, settings_override=None):
    docx_path = out_dir / (pdf_path.stem + ".docx")
    cv = Converter(str(pdf_path))
    try:
        kwargs = settings_override or {}
        cv.convert(str(docx_path), start=0, end=None, **kwargs)
        logging.info(f"OK: {pdf_path.name} -> {docx_path.name}")
    except Exception as e:
        logging.exception(f"FAIL: {pdf_path.name}: {e}")
        # 可在这里加 fallback，比如关闭 lattice 再试一次
    finally:
        cv.close()

def batch_convert(input_dir: str, output_dir: str, settings_override=None):
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(in_dir.glob("*.pdf"))
    logging.info(f"Found {len(pdf_files)} pdf files")

    for pdf in pdf_files:
        convert_single(pdf, out_dir, settings_override=settings_override)

if __name__ == "__main__":
    batch_convert("input_pdfs", "output_docx")
```

几点实践要点：

+ **每个文件都 try/except 包裹**，避免一个坏文件拖死整批
+ 保留 conversion log（info + error），方便对照具体文件
+ 可以在失败时自动改用 `parse_lattice_table=False` 再试一次（你可以记录“用 fallback 成功”的文件清单）

如果你后续要并行（多进程）：

+ 利用 `concurrent.futures.ProcessPoolExecutor`，对“文件粒度”做多进程
+ 注意控制最大并发数（比如 CPU 核心数或者略少）
+ pdf2docx 内部对“页”也可以多进程，但只在 `start/end` 连续时启用([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))

---

## 五、质量控制策略（特别针对法律合同）
因为“完美还原”在实践中永远是“尽量接近 100%”，所以你在 overview 里最好明确一个 QA 策略，而不是假设“pdf2docx 永远没问题”。

**建议的 QA 体系：**

1. **金样本集（Golden Set）**
    - 至少选 20–50 份典型合同：
        * 不同模版（租赁、采购、服务、保密等）
        * 不同复杂度（简单条款 vs 超多表格的技术协议）
    - 每次改配置/升级 pdf2docx 都跑一遍，做“回归对比”
2. **自动检查项（可写工具脚本）：**

对比 PDF 与 DOCX 的：

    - 页数（Page count）
    - 表格数量（粗略可用 pdf2docx 的 `extract_tables()` 来数）([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))
    - 特定关键字段是否存在（比如“合同编号”、“甲方”、“乙方”这些关键词）

通过这些自动指标快速筛出“明显炸掉”的结果。

3. **人工 spot check（重点在表格 & 页眉）**

对自动指标没问题的文档：

    - 抽样逐页对照表格：结构 & 合并单元格
    - 检查每页上方/下方是否重复正文内容（页眉/页脚）
    - 检查长篇条款是否被拆成奇怪的段落或放进表格
4. **问题分类 & 配置迭代**
    - 把问题归类：
        * 某类模板总是出问题 → 单独调配置
        * 某个参数调整会对绝大部分文档更好 → 合并到全局配置
    - 记录配置版本号和对应的摘要（可以用 git 管理 YAML 配置）

---

## 六、部署到 Web 服务的架构建议（概览级）
当你从“脚本”升级到“Web 服务”，整体架构可以是这样：

### 6.1 典型结构
+ **API 层（FastAPI / Flask）**
    - 提供 `POST /convert`：上传 PDF，返回任务 ID 或直接返回 DOCX 文件
    - 限制上传大小/页数（例如 ≤ 200 页或 ≤ 20MB）
+ **任务层（Celery / RQ / 自己封装的进程池）**
    - 把长时间的转换任务移到后台
    - 控制并发度（例如同时最多跑 4–8 个转换任务）
+ **存储层**
    - 临时文件目录（PDF 上传、DOCX 输出）
    - 若要保存历史结果，可以用对象存储（比如本地 MinIO/S3）
+ **监控 & 日志**
    - log：记录每个任务的输入文件名、页数、用时、是否使用 fallback 参数
    - metric：每分钟处理文档数、失败率等

在任务处理逻辑里，就调用你上面写好的 `convert_single()`，完全复用同一套 pdf2docx 参数和 QA 策略。

### 6.2 环境与依赖
+ 镜像 / 服务器需要安装：
    - Python 3.x
    - `pdf2docx`（间接依赖 PyMuPDF）([PyPI](https://pypi.org/project/pdf2docx/?utm_source=chatgpt.com))
    - 对应系统库（MuPDF 动态库，通常 pip 安装时已经解决，但在某些 Linux 需要额外包）
+ 字体：
    - 如果合同有中文/特殊字体，尽量在服务器上安装同样字体，避免 Word 打开时替换字体导致布局变化

---

## 七、小结：可以直接写在 overview 里的核心结论
你可以在项目总述里，用接近下面这种总结：

本项目面向“WPS/Word 导出的合同类 PDF 文档”，目标是在保证版面和表格结构尽可能一致的前提下，将 PDF 高保真转换为可编辑的 Word（DOCX）文件。

在技术路线选择上，我们采用开源库 **pdf2docx** 作为核心转换引擎。该库基于 PyMuPDF 提取 PDF 中的文本、图片和线条信息，通过规则进行版面分析，并使用 `python-docx` 生成 DOCX 文档。它原生支持复杂表格（包括合并单元格、嵌套表格、部分边框隐藏等），并尽可能保留字体、对齐方式等格式信息，非常适合本项目所处理的“标准电子 PDF” 合同场景。([PyPI](https://pypi.org/project/pdf2docx/?utm_source=chatgpt.com))

在实现层面，我们围绕 `pdf2docx.Converter` 封装统一的转换接口，并通过 `default_settings` 机制进行参数调优：对可能导致错误的 lattice 模式表格解析配置进行优化（必要时关闭 `parse_lattice_table` 以避免复杂线条误判表格），同时利用 `debug_page()` 和 `serialize()` 等调试能力，对典型样本合同进行细粒度排版分析。([pdf2docx.readthedocs.io](https://pdf2docx.readthedocs.io/en/latest/api/pdf2docx.converter.html))

为提升稳定性和可维护性，系统引入“金样本文档集”作为回归测试基础，每次配置调整或版本升级前均对其进行批量转换，并对页数、表格数、关键字段等自动指标及人工 spot check 结果进行验证。对于页眉页脚等特殊版面元素，则采用“转换后检测重复内容 + 半自动迁移到 Word 页眉/页脚区域”的方式进行补充处理。

在部署层面，当前阶段采用 Python 脚本方式对文件夹内 PDF 进行批量转换；后续将基于 FastAPI 和任务队列封装为 Web 服务，对外提供 `POST /convert` 等接口，并通过并发控制、日志记录及错误隔离机制，保证系统在高并发文档转换场景下的性能与可用性。

---

如果你愿意，下一步我可以：

+ **帮你写一份真正可以直接 copy 进文档的“项目总览.md”完整版本**，
+ 或者帮你设计“合同金样本评测用的小脚本”，专门对比表格数量和关键字段。

