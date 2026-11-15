# office-toolbox

| 脚本 | 作用 |
|---|---|
| `docx2md.py` | `.docx` → `.md` |
| `md2docx.py` | `.md` → `.docx`（公式写成可编辑的 Word 公式） |
| `docx2pdf.py` | `.doc` / `.docx` → `.pdf`（需要 Windows + Microsoft Word） |
| `pptx2md.py` | `.pptx` → `.md` |
| `fix_punct.py` | `.md` 标点、计量单位与上下标后处理 |
| `normalize_docx.py` | `.docx` 字体、字号、行距、标题与多级编号标准化 |
| `pdf2md.py` | `.pdf` → `.md` / `.json`（本地 OXR-1.0 OCR） |
| `mdviewer.html` | 独立的 Markdown 查看器 |

## 安装

```bash
pip install -r requirements.txt
```

要求 Python 3.8 及以上（开发与验证环境为 3.13）。除 `pdf2md.py` 外的脚本都是单文件，可以单独拷走使用。

- `docx2pdf.py` 需要本机安装 Microsoft Word；`pywin32` 已按平台条件写入依赖。
- `pptx2md.py` 需要 `python-pptx`；把 EMF/WMF 图片转成 PNG 需要 Pillow，缺了只会跳过这类图片。

`pdf2md.py` 还需要同目录的 `OXR/` 子目录和约 3 GB 的模型权重。权重不入库，需要自行下载：

```bash
pip install modelscope
modelscope download SoMark/OXR-1.0 --local_dir ./OXR/models/OXR-1.0

# 或者
huggingface-cli download SoMarkAI/OXR-1.0 --local-dir ./OXR/models/OXR-1.0
```

推理默认使用 CUDA；没有 GPU 时可以加 `--device cpu`，但会慢一个数量级。

## 便携版

`python_embeded` 目录包含无 GPU 加速的 Python 与 PyTorch 运行环境，可用于没有 Python 环境的机器。把下面命令里的 `python` 替换成 `.\python_embeded\python.exe` 即可。使用 `pdf2md.py` 前建议先安装 GPU 加速的 torch/torchvision。

## 用法

### docx → md

```bash
python docx2md.py path/to/file.docx      # 转换单个文件
python docx2md.py -d path/to/dir         # 转换目录下所有 .docx
```

结果写在与 docx 同一目录、文件名相同的 `.md` 文件，图片导出到同目录的 `<md名>_media/`。目录模式下会打印 `完成: 成功 N / 共 N 个文件`，任一文件失败则退出码为 1。

### md → docx

```bash
python md2docx.py path/to/file.md        # 转换单个文件
python md2docx.py -d path/to/dir         # 把 path 当作目录，转换其中所有 .md
python md2docx.py -n path/to/file.md     # 标题套用多级自动编号
```

结果写在与 md 同一目录、文件名相同的 `.docx` 文件。

`-n`（`--heading-numbering`）会给标题样式挂上 Word 多级自动编号，并去掉标题文字里已有的序号。`###### 图…` / `###### 表…` 分别写成 Heading 8 / Heading 9，图号、表号与 `normalize_docx.py` 一致；普通 `######` 仍是 Heading 6，不参与自动编号。

支持 GFM 标题、列表、表格、引用、代码、图片、链接，行内 HTML 上下标/下划线，以及 `$…$` / `$$…$$` 公式；细节见 `AGENTS.md`。

### docx → pdf

```bash
python docx2pdf.py                  # 转换当前工作目录
python docx2pdf.py path/to/dir      # 转换指定目录
python docx2pdf.py path/to/dir -r   # 递归处理子目录
python docx2pdf.py path/to/dir -f   # 覆盖已存在的 PDF
```

结果写在与源文件同目录、文件名相同的 `.pdf`。支持 `.doc` / `.docx` / `.docm` / `.dot` / `.dotx`，会自动跳过 Word 临时文件，默认不覆盖已有 PDF。需要本机安装 Microsoft Word。

### pptx → md

```bash
python pptx2md.py path/to/file.pptx          # 转换单个文件
python pptx2md.py -d path/to/dir             # 转换目录下所有 .pptx / .pptm
python pptx2md.py file.pptx --no-notes       # 不输出演讲者备注
python pptx2md.py file.pptx --slide-markers  # 每页前加 <!-- slide N -->
```

结果写在与 pptx 同一目录、文件名相同的 `.md` 文件，图片导出到同目录的 `<md名>_media/`。

### pdf → md（本地 OCR）

```bash
python pdf2md.py path/to/file.pdf             # 转换单个文件
python pdf2md.py -d path/to/dir               # 转换目录下所有 .pdf
python pdf2md.py -d path/to/dir --recursive   # 递归处理子目录
python pdf2md.py -d path/to/dir -o path/to/out --formats md
```

结果默认写在与 PDF 同一目录、文件名相同的 `.md` / `.json`；用 `-o` 指定输出目录时会保留相对路径结构。插图导出到同目录的 `<文件名>_media/`，Markdown / JSON 里写相对路径。

常用参数：

| 参数 | 作用 |
|---|---|
| `-o`, `--out` | 输出目录（默认写在 PDF 旁边） |
| `--recursive` | 配合 `-d` 递归子目录 |
| `--formats md,json` | 输出格式，逗号分隔；默认只输出 `.md` |
| `--pages 1-5,8` | 只处理指定页（1 起） |
| `--skip-existing` | 输出已存在就跳过 |
| `--dpi 150` | PDF 渲染 DPI |
| `--keep-header-footer` | 保留页眉页脚（默认丢弃） |
| `--device auto/cuda/cpu` | 推理设备，默认有 CUDA 就用 CUDA |
| `--max-tokens 8192` | 单个版面块的最大生成 token 数 |
| `--batch-size 16` | 文本模型合批的批大小；1 表示不合批 |
| `--reading-order llm/geo` | 阅读序方式；`llm` 质量最高，`geo` 更快但对复杂版式有局限 |
| `--chunk-pages 8` | 每批送进流水线的页数 |
| `--layout-batch-size 8` | 版面分析批大小 |
| `--layout-concurrency 2` | 版面分析在途批数 |
| `--torch-threads 0` | torch/OpenCV 的 CPU 线程数，0 表示默认 |
| `--graph-decode auto` | 是否把逐 token 解码捕获成 CUDA Graph，`auto`/`on`/`off` |
| `--graph-cache-len 2048` | 图解码预留的 KV 缓存长度 |
| `--model <目录>` | 换一份模型权重 |
| `--verbose` | 打印调试日志 |

### md 标点修正

```bash
python fix_punct.py 文件1.md [文件2.md ...]
```

原地重写文件，保留原行尾，幂等；非 UTF-8 编码的文件会先转成 UTF-8 再处理。只处理以中文为主的行，公式、代码、URL 与链接地址保持原样。

### docx 格式标准化

```bash
python normalize_docx.py path/to/file.docx
```

原地重写文件，统一字体、字号、行距、标题格式与多级编号：

- 中文标题黑体、正文宋体，英文 Times New Roman；
- 正文小四、1.5 倍行距；
- 标题 1 小三、段前分页；标题 2~9 小四；
- 标题 1~9 加粗、无倾斜、段前段后为 0，标题 8/9 居中；
- 标题 1~7 是章节号，8 级是图号、9 级是表号，图/表号随二级章节重置；
- 自动替换常见计量单位写法，如 `m3/h` → `m³/h`。

### Markdown 查看器

用浏览器直接打开 `mdviewer.html`。点「打开文件」或把 `.md` / `.smd` / `.txt` 文件拖进窗口打开；拖入文件夹不支持。

预览按源文件逐行成段：没有被空行分隔的连续行也各自显示为一个段落。本地相对路径图片需要在底部提示条点一次「授权打开所在文件夹…」，选择 md 所在目录后即可显示。

## 已知限制

- `docx2md.py` 只处理正文，不提取页眉页脚、脚注尾注与书签；图片只导出正文里的图片。
- 含合并单元格的表格会整体输出为 HTML（普通表格仍是 GFM），嵌套表格会被拍平。
- `md2docx.py` 使用 python-docx 默认页面设置，不继承原 docx 的版式；图片按行内方式插入。
- `md2docx.py` 的公式转换只覆盖 `docx2md.py` 输出的 LaTeX 子集，解析不了时退回 LaTeX 原文。
- `pptx2md.py` 只读 `.pptx` / `.pptm`，不保留版式、动画与母版；OLE 对象只导出预览图。
- `pdf2md.py` 把每页重新 OCR，不读取 PDF 文本层；插图是位图裁切，图表数据不会变成表格。
- `pdf2md.py` 依赖上游 `OXR/` 子目录，`mdviewer.html` 依赖上游 `lib/` 子目录；许可证见各自的 `LICENSE`。

## 开发

实现细节、往返约定与常见坑见 `AGENTS.md`。
