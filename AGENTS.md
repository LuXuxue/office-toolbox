# AGENTS.md

给在本仓库工作的 AI 编码助手看的项目约定与验证流程。人读的说明见 `README.md`。

## 仓库里有什么

```
docx2md.py                .docx → .md
md2docx.py                .md → .docx
docx2pdf.py               .doc / .docx → .pdf
pptx2md.py                .pptx → .md
pdf2md.py                 .pdf → .md / .json（本地 OXR 推理，依赖 OXR/）
fix_punct.py              .md 标点修正（英文标点 → 中文标点；附带计量单位等自动替换）
normalize_docx.py         .docx 格式标准化（主题字体、字号行距、标题多级列表）
mdviewer.html             独立的 Markdown 查看器（浏览器打开即用，按钮选文件/拖入 .md，配色跟随系统，源文件里没有空行分隔的连续行也逐行成段，相对图需在提示条授权所在文件夹一次）
mdviewer_update_smd.bat   mdviewer.html 的后端渲染组件更新程序
requirements.txt
lib/                      mdviewer.html 的后端渲染组件资源库
OXR/
  oxr/                        上游流水线最小子集（版面结果 → md/json）
  services/layout_analyze/    D-FINE 版面分析推理代码
  models/OXR-1.0/             模型权重（gitignore，自行下载）
```

`docx2md.py` 与 `md2docx.py` 是**单向转换器**，`md2docx.py` 的公式转换器专门服务于 `docx2md.py` 的 LaTeX 输出。改动一侧的输出格式，另一侧必须同步跟上，否则往返会静默损坏。行内格式的对应关系是加粗/斜体/删除线/下划线/上标/下标 ↔ `**`/`*`/`~~`/`<u>`/`<sup>`/`<sub>`，两侧都要认。`md2docx.py` 除 GFM 表格外还会解析 `pdf2md.py` 输出的 HTML 表格（`table`/`tr`/`td`/`th`、`rowspan`/`colspan`、单元格里的 `br`/`sup`/`sub`），并支持独占两行的 `$$` 围栏公式；`docx2md.py` 遇到含合并单元格的表格时也用 HTML 写出来，普通表格仍用 GFM，保证 `rowspan`/`colspan` 往返不丢。`docx2pdf.py` 只负责调用本机 Microsoft Word 导出 PDF，不参与 Markdown 往返。`pptx2md.py` 只做 pptx → md，目前没有反向脚本，不改动上面这对往返。`fix_punct.py` 做 md 的标点修正（中文文本里夹杂的英文标点替换为中文标点，含英文括号 `()` → `（）`，公式、代码、URL、英文段落不动）并附带计量单位等自动替换（如 m3/h → m³/h，只作用于未受保护的区域；非 UTF-8 编码的 md 先自动转 UTF-8 再处理，原来就是则跳过），是单向后处理，不参与 Markdown 往返。`normalize_docx.py` 只做 docx 的格式标准化（字体：标题黑体/正文宋体、英文 Times New Roman；正文小四 1.5 倍行距、标题 1 小三、标题 2~9 小四；标题 1~9 加粗、无倾斜、段前段后为 0、标题 8/9 居中；多级列表，其中 8/9 级为图表编号；无间隔样式五号 1 倍行距；自动替换，如 m3/h → m³/h），是单向后处理，不参与 Markdown 往返。

`pdf2md.py` 是 pdf → md/json 的单向转换器（md 之后可以交给 `md2docx.py`）。它自己只负责 CLI、PDF 渲染、本地模型加载和文件输出，版面分析之后的流水线全部复用 `OXR/oxr/pipeline` 里的上游代码。`OXR/` 是从 [SoMarkAI/OXR](https://github.com/SoMarkAI/OXR) 复制的最小运行子集：`oxr/`（去掉 `cli/`、`server/`）加 `services/layout_analyze/dfine_la`，许可证见 `OXR/LICENSE`、`OXR/NOTICE`。模型权重 `OXR/models/OXR-1.0` 不在库里，用 modelscope 或 huggingface-cli 下载。

`mdviewer.html` 使用的渲染程序来自于 [SoMarkAI/SoMarkDownViewer](https://github.com/SoMarkAI/SoMarkDownViewer)，这里只保留最小需要的资源子集：`lib/`（去掉编辑功能）。许可证见 `lib/LICENSE`。本地相对路径图片通过渲染后把 `<img>` 换成 object URL 来显示，详见「环境陷阱」第 17 条。

## 硬性约束

1. 注释用中文。
2. **`OXR/` 和 `lib/` 是上游代码子集，能不改就不改**。本地推理的接线（替换 `model_client` / `layout_client`、prompt、生成参数）都在 `pdf2md.py` 里；确实要改 `OXR/` 时在提交说明里写清改了什么、为什么，因为升级上游时会整块覆盖。
3. 不要提交模型权重、`OXR/models/` 与转换产物（`*_media/`、`outputs/`）。

## 环境陷阱

这几个坑都实际踩过，会浪费大量时间：

### 1. 行尾必须是 CRLF

仓库 `core.autocrlf=true`，所有 Python 源文件（含 `pdf2md.py`）都是 CRLF。**任何用脚本重写源文件的操作都必须用 `newline=''`**，否则整个文件会被改成 LF，`git diff` 爆炸成上千行。

改完检查：

```bash
python -c "d=open('docx2md.py','rb').read(); print(d.count(b'\r\n'), d.count(b'\n'))"
```

两个数必须相等。

### 2. 不要用 PowerShell here-string 写测试数据

`@"..."@` 会做 `$VAR` 插值，而 `\$` 在 PowerShell 里不是转义序列——写 LaTeX 测试用例必炸。**用 `write` 工具创建测试脚本文件**，不要 here-string。

### 3. 控制台是 GBK

中文输出会显示成乱码，`×`(U+00D7) 会显示成 `*`（曾导致过保真度检查的误报）。**验证特殊字符要对比码点，不要看控制台输出**：

```python
print([hex(ord(c)) for c in s])
```

### 4. OMML 的命名空间

公式元素在 `m` 命名空间（`http://schemas.openxmlformats.org/officeDocument/2006/math`），**属性也在 `m` 上，不是 `w`**。`docx2md.py` 里用 `_m()` / `_mtag()`，`md2docx.py` 里用 `_mtag()`。用 `qn('w:val')` 读 OMML 属性会永远拿到 `None`。

### 5. Word 的标题样式名不可信

中文模板的标题样式常叫「1级标题」「2级标题」，不是内置的 `Heading N`/`标题 N`。**认标题要看 `w:outlineLvl`（大纲级别）**，那才是 Word 导航窗格与目录使用的依据；样式名只能当兜底。`w:outlineLvl val="9"` 表示「正文」，写了它就不再是标题（段落级与样式级都算数）。

同理，编号的 `numFmt` 也不等于显示效果：`w:isLgl`（legal 编号）会让本级 `lvlText` 里的序号一律显示为阿拉伯数字，一级是中文「二.」时二级仍显示「2.1.」。

### 6. 一个 run 不等于一个字

Word 会把连续文字拆成多个 run（拼写检查、格式切换、编辑痕迹），逐个 run 加 `**` 会产生 `**前半****后半**`。要按「格式相同的相邻片段」合并后再加标记。

### 7. 链接要写在最外层

超链接整体算一个片段，链接文字的加粗写在括号里面：`[**文字**](url)`。写成 `**[文字](url)**` 的话，`md2docx` 的 `**…**` 分支会先匹配上，把地址当正文写进文档。

### 8. 表格的行要排除嵌套表格

`w:tbl` 里允许嵌套 `w:tbl`。直接用 `.iter(qn('w:tr'))` 会把内层表格的行也遍历出来，而 `_tc_text()` 已经把内层文字算进父单元格，最终同一段文字出现两次；但也不能简单用 `findall` 只取直接子元素，因为行可能被包在 `w:sdt` 里。统一走 `_table_rows()`（按最近的 `w:tbl` 祖先判断）。`_extract_table_grid()` 与 `_table_header_flags()` 的行数必须一致，否则表头标记会因长度对不上而被静默忽略。

### 9. 块级公式的 `$$…$$` 要整段匹配

`docx2md.py` 对 `m:oMathPara`（块级公式）输出 `$$…$$`，对 `m:oMath`（行内公式）输出 `$…$`。`md2docx.py` 里如果先按 `$…$` 切分，`$$x+1$$` 会被拆成「`$`」+「`$x+1$`」+「`$`」——表格单元格里的块级公式就变成一个字面量 `$` 加一个行内公式。所以切分时 `$$…$$` 必须排在 `$…$` 前面。

独占一行的 `$$…$$` 走 `display_match` 分支，表格单元格里的 `$$…$$` 只能靠 `_add_formatted_text()` 的切分顺序兜住。改这两处都要一起改。

### 10. `w:pPr` 的子元素有顺序要求

OOXML 规定 `w:pPr` 里 `w:pStyle` 必须排在 `w:numPr` 前面。`pPr.insert(0, numPr)` 会把 `w:numPr` 插到最前面，写出的文档虽然多数时候能打开，但顺序不合规，Word 重新保存时会重排。统一用 `pPr.get_or_add_numPr()`（python-docx 自己会插到正确的位置），不要手写 `insert`。

### 11. 超链接有好几种写法，只认 `w:hyperlink` 会丢链接

Word 里的链接地址不一定在 `w:t` 旁边。抽文字只能拿到链接文字，地址会整个丢掉：

- `w:hyperlink` 带 `r:id` — 外部链接，地址在 `doc.part.rels` 的关系目标里；
- `w:hyperlink` 只有 `w:anchor`（或 `w:docLocation`）— 文档内部的书签链接，没有 `r:id`，要转成 `#书签`；
- `w:fldSimple` 与复杂域（`w:fldChar` + `w:instrText`）— `HYPERLINK "地址"` 域，地址只在域代码里；
- 表格单元格里的链接 — 单元格走 `_tc_text()`，不共用段落的行内渲染，链接要在那里单独处理；
- 文本框（`w:drawing` / `w:pict` 里的 `w:txbxContent`）— 里面的段落要单独渲染，否则里面的链接只是文字。

`mc:AlternateContent` 里 `mc:Choice` 与 `mc:Fallback` 是同一份内容的两种写法，只取一个，否则同一段文字会被算两遍。

### 12. 写超链接要用 `w:hyperlink` + 关系，不能只染颜色

`md2docx.py` 里把 `[文字](地址)` 写成「蓝色带下划线的普通 run」看着像链接，其实不可点击，`docx2md` 再读回来时也拿不到地址。真正的超链接长这样：

- 外部地址 — `<w:hyperlink r:id="rIdN">`，`rIdN` 来自 `para.part.relate_to(url, RT.HYPERLINK, is_external=True)`；
- 文档内部的书签 — `<w:hyperlink w:anchor="书签名">`，**不建关系**（内部链接没有外部地址）。

两条坑：

- 链接文字的 run 在 `w:hyperlink` 元素里，**不是** `w:p` 的直属子元素。`para.runs` 只返回直属的 `w:r`，遍历它会漏掉链接文字——表格表头加粗那一步就曾因此漏加。统一走 `_iter_runs()`（用 `para._p.iter(qn('w:r'))`）。
- `w:hyperlink` 必须插在链接文字原来的位置（`run._r.addprevious(hl)` 后再把 run 搬进去）。直接 `append` 到段尾会打乱顺序——链接后面跟着 `m:oMath` 时公式会跑到链接前面。

链接文字套 `Hyperlink` 字符样式（Word 内置的蓝色带下划线样式，python-docx 的默认模板里没有，要自己建），同时保留直接格式兜底。

### 13. 相邻表格只能按表头完全一致来合并

`merge_consecutive_tables()` 只服务于「跨页被拆成两个 `w:tbl` 的同一张表」，判据必须是表头完全一致。报告里的表单表（如 3.3 管段检测成果表）每张都以「录像文件 / 起始井号 / 终止井号」开头，字符集几乎相同，若按字符相似度合并，就会把后一张表的首行当成重复表头丢掉。表头不同就各自成表，宁可多一张表也不能丢行。

### 14. `pdf2md.py` 本地推理的三个坑

这三处都是踩过的，改动 `pdf2md.py` 的模型加载/生成部分时别退回去：

- **`eos_token_id` 必须显式传 `[151643, 151645]`**。模型内层 `llm.generation_config` 里只有 `151643`，而模型实际用 `<|im_end|>`（`151645`）结束；漏掉它就会一路生成到 `--max-tokens`（8192 起步），看起来像卡死。两个值都写在 `OXR/models/OXR-1.0/generation_config.json` 里，代码用 `GenerationConfig.from_pretrained()` 读出后作为 `eos_token_id` 传给 `generate`，别只取内层 `llm` 的那一个、也别硬编码。
- **`indicator_token_indices` 要重建**。用 `from_pretrained` 低内存加载时它是非持久化 buffer，checkpoint 不覆盖，可能是未初始化内存（任何随机值都可能），接着触发 CUDA `device-side assert`（`srcIndex < srcSelectDimSize`）。加载后按 `visual_vocab_size` 与 `INDICATOR_IDS` 长度重建。
- **不要用 `trust_remote_code`**。transformers 5.x 会把 `modeling_oxr.py` / `configuration_oxr.py` 复制到不同缓存子目录，`Siglip2NavitConfig` 注册失效，报 `Unrecognized configuration class`。现在的做法是把模型目录挂成临时包 `_oxr_local_model` 再导入，保证所有类来自同一份源码。

### 15. 模型权重不在仓库里

`OXR/models/` 已写进 `.gitignore`（约 3 GB）。换机器时重新下载：

```bash
pip install modelscope
modelscope download SoMark/OXR-1.0 --local_dir ./OXR/models/OXR-1.0
```

`pdf2md.py` 校验 `OXR/models/OXR-1.0/config.json` 与 `layout_analysis/last.pth`，缺了会直接退出并打印这条命令。

### 16. `pdf2md.py` 的提速结构不能被“顺手简化”掉

现有的速度来自几处互相咬合的改动，单独拿掉任一处都会退回原来的数倍耗时：

- **合批的 key 是 `(kind, options)`，不含 `limit`**。阅读序的提示词里带本页 bbox、token 上限也随本页块数变，按 prompt 或 limit 分桶会让每页各自成批，每批都要按 16 行补齐做预填充，固定开销摊不薄。阅读序必须传 `kind="reading_order"`，长短差异由 `run_batch` 里的按行 `row_limit` 截断兜住。别把 `limit` 加回 key。
- **图解码器按「批大小 × 缓存长度」分桶复用**（`GreedyGraphDecoderPool`）。单步解码要读完整段预留缓存，预留越多单步越慢；所以 1 行请求不能塞进 16 行的图（批太小 → 批数多、总预填充开销大；批太大 → 没有净收益）。
- **池子的字节预算只能给 3 GiB**。池子里装的是已经捕获的图；显存一紧就开始抖动，比重新捕获几张图贵得多。
- **`--chunk-pages` 默认 8 是有原因的**。显存峰值随它线性上涨，一旦逼近 12 GB 就开始抖动。
- **分段解码在本机不划算**。按步数切段、把已结束的行剔出去续跑，每次续跑都要换一张图并重做预填充，省下的步数全被吃掉。只在缓存装不下时才续跑。
- **PDF 按 chunk 懒渲染，版面预取下一块**。渲染与上一块的推理重叠，`_layout_prefetched` 用 `id(images)` 作键，所以「渲染 + 版面预取」返回的列表对象必须原样传给流水线，别中途再复制一份。
- **视觉塔跑独立 CUDA 流，与下一批的解码重叠**（`VISION_OVERLAP`）。三分段流水线：`_prepare_loop`（CPU 预处理）→ `_vision_loop`（视觉塔，独立流 + event）→ `_generate_loop`（预填充 + 解码，默认流 `wait_event`）。视觉塔是算力受限、解码是带宽受限，两者能互相填空。几条不能动的约束：
  - 解码热路径的 `torch.cuda.synchronize(device)` 必须用 `current_stream().synchronize()`，设备级同步会把重叠整个抹掉（这是收益的大头）。
  - 视觉提交与 `torch.cuda.graph()` 捕获共用 `_CUDA_CAPTURE_LOCK`：捕获期间不能让另一条流提交 kernel，但只锁“提交”、不锁等待，否则重叠同样消失。
  - 嵌入在视觉流上分配、在默认流上使用，`run_batch` 里要 `inputs_embeds.record_stream(...)`，否则视觉流的下一批可能提前复用这块显存。
  - **`VISION_CHUNK_PATCHES` 必须切块，且上限 ≤ 8192**。两条流各有各的显存池（显存块只能被同一条流复用），整批整页图一次前向就会把显存顶到 12 GB 卡的边界、触发驱动换页，不是慢一点而是慢几倍。切块不改变视觉总计算量。
  - 8192 与 6144 一样稳，切得少，前向的 bf16 累加顺序更接近整批——这是既有数值特性，不是排版错误。
- **裁切去重缓存（`_crop_cache`）只缓存识别出的正文与格式，且每次 `process_one` 清空**。它按「裁切像素 + 块类型」判重，用来让重复的页眉/页脚/页码只识别一次。两条约束都不能松：
  - 不能把整个 `Block` 存下来复用。`Block` 里的 `bbox` / `text_before` / `text_after` 属于被命中的那个块，复用时会把这套上下文带到别的块上 —— 同进程里对同一份 PDF 跑第二遍，输出会多出一个重复的标题块。
  - 不能跨文件/跨次保留。缓存挂在 `LocalEngine` 上，第二次处理同一份 PDF 时所有裁切都会命中，识别请求一个都不发，模型参数一改也不会失效——这不再是提速，而是拿上一次的结果顶替本次推理。
- **合批要按「预估输出长度」挑，不能让短请求陪跑**（`_length_hint` / `_best_length_group`）。一批里的行同步推进，解码步数等于最长那条的输出长度，短请求全程陪跑。hint 用裁切像素面积估识别长度、用提示词长度估阅读序长度，按 hint 排序后用 DP 把序列切成若干**连续**的组（每组代价 = 组内最大 hint，组大小 ∈ [`batch_size`/2, `batch_size`]），每次只交出一组。两条约束：
  - hint 只能用于排序，不能进 `key`。进了 key 就会把同类拆成许多小批，预填充反而更贵。
  - 「攒批」要等到队列这一波到齐：一个 chunk 的识别请求几乎同时到达，判据是「队列连着 30 ms 没再增长」或够两批。只看是否凑满一批的话，第一批永远是从最先到的 16 条里硬凑，DP 没有选择余地。
- **generate 兜底只续跑没完成的行，退化行要逐行掐断**。图解码装不下预留缓存时不能把整批退回 `generate` 重跑：`DynamicCache` 随**实际**生成长度增长，批 16、生成到 8192 时约 16.8 GB，16 GB 卡直接顶满并触发驱动换页，表现就是「卡住、显存几乎全满」（GB 50268 第 4 页单页实测 424s → 18s，同批 8 页 368s → 27s）。三条约束：
  - `_graph_decode_rows` 返回 `(tokens, residual, degenerate_rows)`。`residual` 只含没跑完的行（它们的 `embeds` 已经拼好已生成前缀），由 `_generate_rows` 按 `GENERATE_KV_BUDGET_BYTES`（3 GiB）拆组，单组最坏占用不超过预算；已经结束的行不重跑。
  - 退化检测（模型陷入「1、2、3…」数数或复读）用 `_DegenerateLogitsProcessor`，**不能用 `StoppingCriteria`**：停止条件只能停整批，同批还有正常行时退化行照样陪跑到 `max_new_tokens`。处理器把退化行的下一步强制成 eos，其余行不受影响；退化行标成 `finish_reason="length"`，交给 recognition 既有的 repetition_penalty 恢复重跑。
  - 图解码里的同一套检测直接把该行的 token 在检测点截断（`stop_at`），避免后续重放再往 `history` 里写重复内容。判据（`_is_degenerate_text`）有两条，都不能放宽，否则会误伤正常表格与编号：①末尾重复同一个片段、周期 16~400 字不定（**不能写死周期**：GB 50268 第 148 页的表格复读的是 29 字句子，16/24/32… 完全抓不到），且重复区域连起来至少 `DEGENERATE_REPEAT_SPAN`=256 字，纯标点片段（目录点线）跳过；②尾部 50 个数字连续递增/相同，且数字之间只有标点分隔符（纯空格分隔的数字串是阅读序的正常输出 `0 1 2 3 …`）。

### 17. `mdviewer.html` 里本地相对图的三个坑

`mdviewer.html` 直接在浏览器里跑，本地图片受文件系统访问权限限制，这几个点都实际踩过：

- **图片 destination 带空格时不发 `<img>`**。SoMarkDown（markdown-it）按 CommonMark 解析，`![](./名字 带空格_media/a.png)` 会被当普通文字输出，图片整个不显示。`encodeImageDestinations()` 在渲染前把空格统一编码成 `%20`（跳过代码块、已是 `<...>` 的 destination 以及可选的 `"title"`），`render()` 传它的返回值。pdf2md/docx2md 输出目录名几乎都带空格（中文 + 空格），别把这个预处理拿掉。
- **浏览器拿不到被打开本地文件的绝对路径**，相对图无法解析到 md 旁边的目录，只能在渲染后把 `<img src>` 换成 object URL：先建「相对路径 → File」索引，再读取并 `createObjectURL`。`img` 原始 src 存在 `dataset.origSrc` 里，重复解析时用它而不是 `img.currentSrc`（否则上次换的 blob: 会被当原始地址）。按 md 目录拼完还找不到时退化按文件名匹配。
- **file:// 页面 fetch(file://) 被 Chromium 禁止**。`?file=` 指定本地 md 会加载失败，只能在失败后弹提示条，等用户手势点「授权并打开…」选一次 md 所在文件夹再读取——不弹出来就永远打不开，别改回纯直载。另外：drop 删除了文件夹遍历，只收 `.md`/`.smd`/`.txt` 文件本身，文件夹拖放会给出明确提示；之前遍历目录树的实现有边界情况（md 不在文件夹根层或 `webkitGetAsEntry` 返回 null 时会静默打不开），别把它加回去。

### 18. `md2docx.py` 的 HTML 表格有两个坑

- **合并单元格要在写内容之前做完**。`python-docx` 的 `cell.merge()` 会把被合并格里的段落一起带进来；先写内容再合并容易留下重复段落。现在的顺序是：建空表 → 按 `rowspan`/`colspan` 合并 → 只往合并后的左上角单元格写内容。
- **判断「哪些 run 是这次新增的」不能存 `id()`**。lxml 的元素代理对象被回收后 `id` 会复用，`{id(r) for r in para._p.iter(qn('w:r'))}` 会把旧 run 也当成新 run，导致 `<sup>` 的格式扩散到前文、随后又被 `<sub>` 覆盖成下标。必须 `set(para._p.iter(qn('w:r')))` 直接存元素本身。

### 19. `docx2md.py` 的合并表格不能复用 `_extract_table_grid()`

`_extract_table_grid()` 会把 `vMerge` 的续行重复首行文字、把 `gridSpan` 拆成空格子，用来生成 GFM 够用，但 `rowspan`/`colspan` 的原始信息已经丢了。写 HTML 表格时直接遍历 `w:tc`、按 `gridSpan`/`vMerge` 输出属性；表头只认 `w:tblHeader`，不要按拍平后的网格猜（合并表头覆盖的列不全，猜出来会错）。普通无合并表格仍走原来的 GFM 路径。

### 20. 行内标记的嵌套顺序要跟 `md2docx.py` 的解析顺序对齐

`md2docx.py` 的行内解析是**分层下降**的：`_add_formatted_text`（公式）→ `_add_html_segment`（HTML 标签）→ `_add_text_segment`（代码段/图片/链接）→ `_add_formatted_runs`（`~~`/`***`/`**`/`*`）。外层先被切走，里层才有机会解析，所以 `docx2md.py` 输出的标记顺序必须是：

- **HTML 标签最外层**：`<sup>**2**</sup>`、`<u>~~作废~~</u>`。写成 `**<sup>2</sup>**` 时 `_add_html_segment` 会先把 `<sup>` 切走，剩下的 `**` 片段匹配不到成对星号，会被当正文写进 Word。
- **`~~` 在 `**`/`*` 之外**：`~~***文字***~~`。写成 `***~~文字~~***` 时 `bi` 正则的 `***` 配对会落在错误的位置，内文变成 `文字~~`，同样写进正文。
- `_segments_to_markdown()` 的叠加顺序（斜体 → 加粗 → 删除线 → `<u>` → `<sup>`/`<sub>`）就是据此定的，别「顺手」调换。
- `_add_marked_runs()` 让 `**…**`、`*…*`、`~~…~~` 里的内文再走一遍 `_add_formatted_text()`，所以标签/公式嵌在粗体里面也能解析；不要退回 `para.add_run(内文)`。

两个配套点：超链接里的 run 不输出 `<u>`（Hyperlink 样式自带下划线，见 `_in_hyperlink()`）；单元格走独立的 `_collect_cell_run_text()`，行内格式的改动必须在这条路径上也做一遍，否则表格里的 `m²` 会退回普通文字。

### 21. 空的分节符段落不占编号

用「分节符」结束封面/目录那一节时，Word 把 `w:sectPr` 放进最后一个段落的 `w:pPr`，那个段落标记常带着「标题 1」样式（继承编号）。它没有文字，**Word 不把它算进编号序列**：正文第一个一级标题仍是「1」。`docx2md.py` 原来照常推进计数器，整篇标题就整体后移一位（`1 概述` 变成 `2 概述`）。判据见 `NumberingResolver._is_empty_section_break()`：带 `w:sectPr` 且没有可见内容（文字/公式/图片/对象）才跳过。

这条规则是拿 Word 实测出来的，别按「空段落一律不占号」改：不带分节符的空标题段落 Word 照样占号（`w:start` 也从 0 开始计数，别用 `or 1` 把 0 吞掉）。

## 改完自检

- [ ] 行尾仍是 CRLF
- [ ] `docx2md.py` 与 `md2docx.py` 同步改：`docx2md.py` 新写出的标记，`md2docx.py` 要能解析回 Word；`md2docx.py` 新写进 docx 的格式（如 HTML 表格合并、`<sup>`/`<sub>`/`<u>`），`docx2md.py` 也要能还原成同样的标记
- [ ] 注释是中文
- [ ] `git status` 里没有模型权重、`OXR/models/`、`*_media/`、`outputs/`
