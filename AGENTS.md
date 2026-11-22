# AGENTS.md

给在本仓库工作的 AI 编码助手看的项目约定与验证流程。人读的使用说明见 `README.md`。

## 仓库里有什么

```
docx2md.py                .docx → .md
md2docx.py                .md → .docx
docx2pdf.py               .doc / .docx → .pdf
pptx2md.py                .pptx → .md
pdf2md.py                 .pdf → .md / .json（本地 OXR 推理，依赖 OXR/）
pdfcompress.py            .pdf → .pdf（按 DPI 重编码图像压缩；可选整页 PNG 化）
image-moss.py             调本地 ComfyUI 出图（PNG；依赖 ComfyUI 与模型权重）
fix_punct.py              .md 标点修正（英文标点 → 中文标点；附带计量单位、上下标后处理）
normalize_docx.py         .docx 格式标准化（主题字体、字号行距、标题多级列表）
mdviewer.html             独立的 Markdown 查看器
mdviewer_update_smd.bat   mdviewer.html 的后端渲染组件更新程序
requirements.txt
pngquant/                 pdfcompress.py 用的 pngquant 可执行文件
lib/                      mdviewer.html 的后端渲染组件资源库
OXR/
  oxr/                        上游流水线最小子集（版面结果 → md/json）
  services/layout_analyze/    D-FINE 版面分析推理代码
  models/OXR-1.0/             模型权重（gitignore，自行下载）
```

## 硬性约束

1. 注释用中文。
2. **`OXR/` 和 `lib/` 是上游代码子集，能不改就不改**。本地推理的接线（替换 `model_client` / `layout_client`、prompt、生成参数）都在 `pdf2md.py` 里；确实要改 `OXR/` 时在提交说明里写清改了什么、为什么，因为升级上游时会整块覆盖。
3. 不要提交模型权重、`OXR/models/` 与转换产物（`*_media/`、`outputs/`）。

## 转换器约定

`docx2md.py` 与 `md2docx.py` 是**单向转换器**，改动一侧的输出格式，另一侧必须同步跟上，否则往返会静默损坏。

- **行内格式** — 加粗/斜体/删除线/下划线/上标/下标 ↔ `**`/`*`/`~~`/`<u>`/`<sup>`/`<sub>`；相邻同格式 run 要合并；HTML 标签放最外层，`~~` 在 `**`/`*` 之外；超链接整体写成 `[文字](地址)`，链接文字的加粗写在括号里。
- **字面记号** — 正文里原样出现的 `\`、`` ` ``、`*`、`_`、`~`、`[`、`]` 必须写成 `\X`（否则 `1*2*3=6` 回转时会被当成斜体，星号消失）；`md2docx.py` 解析格式前把 `\X` 换成私用区占位符，写进 Word 时再还原，代码段里的反斜杠按字面量保留。竖线由表格渲染单独转义，`$` 与 `<` 走公式/HTML 标签的解析路径。
- **标题** — `docx2md.py` 优先按 `w:outlineLvl` 认标题，6～9 级统一压成 `######`。`md2docx.py` 回写时 `###### 图…` → Heading 8、`###### 表…` → Heading 9，其余 `######` → Heading 6；普通 6/7 级不参与自动编号。
- **图表编号** — 加 `-n` 时 Heading 8/9 的编号必须与 `normalize_docx.py` 一致：`图 %1.%2-%8`、`表 %1.%2-%9`，带 `w:lvlRestart=2`；写入前先剥掉标题文字里已有的 `图 1.1-1` / `表 1.1-1`，否则往返会重复编号。
- **表格** — 普通表格用 GFM；含 `rowspan`/`colspan` 的表格用 HTML。`docx2md.py` 合并跨页表格时只认表头完全一致；`md2docx.py` 写 HTML 表格时先合并单元格再写内容。
- **公式** — OMML 元素在 `m` 命名空间，属性也是 `m`；`md2docx.py` 切分时 `$$…$$` 必须排在 `$…$` 前面。
- **行内 HTML** — `<sup>`/`<sub>`/`<u>`/`<b>`/`<i>`/`<br>` 要能写回 Word；判断新增 run 不能存 `id()`，要存元素本身。
- **其他转换器** — `fix_punct.py` 只做 md 后处理，公式/代码/URL/链接地址不替换，非 UTF-8 先转 UTF-8；`pptx2md.py` 只做 pptx → md；`normalize_docx.py` 只在样式层统一格式并套用 1~9 级编号；`pdf2md.py` 的本地推理接线都在自己文件里，`OXR/` 尽量不改。

## 图像生成

`image-moss.py` 不启动 ComfyUI 服务：把 ComfyUI 目录插进 `sys.path`，直接调用 ComfyUI 自带的节点类（`UNETLoader` / `CLIPLoader` / `VAELoader` / `CLIPTextEncode` / `SkipLayerGuidanceDiTSimple` / `EmptyFlux2LatentImage` / `KSampler` / `VAEDecode`）依次执行，存图用 Pillow。**不要引入第三方节点或 ComfyUI 私有接口**，否则 ComfyUI 升级就没法用了。

- `--comfyui` 默认 `E:\ComfyUI`（可用环境变量 `COMFYUI_PATH` 覆盖），目录下必须有 `nodes.py`。
- 默认参数：seed=1、steps=6、cfg=1、`er_sde`/`simple`、1536×1024；宽高按 16 四舍五入。
- 默认模型：`MOSS\moss_image_sft_4step_v2.0_int8_convrot.safetensors` + `Qwen3.5-9B-heretic-v2_int8_convrot.safetensors` + `flux2-vae_small.safetensors`，三个名字都能用参数替换。
- 输出默认写到执行命令时的当前目录，文件名是 `年月日-时分秒.png`；`-o/--output-dir`、`--name` 可改。
- 加速按 ComfyUI 的 `--use-ck-attention --fast-disk` 默认打开，`--no-ck-attention` / `--no-fast-disk` 可关。

## 环境陷阱

### 1. 行尾必须是 CRLF

仓库 `core.autocrlf=true`，所有源文件都是 CRLF。**任何用脚本重写文件的操作都必须用 `newline=''`**，否则整个文件会被改成 LF，`git diff` 爆炸成上千行。

```bash
python -c "d=open('docx2md.py','rb').read(); print(d.count(b'\r\n'), d.count(b'\n'))"
```

两个数必须相等。

### 2. 不要用 PowerShell here-string 写测试数据

`@"..."@` 会做 `$VAR` 插值，而 `\$` 在 PowerShell 里不是转义序列——写 LaTeX 测试用例必炸。用 `write` 工具创建测试脚本文件。

### 3. 控制台是 GBK

中文输出会显示成乱码，`×`(U+00D7) 会显示成 `*`（曾导致过保真度检查的误报）。**验证特殊字符要对比码点**：

```python
print([hex(ord(c)) for c in s])
```

### 4. OMML 的命名空间

公式元素在 `m` 命名空间（`http://schemas.openxmlformats.org/officeDocument/2006/math`），**属性也在 `m` 上，不是 `w`**。`docx2md.py` 里用 `_m()` / `_mtag()`，`md2docx.py` 里用 `_mtag()`。用 `qn('w:val')` 读 OMML 属性会永远拿到 `None`。

### 5. Word 的标题样式名不可信

中文模板的标题样式常叫「1级标题」「2级标题」，不是内置的 `Heading N`/`标题 N`。**认标题要看 `w:outlineLvl`（大纲级别）**；`w:outlineLvl val="9"` 表示「正文」，写了它就不再是标题（段落级与样式级都算数）。样式名只能当兜底。

编号的 `numFmt` 也不等于显示效果：`w:isLgl`（legal 编号）会让本级 `lvlText` 里的序号一律显示为阿拉伯数字。

### 6. 一个 run 不等于一个字

Word 会把连续文字拆成多个 run（拼写检查、格式切换、编辑痕迹），逐个 run 加 `**` 会产生 `**前半****后半**`。要按「格式相同的相邻片段」合并后再加标记。

### 7. 链接要写在最外层

超链接整体算一个片段，链接文字的加粗写在括号里面：`[**文字**](url)`。写成 `**[文字](url)**` 的话，`md2docx` 的 `**…**` 分支会先匹配上，把地址当正文写进文档。

### 8. 表格的行要排除嵌套表格

`w:tbl` 里允许嵌套 `w:tbl`。直接用 `.iter(qn('w:tr'))` 会把内层表格的行也遍历出来，而 `_tc_text()` 已经把内层文字算进父单元格，最终同一段文字出现两次；也不能简单用 `findall` 只取直接子元素，因为行可能被包在 `w:sdt` 里。统一走 `_table_rows()`（按最近的 `w:tbl` 祖先判断）。`_extract_table_grid()` 与 `_table_header_flags()` 的行数必须一致。

### 9. 块级公式的 `$$…$$` 要整段匹配

`docx2md.py` 对 `m:oMathPara`（块级公式）输出 `$$…$$`，对 `m:oMath`（行内公式）输出 `$…$`。`md2docx.py` 里如果先按 `$…$` 切分，`$$x+1$$` 会被拆成「`$`」+「`$x+1$`」+「`$`」。切分时 `$$…$$` 必须排在 `$…$` 前面；独占一行的 `$$…$$` 走 `display_match` 分支，表格单元格里的 `$$…$$` 靠 `_add_formatted_text()` 的切分顺序兜住。

### 10. `w:pPr` 的子元素有顺序要求

OOXML 规定 `w:pPr` 里 `w:pStyle` 必须排在 `w:numPr` 前面。`pPr.insert(0, numPr)` 会把 `w:numPr` 插到最前面，写出的文档虽然多数时候能打开，但顺序不合规。统一用 `pPr.get_or_add_numPr()`，不要手写 `insert`。

### 11. 超链接有好几种写法

Word 里的链接地址不一定在 `w:t` 旁边。抽文字只能拿到链接文字，地址会整个丢掉：

- `w:hyperlink` 带 `r:id` — 外部链接，地址在 `doc.part.rels` 的关系目标里；
- `w:hyperlink` 只有 `w:anchor`（或 `w:docLocation`）— 文档内部的书签链接，要转成 `#书签`；
- `w:fldSimple` 与复杂域（`w:fldChar` + `w:instrText`）— `HYPERLINK "地址"` 域，地址只在域代码里；
- 表格单元格、文本框、内容控件与修订里也可能有链接，要分别处理。

`mc:AlternateContent` 里 `mc:Choice` 与 `mc:Fallback` 是同一份内容的两种写法，只取一个，否则同一段文字会被算两遍。

### 12. 写超链接要用 `w:hyperlink` + 关系

`md2docx.py` 里把 `[文字](地址)` 写成「蓝色带下划线的普通 run」看着像链接，其实不可点击，`docx2md` 再读回来时也拿不到地址。外部地址用 `<w:hyperlink r:id="rIdN">` + `para.part.relate_to(url, RT.HYPERLINK, is_external=True)`；文档内部书签用 `<w:hyperlink w:anchor="书签名">`，**不建关系**。

链接文字的 run 在 `w:hyperlink` 里，**不是** `w:p` 的直属子元素，遍历要统一走 `_iter_runs()`。`w:hyperlink` 还必须插在链接文字原来的位置（`run._r.addprevious(hl)` 后再把 run 搬进去），直接 `append` 会打乱顺序。链接文字套 `Hyperlink` 字符样式（python-docx 默认模板里没有，要自己建），同时保留直接格式兜底。

### 13. 相邻表格只能按表头完全一致来合并

`merge_consecutive_tables()` 只服务于「跨页被拆成两个 `w:tbl` 的同一张表」，判据必须是表头完全一致。报告里的表单表每张都以「录像文件 / 起始井号 / 终止井号」开头，字符集几乎相同，若按字符相似度合并，就会把后一张表的首行当成重复表头丢掉。表头不同就各自成表，宁可多一张表也不能丢行。

### 14. `pdf2md.py` 本地推理的三个坑

- **`eos_token_id` 必须显式传 `[151643, 151645]`**。模型内层 `llm.generation_config` 里只有 `151643`，而模型实际用 `<|im_end|>`（`151645`）结束；漏掉它就会一路生成到 `--max-tokens`。用 `GenerationConfig.from_pretrained()` 读出后传给 `generate`，别硬编码。
- **`indicator_token_indices` 要重建**。低内存加载时它是非持久化 buffer，checkpoint 不覆盖，可能是未初始化内存，随后触发 CUDA `device-side assert`。加载后按 `visual_vocab_size` 与 `INDICATOR_IDS` 长度重建。
- **不要用 `trust_remote_code`**。transformers 5.x 会把模型文件复制到不同缓存子目录，导致配置类注册失效。现在的做法是把模型目录挂成临时包 `_oxr_local_model` 再导入。

### 15. 模型权重不在仓库里

`OXR/models/` 已写进 `.gitignore`（约 3 GB），下载方式见 `README.md`。`pdf2md.py` 校验 `OXR/models/OXR-1.0/config.json` 与 `layout_analysis/last.pth`，缺了会直接退出。

### 16. `pdf2md.py` 的提速结构不能被“顺手简化”掉

- **合批 key 是 `(kind, options)`，不含 `limit`**。阅读序传 `kind="reading_order"`，长短差异由 `run_batch` 里的按行 `row_limit` 截断兜住。
- **图解码器按「批大小 × 缓存长度」分桶复用**（`GreedyGraphDecoderPool`）；池子的字节预算只能给 3 GiB；`--chunk-pages` 默认 8 是有原因的，显存峰值随它线性上涨。
- **PDF 按 chunk 懒渲染，版面预取下一块**。`_layout_prefetched` 用 `id(images)` 作键，所以「渲染 + 版面预取」返回的列表对象必须原样传给流水线。
- **视觉塔跑独立 CUDA 流，与下一批的解码重叠**：解码热路径用 `current_stream().synchronize()`；视觉提交与 `torch.cuda.graph()` 捕获共用 `_CUDA_CAPTURE_LOCK`；`run_batch` 里要 `inputs_embeds.record_stream(...)`；`VISION_CHUNK_PATCHES` 必须切块且上限 ≤ 8192。
- **裁切去重缓存（`_crop_cache`）只缓存识别出的正文与格式，且每次 `process_one` 清空**；不能复用整个 `Block`，也不能跨文件/跨次保留。
- **合批按「预估输出长度」挑**（`_length_hint` / `_best_length_group`）：hint 只用于排序，不能进 `key`；攒批要等队列连续 30 ms 没增长或够两批。
- **图解码装不下预留缓存时只续跑没完成的行**（`residual`），按 `GENERATE_KV_BUDGET_BYTES` 拆组；退化行用 `_DegenerateLogitsProcessor` 逐行强制 eos，**不能用 `StoppingCriteria`**；`_is_degenerate_text` 的判据不要放宽。

### 17. `mdviewer.html` 里本地相对图的三个坑

- **图片 destination 带空格时要先编码成 `%20`**，否则 markdown-it 按 CommonMark 会把它当普通文字，图片不显示。
- **浏览器拿不到被打开本地文件的绝对路径**，渲染后要把 `<img src>` 换成 object URL；原始 src 存在 `dataset.origSrc` 里，重复解析时用它而不是 `img.currentSrc`。
- **file:// 页面 fetch(file://) 被 Chromium 禁止**，`?file=` 加载失败后只能弹提示条，等用户手势点「授权并打开…」选一次文件夹。drop 只收 `.md`/`.smd`/`.txt` 文件本身，不要把文件夹遍历加回来。

### 18. `md2docx.py` 的 HTML 表格有两个坑

- **合并单元格要在写内容之前做完**。`cell.merge()` 会把被合并格里的段落一起带进来；顺序是：建空表 → 按 `rowspan`/`colspan` 合并 → 只往合并后的左上角单元格写内容。
- **判断「哪些 run 是这次新增的」不能存 `id()`**。lxml 元素代理对象被回收后 `id` 会复用，必须 `set(para._p.iter(qn('w:r')))` 直接存元素本身。

### 19. `docx2md.py` 的合并表格不能复用 `_extract_table_grid()`

`_extract_table_grid()` 会把 `vMerge` 的续行重复首行文字、把 `gridSpan` 拆成空格子，`rowspan`/`colspan` 的原始信息已经丢了。写 HTML 表格时直接遍历 `w:tc`、按 `gridSpan`/`vMerge` 输出属性；表头只认 `w:tblHeader`，不要按拍平后的网格猜。普通无合并表格仍走原来的 GFM 路径。

### 20. 空的分节符段落不占编号

用「分节符」结束封面/目录那一节时，Word 把 `w:sectPr` 放进最后一个段落的 `w:pPr`，那个段落标记常带着「标题 1」样式（继承编号）。它没有文字，**Word 不把它算进编号序列**：正文第一个一级标题仍是「1」。判据见 `NumberingResolver._is_empty_section_break()`：带 `w:sectPr` 且没有可见内容才跳过。不带分节符的空标题段落 Word 照样占号，别按「空段落一律不占号」改。

### 21. `image-moss.py` 的 ComfyUI 接线

- **`comfy.cli_args` 不会自动解析命令行**：脚本不经过 `main.py`，`args` 全是默认值。`--use-ck-attention` / `--fast-disk` 要在 **`import nodes`（进而 import `comfy.model_management`、`comfy.ldm.modules.attention`）之前**直接写 `comfy.cli_args.args.use_ck_attention` / `args.fast_disk`；attention 后端是 import 时绑定的，晚设无效。跑起来会打印实际绑定的 `optimized_attention.__name__`，正常应为 `attention_comfy_kitchen_int8`。
- **模型要在退出前主动释放**：`del` 模型后调用 `comfy.model_management.unload_all_models()`，否则解释器退出时 `ModelPatcher.__del__` 会因模块已清空报 `'NoneType' object has no attribute 'ON_DETACH'`。
- **IMAGE 张量是 `[B,H,W,C]` 浮点 0~1**，存 PNG 前要 `detach().float().cpu()` 再转 uint8，PIL 只吃 3 维。
- **`VAEDecode.decode(vae, samples)` 的参数顺序与 `KSampler.sample()` 不同**，别凭节点输入顺序猜。

### 22. 4 GB 卡（GTX 1650）的 dtype 与 `--gtx1650` 预设

- **用 bf16，不要 fp16，也不要 fp32**。sm75 没有 bf16 硬件指令，torch 会走按 fp32 的通道（实测瘦 GEMM 1.6 TFLOPS，与 fp32 的 2.0 相当），而权重只占一半（2.6 GB vs 5.2 GB）；fp16 的瘦 GEMM 掉到 0.3 TFLOPS（慢 6 倍），模型级单请求慢 2 倍；fp32 权重 5.2 GB 装不进 4 GB，靠驱动换页后慢 8 倍。三种精度的识别结果逐字相同。
- **bf16 的注意力在 sm75 上没有 memory-efficient 内核**（只有 efficient 的 fp16 分支），只能走 math，会把 `(头 × 序列²)` 实体化：整页 3456 patch 单层就要 1.8 GiB。所以阅读序的整页图单独走 `READING_ORDER_MAX_PIXELS`（`--gtx1650` 给 320000），识别请求仍用 `MAX_PIXELS`；实测输出与全分辨率版逐字一致。
- **Windows 上显存超卖不报 OOM**：驱动把显存换页到内存，表现是“能跑，但 3 s/token”。判断依据是 `nvidia-smi` 的 used 贴着物理上限（4096 MiB）且 `reserved` 超过 `allocated` 一倍以上。`expandable_segments:True` 在这套 Windows + torch 2.15 上**能生效但必须禁用**：VMM 打映射时会 `Failed to create GPU mapping` / `cuMemSetAccess` 返回 OOM，随后分配器内部断言 `!handles_.at(i)` 直接崩掉整个文件（短跑 2~4 页看不出来，整本跑到第 6 批必踩）。真正的杠杆是把 chunk / 批 / 图缓存压小，再加 `RELEASE_CACHE_EACH_CHUNK`。
- **换出图解码器要真的还显存**：从池里 `pop` 掉只是丢引用，KV 缓存由 CUDA Graph 持有，要 `del` + `gc.collect()` + `torch.cuda.empty_cache()` 才还给驱动（`_release_cuda_memory()`）。捕获新图前必须先做完这一步，否则新旧两张图短暂共存，正好够把 4 GB 卡顶到换页甚至 OOM。
- **`--gtx1650` 只是预设，显式参数优先**（判据是仍等于 argparse 默认值）；`GRAPH_POOL_BUDGET_BYTES`、`GENERATE_KV_BUDGET_BYTES`、`VISION_CHUNK_PATCHES`、`VISION_OVERLAP`、`READING_ORDER_MAX_PIXELS`、`RELEASE_CACHE_EACH_CHUNK` 这些内部旋钮由 `_apply_gtx1650_preset` 改模块级全局。
- **捕获新图前先按预算淘汰旧图**：StaticCache 创建时一次性分配，新旧两张图共存会多占 400~900 MiB，正好够把 4 GB 卡顶到换页。

### 23. 版面模型（D-FINE）单张前向会输出全 NaN

- **判据**：只送 1 张页面时，`pred_boxes` / `pred_scores` 300 个候选全是 NaN，`layout_postprocess` 抛 `Invalid layout response: page 0 candidate 0 score must be finite`，整份文件作废；同一页跟着别的页一起送就完全正常。与页内容、显存、dtype 都无关，纯 batch 维度的坑。
- **踩法**：总页数是奇数时最后一批只有 1 页；`--pages 37-37` 这种单页调试也一样。
- **修法**：`LocalLayoutRuntime.infer()` 里单张时复制一张凑成 2 张，取第 1 行。实测该行与“和别的页同批”时**逐元素相同**（最大差 0.0），批内位置也不影响结果，代价只有一次多余的页面前向（约 0.5 s）。这是 pdf2md 的接线层，不用改 `OXR/services/`。

### 24. pdfcompress.py 的两阶段图像替换与整页判断

- **MuPDF 保存时会把会话内动过的图像流重编码为 raw RGB + Flate**：`replace_image()` / `insert_image()` 不保留原 filter。图像在 PDF 内的真实体积 ≈ `zlib.compress(裸 RGB, 6)`，**不能用磁盘文件体积**（如 pngquant 产物）当参照，PNG 候选的入 PDF 估计必须用裸像素压缩。
- **JPEG 替换必须走字节手术（stage 2）**：stage 1 保存后，把 JPEG 候选字节直接拼进已保存文件的图像对象（字典 `/Filter /DCTDecode` + 流），批量平移所有编号更高的对象偏移，最后重建 XRef 表（一遍扫 `N 0 obj`，startxref 在最后一行）。
- **`replace_image` 会清掉 `/SMask`**：先记下原 SMask 引用，替换后用 `xref_set_key(x, "SMask", "N 0 R")` 写回（普通字符串即可）。
- **读原 XRef 流要撤 PNG 预测**：MuPDF 存的 XRef 流带 `/DecodeParms << /Columns N /Predictor 12 >>` + FlateDecode。zlib 解压后必须逐行做 PNG 反预测（每行 N+1 字节：首字节是行内 filter 类型 0/2/3/4，其余是数据，Up 预测要用上一行值）。不撤预测，条目表整体错位，重建 XRef 会抄进垃圾条目。
- **扫对象不能要求 `<<`**：存在标量对象（`N 0 obj\n12345\nendobj`，纯数字被间接引用），没有字典。扫描模式用 `N 0 obj` + 空白即可；若要求 `<<`，这类对象会被误标 free，引用它的地方变成悬挂引用（MuPDF `insert_pdf` 报 `cannot find object in xref`，逐页读取却正常，极难发现）。
- **整页替换的参照用“该页在文档中的边际存储”**：内容流 +（图像大小 ÷ 引用该图像的页数），字体不计（全文档共享只存一份）。**不要用“该页单独提取成单页 PDF”的文件体积**：那会把共享字体、共享图像完整复制到每页，文本密集页的判断必然失真。
- **全局守卫（剪枝式）**：逐页的“边际存储”判断不可单独依赖（共享对象会被每页重编码、对象头与 XRef 有固定开销）。先用全部胜出页构建整文件，不更小就只保留胜算最强的一半重建，再不更小就保留四分之一；仍不更小才保留图片压缩版。胜算 = 边际存储 − PNG 入 PDF 体积估计。
- **页面资源里声明但从未绘制的图像**（`get_image_rects` 返回空）要跳过，否则 stage 1 崩。
- pngquant 退出码 98 = “压缩后没变小”，属正常结束，按接受处理。

## 改完自检

- [ ] 行尾仍是 CRLF
- [ ] `docx2md.py` 与 `md2docx.py` 同步改：`docx2md.py` 新写出的标记，`md2docx.py` 要能解析回 Word；`md2docx.py` 新写进 docx 的格式，`docx2md.py` 也要能还原成同样的标记
- [ ] 改到标题/编号时，`###### 图…`、`###### 表…` 仍分别还原成 Heading 8/9，`-n` 的图表编号仍与 `normalize_docx.py` 一致，往返不出现重复图号/表号
- [ ] 注释是中文
- [ ] 改到 `image-moss.py` 时能正常出图（快速冒烟：`--steps 1 --width 512 --height 512`）
- [ ] 改到 `pdfcompress.py` 时跑一份真实 PDF 冒烟（几页即可），确认输出不大于输入、页面渲染正常
- [ ] `git status` 里没有模型权重、`OXR/models/`、`*_media/`、`outputs/`
