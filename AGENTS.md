# AGENTS.md

给在本仓库工作的 AI 编码助手看的项目约定与验证流程。人读的说明见 `README.md`。

## 仓库里有什么

```
docx2md.py         .docx → .md
md2docx.py         .md → .docx
docx2pdf.py        .doc / .docx → .pdf
pptx2md.py         .pptx → .md
pdf2md.py          .pdf → .md / .json（本地 OXR 推理，依赖 OXR/）
fix_punct.py       .md 标点修正（英文标点 → 中文标点）
normalize_docx.py  .docx 格式标准化（主题字体、字号行距、标题多级列表）
requirements.txt
OXR/
  oxr/                        上游流水线最小子集（版面结果 → md/json）
  services/layout_analyze/    D-FINE 版面分析推理代码
  models/OXR-1.0/             模型权重（gitignore，自行下载）
```

`docx2md.py` 与 `md2docx.py` 是**单向转换器**，`md2docx.py` 的公式转换器专门服务于 `docx2md.py` 的 LaTeX 输出。改动一侧的输出格式，另一侧必须同步跟上，否则往返会静默损坏。`docx2pdf.py` 只负责调用本机 Microsoft Word 导出 PDF，不参与 Markdown 往返。`pptx2md.py` 只做 pptx → md，目前没有反向脚本，不改动上面这对往返。`fix_punct.py` 只做 md 的标点修正（中文文本里夹杂的英文标点替换为中文标点，含英文括号 `()` → `（）`，公式、代码、URL、英文段落不动），是单向后处理，不参与 Markdown 往返。`normalize_docx.py` 只做 docx 的格式标准化（字体：标题黑体/正文宋体、英文 Times New Roman；正文小四 1.5 倍行距、标题 1 小三、标题 2~9 小四；标题 1~9 加粗、无倾斜、段前段后为 0、标题 8/9 居中；多级列表，其中 8/9 级为图表编号；无间隔样式五号 1 倍行距；自动替换，如 m3/h → m³/h），是单向后处理，不参与 Markdown 往返。

`pdf2md.py` 是 pdf → md/json 的单向转换器（md 之后可以交给 `md2docx.py`）。它自己只负责 CLI、PDF 渲染、本地模型加载和文件输出，版面分析之后的流水线全部复用 `OXR/oxr/pipeline` 里的上游代码。`OXR/` 是从 [SoMarkAI/OXR](https://github.com/SoMarkAI/OXR) 复制的最小运行子集：`oxr/`（去掉 `cli/`、`server/`）加 `services/layout_analyze/dfine_la`，许可证见 `OXR/LICENSE`、`OXR/NOTICE`。模型权重 `OXR/models/OXR-1.0` 不在库里，用 modelscope 或 huggingface-cli 下载。

## 硬性约束

1. 注释用中文。
2. **`OXR/` 是上游代码子集，能不改就不改**。本地推理的接线（替换 `model_client` / `layout_client`、prompt、生成参数）都在 `pdf2md.py` 里；确实要改 `OXR/` 时在提交说明里写清改了什么、为什么，因为升级上游时会整块覆盖。
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
- **`indicator_token_indices` 要重建**。用 `from_pretrained` 低内存加载时它是非持久化 buffer，checkpoint 不覆盖，可能是未初始化内存（比如 `[2342367920352, 0, 0, 0]`），接着触发 CUDA `device-side assert`（`srcIndex < srcSelectDimSize`）。加载后按 `visual_vocab_size` 与 `INDICATOR_IDS` 长度重建。
- **不要用 `trust_remote_code`**。transformers 5.x 会把 `modeling_oxr.py` / `configuration_oxr.py` 复制到不同缓存子目录，`Siglip2NavitConfig` 注册失效，报 `Unrecognized configuration class`。现在的做法是把模型目录挂成临时包 `_oxr_local_model` 再导入，保证所有类来自同一份源码。

### 15. 模型权重不在仓库里

`OXR/models/` 已写进 `.gitignore`（约 3 GB）。换机器时重新下载：

```bash
pip install modelscope
modelscope download SoMark/OXR-1.0 --local_dir ./OXR/models/OXR-1.0
```

`pdf2md.py` 校验 `OXR/models/OXR-1.0/config.json` 与 `layout_analysis/last.pth`，缺了会直接退出并打印这条命令。

## 改完自检

- [ ] 行尾仍是 CRLF
- [ ] 一侧改了输出格式的话，另一侧同步跟上了
- [ ] 注释是中文
- [ ] `git status` 里没有模型权重、`OXR/models/`、`*_media/`、`outputs/`
