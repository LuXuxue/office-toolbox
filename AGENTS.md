# AGENTS.md

给在本仓库工作的 AI 编码助手看的项目约定与验证流程。人读的说明见 `README.md`。

## 仓库里有什么

```
docx2md.py     .docx → .md，2034 行
md2docx.py     .md → .docx，1044 行
docx2pdf.py    .doc / .docx → .pdf，258 行（调用本机 Word）
requirements.txt  python-docx；pywin32 仅 Windows（docx2pdf.py 用）
```

`docx2md.py` 与 `md2docx.py` 是**单向转换器**，`md2docx.py` 的公式转换器专门服务于 `docx2md.py` 的 LaTeX 输出。改动一侧的输出格式，另一侧必须同步跟上，否则往返会静默损坏。`docx2pdf.py` 只负责调用本机 Microsoft Word 导出 PDF，不参与 Markdown 往返。

## 硬性约束

1. 注释用中文。

## 环境陷阱

这几个坑都实际踩过，会浪费大量时间：

### 1. 行尾必须是 CRLF

仓库 `core.autocrlf=true`，三个 Python 源文件都是 CRLF。**任何用脚本重写源文件的操作都必须用 `newline=''`**，否则整个文件会被改成 LF，`git diff` 爆炸成上千行。

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

## 改完自检

- [ ] 行尾仍是 CRLF
- [ ] 一侧改了输出格式的话，另一侧同步跟上了
- [ ] 注释是中文
