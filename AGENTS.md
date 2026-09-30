# AGENTS.md

给在本仓库工作的 AI 编码助手看的项目约定与验证流程。人读的说明见 `README.md`。

## 仓库里有什么

```
docx2md.py     .docx → .md，1791 行
md2docx.py     .md → .docx，987 行
requirements.txt  只有 python-docx
```

两个脚本是**单向转换器**，`md2docx.py` 的公式转换器专门服务于 `docx2md.py` 的 LaTeX 输出。改动一侧的输出格式，另一侧必须同步跟上，否则往返会静默损坏。

## 硬性约束

1. 注释用中文。

## 环境陷阱

这几个坑都实际踩过，会浪费大量时间：

### 1. 行尾必须是 CRLF

仓库 `core.autocrlf=true`，两个源文件都是 CRLF。**任何用脚本重写源文件的操作都必须用 `newline=''`**，否则整个文件会被改成 LF，`git diff` 爆炸成上千行。

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

## 改完自检

- [ ] 行尾仍是 CRLF
- [ ] 一侧改了输出格式的话，另一侧同步跟上了
- [ ] 注释是中文
