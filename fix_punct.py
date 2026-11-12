# -*- coding: utf-8 -*-
"""将 Markdown 文件中文本里夹杂的英文标点替换为中文标点。

用法:
    python fix_punct.py 文件1.md [文件2.md ...]

规则:
- 仅对“以中文为主”的行做替换(中文字数 >= 英文单词数),
  因此整段英文不受影响;
- , ; : ? ! ( ) 在前一字符或后一字符为中文时替换,
  数字(如 10,000)、时间(如 14:30)、版本号(如 1.0)里的标点保留;
- . 在前一字符或后一字符为中文、或前一字符为收尾字符且处于行尾时替换;
- 双引号按全文配对转换为 “ ”(嵌套时用 ‘ ’),
  方向用错(如 ““ 或 ””)也会纠正;英文段落中的引号不动;
- 公式($...$、$$...$$)、代码(``` … ``` 与 `...`)、
  URL 与链接地址、代码块中的标点一律不替换。
- 自动替换计量单位等写法（如 m3/h → m³/h），
  同样只作用于未受保护的区域;
- 自动把上/下标标签转成对应的 UTF-8 字符，
  如 <sub>3</sub> → ₃、<sup>2</sup> → ²；
  内容含无法映射的字符(字母、** 等标记、公式)时保持原样;
- 自动替换用到 UTF-8 字符(如 ³ ₂),
  非 UTF-8 编码的 md 文件会先自动转换为 UTF-8 再处理(原来就是则跳过)。
"""

import re
import sys

# 需要替换的英文标点及其对应的中文标点
ZP_MAP = {
    ',': '，',
    ';': '；',
    ':': '：',
    '?': '？',
    '!': '！',
    '(': '（',
    ')': '）',
}

# 自动替换表：(原文, 替换文)，按顺序应用
_REPLACEMENTS = [
    ('m3/', 'm³/'),
    ('m2/', 'm²/'),
    ('/m3', '/m³'),
    ('/m2', '/m²'),
    ('km2', 'km²'),
    ('BOD5', 'BOD₅'),
    ('NH3', 'NH₃'),
    ('H2S', 'H₂S'),
    ('CO2', 'CO₂'),
]

# 上标标签 <sup>…</sup> 到 UTF-8 字符的映射
_SUP_MAP = {
    '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
    '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
    '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾',
}

# 下标标签 <sub>…</sub> 到 UTF-8 字符的映射
_SUB_MAP = {
    '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
    '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
    '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎',
}

# 上下标标签，内容不允许跨行
_SUBSUP_RE = re.compile(r'<(sup|sub)>(.*?)</\1>', re.IGNORECASE)

# 这些收尾字符后面允许接中文句号
ZH_BOUNDARY = '")]】》”’"”'


def is_zh(ch):
    """判断是否为中文(含 CJK 统一表意文字与中文标点)。"""
    if not ch:
        return False
    cp = ord(ch)
    return (
        0x4E00 <= cp <= 0x9FFF     # CJK 统一表意文字
        or 0x3400 <= cp <= 0x4DBF   # CJK 扩展 A
        or 0x20000 <= cp <= 0x2A6DF  # CJK 扩展 B
        or 0x3000 <= cp <= 0x303F   # CJK 标点
        or 0xFF00 <= cp <= 0xFFEF   # 全角形
        or ch in '“”‘’《》〈〉…—～·'
    )


def protected_spans(line):
    """返回行内不可改动区间(公式、代码、URL、链接地址、有序列表点)。"""
    # 代码围栏整行
    if re.fullmatch(r'\s*(```|~~~).*', line):
        return [(0, len(line))]
    # 整行的块级公式
    if re.fullmatch(r'\s*\$\$.*\$\$\s*', line):
        return [(0, len(line))]
    spans = []
    # 行内代码
    spans += [m.span() for m in re.finditer(r'`[^`\n]*`', line)]
    # 行内公式 $$...$$ 与 $...$
    spans += [m.span() for m in re.finditer(r'\$\$[^\$\n]*\$\$', line)]
    spans += [m.span() for m in re.finditer(r'(?<!\$)\$[^\$\n]*\$', line)]
    # URL 与链接地址
    spans += [m.span() for m in re.finditer(r'<https?://[^\s>]*>', line)]
    spans += [m.span() for m in re.finditer(r'https?://[^\s<>()]*', line)]
    spans += [m.span() for m in re.finditer(r'\]\([^)]*\)', line)]
    # 有序列表标记的点(如 "1. " 或 "1.本" 或 "a. " 或 "b.本")
    m = re.match(r'\s*(?:\d+|[a-zA-Z])\.(?=\D|$)', line)
    if m:
        pos = m.end() - 1
        spans.append((pos, pos + 1))
    return spans


def merge_spans(spans):
    """合并重叠区间。"""
    out = []
    for s, e in sorted(spans):
        if out and s <= out[-1][1]:
            out[-1] = (out[-1][0], max(out[-1][1], e))
        else:
            out.append((s, e))
    return out


def _replace_subsup(seg):
    """把 <sup>…</sup> / <sub>…</sub> 换成对应的 UTF-8 上/下标字符。

    内容里有无法映射的字符(字母、`**` 等 Markdown 标记、公式)时整段保持原样，
    避免把 `<sup>**2**</sup>` 这类带格式的写法改坏。
    """
    def repl(m):
        table = _SUP_MAP if m.group(1).lower() == 'sup' else _SUB_MAP
        body = m.group(2)
        if not body or any(ch not in table for ch in body):
            return m.group(0)
        return ''.join(table[ch] for ch in body)
    return _SUBSUP_RE.sub(repl, seg)


def _apply_replacements_to_seg(seg):
    """对一段文本按顺序应用自动替换表，再转换上/下标标签。"""
    for old, new in _REPLACEMENTS:
        if old in seg:
            seg = seg.replace(old, new)
    return _replace_subsup(seg)


def apply_replacements(line):
    """对行中未受保护的部分做自动替换。

    公式、代码、URL、链接地址里的原文保持不变(比如 LaTeX 里的 CO2 不能变成 CO₂)。
    """
    prot = merge_spans(protected_spans(line))
    if not prot:
        return _apply_replacements_to_seg(line)
    parts = []
    last = 0
    for s, e in prot:
        parts.append(_apply_replacements_to_seg(line[last:s]))
        parts.append(line[s:e])
        last = e
    parts.append(_apply_replacements_to_seg(line[last:]))
    return ''.join(parts)


def fix_line(line, in_code_fence, quote_depth):
    """处理单行。quote_depth 为当前未闭合的引号数(跨行保持)。

    返回 (处理后的行, 更新后的 quote_depth)。
    """
    if in_code_fence:
        return line, quote_depth

    # 自动替换(计量单位等)，必须在受保护区间断开后只替换未受保护部分；
    # 替换会改变长度，后面的 protected_spans 要在替换后的行上重新计算
    line = apply_replacements(line)

    prot = merge_spans(protected_spans(line))
    mask = [False] * len(line)
    for s, e in prot:
        mask[s:e] = [True] * (e - s)

    # 本行是否以中文为主:中文字数 >= 英文单词数
    zh_count = sum(1 for ch in line if is_zh(ch))
    en_words = len(re.findall(r'[A-Za-z]+', line))
    zh_dominant = zh_count >= en_words

    out = []
    for i, ch in enumerate(line):
        if mask[i]:
            out.append(ch)
            continue
        prev = line[i - 1] if i else ''
        nxt = line[i + 1] if i + 1 < len(line) else ''
        if ch in ZP_MAP:
            # 前一字符或后一字符为中文时替换
            if zh_dominant and (is_zh(prev) or is_zh(nxt)):
                out.append(ZP_MAP[ch])
            else:
                out.append(ch)
        elif ch == '.':
            # 前后为中文,或前为收尾字符且处于行尾
            end_of_line = not line[i + 1:].strip()
            if zh_dominant and (
                is_zh(prev) or is_zh(nxt)
                or (prev in ZH_BOUNDARY and end_of_line)
            ):
                out.append('。')
            else:
                out.append(ch)
        elif ch == '"':
            if zh_dominant:
                # 按深度配对:0 为开引号,1 闭外层,>=2 闭内层
                if quote_depth == 0:
                    out.append('“')
                    quote_depth = 1
                elif quote_depth == 1:
                    out.append('”')
                    quote_depth = 0
                else:
                    out.append('’')
                    quote_depth -= 1
            else:
                out.append(ch)
        elif ch in '“”':
            if zh_dominant:
                if ch == '“':
                    # 期望闭引号却出现开引号:按嵌套开引号处理(纠正用错的方向)
                    out.append('‘' if quote_depth else '“')
                    quote_depth += 1
                else:
                    # 期望开引号却出现闭引号:按当前层级的闭引号处理
                    out.append('’' if quote_depth >= 2 else '”')
                    quote_depth = max(quote_depth - 1, 0)
            else:
                out.append(ch)
        else:
            out.append(ch)
    return ''.join(out), quote_depth


def fix_text(text):
    """处理整个文本,保留原有行尾符。"""
    lines = text.split('\n')
    out = []
    quote_depth = 0
    in_code_fence = False
    fence_re = re.compile(r'\s*(```|~~~)')
    for line in lines:
        if fence_re.match(line):
            in_code_fence = not in_code_fence
        line, quote_depth = fix_line(line, in_code_fence, quote_depth)
        out.append(line)
    return '\n'.join(out)


def read_text_auto(path):
    """读取文本并检测编码，返回 (文本, 是否原本就是 UTF-8)。

    依次尝试 UTF-8、GB18030(GBK 的超集，覆盖中文 Windows 常见编码)。
    """
    with open(path, 'rb') as f:
        raw = f.read()
    try:
        return raw.decode('utf-8'), True
    except UnicodeDecodeError:
        pass
    try:
        return raw.decode('gb18030'), False
    except UnicodeDecodeError:
        raise ValueError('无法识别的编码(既不是 UTF-8 也不是 GBK/GB18030)')


def main():
    if len(sys.argv) < 2:
        print('用法: python fix_punct.py 文件1.md [文件2.md ...]')
        return 1
    rc = 0
    for path in sys.argv[1:]:
        try:
            text, is_utf8 = read_text_auto(path)
        except (OSError, ValueError) as e:
            print(f'读取失败: {path}: {e}')
            rc = 1
            continue
        if not is_utf8:
            # 自动替换用到 UTF-8 字符，先把文件转码为 UTF-8（保留原行尾），再进入处理流程
            with open(path, 'wb') as f:
                f.write(text.encode('utf-8'))
            print(f'已转换为 UTF-8: {path}')
        new_text = fix_text(text)
        if new_text == text:
            if is_utf8:
                print(f'无变化: {path}')
        else:
            with open(path, 'w', encoding='utf-8', newline='') as f:
                f.write(new_text)
            print(f'已更新: {path}')
    return rc


if __name__ == '__main__':
    sys.exit(main())
