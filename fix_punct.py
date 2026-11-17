# -*- coding: utf-8 -*-
"""将 Markdown 文件中文本里夹杂的英文标点替换为中文标点。

用法:
    python fix_punct.py 文件1.md [文件2.md ...]

规则:
- 仅对“以中文为主”的行做替换(中文字数 >= 英文单词数),
  因此整段英文不受影响;
- , ; : ? ! ( ) 在前一字符或后一字符为中文时替换,
  数字(如 10,000)、时间(如 14:30)、版本号(如 1.0)里的标点保留;
- 圆括号成对处理:一对括号只要有一侧挨着中文就整对替换,
  避免出现 “（…)” 这种一半中文一半英文的混搭;
  段落内的软换行可跨行配对;
  紧跟在拉丁字母/数字后的单位/标准写法视为记号(如 dB(A)、De63(mm)、2L/(m²·d)),
  一律写成半角括号(源文件用了全角也改回来);
  单位相乘的中点 `·`(U+00B7) 与 `⋅`(U+22C5) 按英文处理,不算中文;
  括号里有中文或前面是中文时照常换全角;
  数量本来就不匹配的括号(缺一边或多一边)保持原样,不自动增删字符;
- . 在前一字符或后一字符为中文、或前一字符为收尾字符且处于行尾时替换;
- 双引号按全文配对转换为 “ ”(嵌套时用 ‘ ’),
  方向用错(如 ““ 或 ””)也会纠正;英文段落中的引号不动;
- HTML 标签(如 <td rowspan="2">)整段保护,标签里的属性引号不动;
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

# HTML 标签(含属性)，整段不参与标点替换。
# 只认常见标签名，避免把正文里的 “a<b>c” 之类误当标签吞掉。
_HTML_TAG_NAMES = (
    'a|abbr|b|big|blockquote|br|caption|center|cite|code|col|colgroup'
    '|dd|del|div|dl|dt|em|font|h[1-6]|hr|i|img|ins|kbd|li|mark|ol|p|pre'
    '|q|s|samp|small|span|strike|strong|table|tbody|td|tfoot|th|thead'
    '|tr|tt|u|ul|var'
)
_HTML_TAG_RE = re.compile(
    r'<!--.*?-->|</?(?:' + _HTML_TAG_NAMES + r')(?:\s[^<>]*)?/?>',
    re.IGNORECASE,
)

# 不带属性的 <sup>/<sub> 要留给 _replace_subsup 处理，不做保护区
_BARE_SUBSUP_RE = re.compile(r'</?(?:sup|sub)>', re.IGNORECASE)

# 圆括号(半角与全角)
OPEN_PARENS = '(（'
CLOSE_PARENS = ')）'

# 单位写法里允许和左括号紧贴的运算符/上标字符(字母数字另算)。
# 不包含 * ^ + - 等 Markdown 标记或通用运算符，避免把 **（`代码`） 这类误判成单位记号。
_UNIT_PREV_CHARS = set('/·⋅%‰°²³¹⁰⁴⁵⁶⁷⁸⁹')

# 这些收尾字符后面允许接中文句号
ZH_BOUNDARY = '")]】》”’"”'


def is_zh(ch):
    """判断是否为中文(含 CJK 统一表意文字与中文标点)。

    单位相乘用的中点 `·`(U+00B7) 不算中文：它在 m²·d 这类单位写法里是乘号，
    和中文的间隔号同形但语义不同，按英文处理，免得把整对括号带成全角。
    """
    if not ch:
        return False
    cp = ord(ch)
    return (
        0x4E00 <= cp <= 0x9FFF     # CJK 统一表意文字
        or 0x3400 <= cp <= 0x4DBF   # CJK 扩展 A
        or 0x20000 <= cp <= 0x2A6DF  # CJK 扩展 B
        or 0x3000 <= cp <= 0x303F   # CJK 标点
        or 0xFF00 <= cp <= 0xFFEF   # 全角形
        or ch in '“”‘’《》〈〉…—～'
    )


def _link_dest_spans(line):
    """链接地址 `](…)` 的区间，地址里的小括号按层数配对。

    只用 `\\]\\([^)]*\\)` 的话，`[图](http://a.com/x(1))` 会在第一个 `)` 处截断，
    剩下的 `)` 被当成正文参与括号配对，从而把这一行里别的括号带跑偏。
    """
    spans = []
    for m in re.finditer(r'\]\(', line):
        i = m.end()
        depth = 1
        while i < len(line) and depth:
            if line[i] == '(':
                depth += 1
            elif line[i] == ')':
                depth -= 1
            i += 1
        if depth == 0:
            spans.append((m.start(), i))
    return spans


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
    spans += _link_dest_spans(line)
    # HTML 标签整体保护，标签里的属性引号/逗号不能被当成正文标点
    for m in _HTML_TAG_RE.finditer(line):
        if _BARE_SUBSUP_RE.fullmatch(m.group(0)):
            continue
        spans.append(m.span())
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


def _near_zh(line, i):
    """判断位置 i 的前后字符是否有一个是中文。"""
    prev = line[i - 1] if i else ''
    nxt = line[i + 1] if i + 1 < len(line) else ''
    return is_zh(prev) or is_zh(nxt)


def _notation_pair(line, o, c):
    """判断一对括号是不是紧贴英文/数字的单位或标准记号。

    例：dB(A)、De63(mm)、2L/(m²·d)、ISO 9001(2015)。
    判据：左括号紧贴前一个英文/数字，或紧跟在 / · % ° 这类单位运算符后面，
    且括号里没有中文(上标、下标、°、℃、中点等都不算中文)。
    记号一律写成半角括号，源文件用了全角也改回来；
    括号里含中文(如 GB12348(A类))或前面是中文(如 压力(MPa))时不算记号。
    前面是反引号、引号、空格时也不算记号，避免把 `代码` 外面的全角括号改掉。
    """
    if o == 0:
        return False
    prev = line[o - 1]
    if not (prev.isascii() and prev.isalnum()) and prev not in _UNIT_PREV_CHARS:
        return False
    body = line[o + 1:c]
    return bool(body.strip()) and not any(is_zh(ch) for ch in body)


def prepare_line(line):
    """自动替换后计算保护区与“是否以中文为主”。

    自动替换会改变长度(m3/ → m³/、<sup>2</sup> → ²)，所以保护区必须在替换后的行上重算。
    返回 (处理后的行, 保护区掩码, 是否以中文为主)。
    """
    line = apply_replacements(line)
    prot = merge_spans(protected_spans(line))
    mask = [False] * len(line)
    for s, e in prot:
        mask[s:e] = [True] * (e - s)
    # 只在“可见文字”里统计，HTML 标签名、URL、代码里的英文不算，否则表格行会被判成英文行
    visible = ''.join(ch for i, ch in enumerate(line) if not mask[i])
    zh_count = sum(1 for ch in visible if is_zh(ch))
    en_words = len(re.findall(r'[A-Za-z]+', visible))
    return line, mask, zh_count >= en_words


def paren_plans(prepared):
    """对段落里成对的圆括号统一风格，返回每行的 {位置: 目标字符}。

    以前逐个字符判断，会出现 “（最小管径De63,最大不超过De200)” 这种
    左括号换了、右括号没换的情况：右括号前是英文、后是 HTML 标签，两侧都不挨中文。
    现在先配对，再决定这一对怎么换：任一侧挨着中文(或本身已是全角)就整对换成全角。

    数量本来就不匹配的括号一律不动，不增删字符：
    多余的 `)` 不会被改成 `）`，没有右括号的 `(` 也不会被改成 `（`，
    否则会凭空造出一个不闭合的全角括号。
    段落内的软换行允许跨行配对；空行是段落边界，调用方分段处理。
    """
    plans = [{} for _ in prepared]
    stack = []  # [(行号, 位置)]，记录还没闭合的左括号
    for li, (line, mask, _) in enumerate(prepared):
        for i, ch in enumerate(line):
            if mask[i]:
                continue
            if ch in OPEN_PARENS:
                stack.append((li, i))
            elif ch in CLOSE_PARENS:
                if not stack:
                    continue  # 落单的右括号：保持原样
                lo, o = stack.pop()
                oline, _, o_dom = prepared[lo]
                c_dom = prepared[li][2]
                # 单位/标准记号(dB(A)、De63(mm))一律写成半角；跨行的括号不按记号处理
                if lo == li and _notation_pair(line, o, i):
                    plans[lo][o] = '('
                    plans[li][i] = ')'
                    continue
                # 挨着中文要同时满足“该行以中文为主”，避免把英文行里的小括号改坏
                zh_side = ((_near_zh(oline, o) and o_dom)
                           or (_near_zh(line, i) and c_dom))
                if zh_side or oline[o] == '（' or ch == '）':
                    plans[lo][o] = '（'
                    plans[li][i] = '）'
    return plans


def fix_line(line, mask, zh_dominant, pplan, quote_depth):
    """把预处理好的单行渲染成中文标点版本。

    quote_depth 为当前未闭合的引号数(跨行保持)，返回 (处理后的行, 更新后的 quote_depth)。
    pplan 是该行成对括号的替换方案(见 paren_plans)。
    """
    out = []
    for i, ch in enumerate(line):
        if mask[i]:
            out.append(ch)
            continue
        prev = line[i - 1] if i else ''
        nxt = line[i + 1] if i + 1 < len(line) else ''
        if ch in OPEN_PARENS or ch in CLOSE_PARENS:
            out.append(pplan.get(i, ch))
        elif ch in ZP_MAP:
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
    """处理整个文本，保留原有行尾符。

    按段落(空行分隔)成批处理：圆括号的配对要在段落范围内看，
    这样软换行拆开的一对括号也能统一风格；围栏代码块整块跳过。
    """
    lines = text.split('\n')
    out = []
    quote_depth = 0
    in_code_fence = False
    fence_re = re.compile(r'\s*(```|~~~)')
    i = 0
    while i < len(lines):
        line = lines[i]
        if fence_re.match(line):
            in_code_fence = not in_code_fence
            out.append(line)
            i += 1
            continue
        if in_code_fence or not line.strip():
            out.append(line)
            i += 1
            continue
        # 收集一个段落：连续的、非空且非围栏的行
        block = []
        while (i < len(lines) and lines[i].strip()
               and not fence_re.match(lines[i])):
            block.append(lines[i])
            i += 1
        prepared = [prepare_line(l) for l in block]
        plans = paren_plans(prepared)
        for (pl_line, mask, zh_dominant), plan in zip(prepared, plans):
            pl_line, quote_depth = fix_line(
                pl_line, mask, zh_dominant, plan, quote_depth)
            out.append(pl_line)
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
