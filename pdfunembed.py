#!/usr/bin/env python3
"""pdfunembed.py —— 清理 PDF 里的不可见元素，节省空间。

清理内容（全部默认执行）：
  1. 嵌入字体：字体描述符里的字体文件流（FontFile/FontFile2/FontFile3），
     并剥离字体名里的 6 位子集前缀（ABCDEF+SimSun → SimSun），
     让查看器按字体名匹配系统字体替代显示。
  2. 结构树（tagged PDF 的 /StructTreeRoot）：Word 直存的 PDF 里常有
     数万个 /StructElem 对象，页面上完全不可见；连同 /MarkInfo 与
     页面上的 /StructParentN 一起移除。
  3. 未使用的 XObject：页面资源里声明了、内容流里从未绘制的对象。
  4. 书签（/Outlines）、图层（/OCG，含页面 /Resources /Properties 里
     的引用）、嵌入文件附件（/Names）、XMP 元数据（/Metadata）。
  5. 隐藏文字：不可见文字绘制操作——文字渲染模式（Tr）为 3/4/6/7
     （不填充不描边）的文字。CAD 导出件里供 SHX 文字选择用的背景字
     属于此类，一并删除。不按颜色判断，正常的白字/灰字不受影响。

用法：
  python pdfunembed.py 输入.pdf [输出.pdf]
  输出默认为 <输入名>-unembed.pdf
"""

import argparse
import os
import re
import sys
import time

import pymupdf

PREFIX_RE = re.compile(r"^[0-9A-Z]{6}\+")


def strip_prefix(name):
    m = PREFIX_RE.match(name)
    return name[m.end():] if m else name


def ref_int(v):
    """'102 0 R' 或 102 → 102。"""
    if isinstance(v, int):
        return v
    return int(str(v).split()[0])


def has_key(doc, x, key):
    t, _ = doc.xref_get_key(x, key)
    return t != "null"


# ---------- 内容流：不可见文字清理 ----------

def tokenize(raw):
    """内容流 → [(kind, start, end)]；kind: num/name/str/hex/arr/op/other"""
    toks = []
    i, n = 0, len(raw)
    delim = (b"(", b")", b"[", b"]", b"{", b"}", b"<", b">", b"/")
    while i < n:
        c = raw[i:i + 1]
        if c.isspace():
            i += 1
            continue
        if c == b"%":
            j = raw.find(b"\n", i)
            i = n if j < 0 else j + 1
            continue
        if c in (b"(", b"[", b"/", b"<"):
            if c == b"(":
                depth, j = 1, i + 1
                while j < n:
                    cj = raw[j:j + 1]
                    if cj == b"\\":
                        j += 2
                        continue
                    if cj == b"(":
                        depth += 1
                    elif cj == b")":
                        depth -= 1
                        if depth == 0:
                            break
                    j += 1
                toks.append(("str", i, j + 1))
                i = j + 1
            elif c == b"<":
                if raw[i + 1:i + 2] == b"<":
                    toks.append(("other", i, i + 1))
                    i += 1
                else:
                    j = raw.find(b">", i)
                    toks.append(("hex", i, j + 1))
                    i = j + 1
            elif c == b"[":
                depth, j = 1, i + 1
                while j < n and depth:
                    cj = raw[j:j + 1]
                    if cj == b"(":
                        k = j + 1
                        while k < n:
                            if raw[k:k + 1] == b"\\":
                                k += 2
                                continue
                            if raw[k:k + 1] == b")":
                                break
                            k += 1
                        j = k + 1
                    elif cj == b"[":
                        depth += 1
                    elif cj == b"]":
                        depth -= 1
                    j += 1
                toks.append(("arr", i, j))
                i = j
            elif c == b"/":
                j = i + 1
                while j < n and not raw[j:j + 1].isspace() \
                        and raw[j:j + 1] not in delim:
                    j += 1
                toks.append(("name", i, j))
                i = j
        elif c in (b"-", b".") or c.isdigit():
            j = i
            while j < n and (raw[j:j + 1].isdigit()
                             or raw[j:j + 1] in (b".", b"-")):
                j += 1
            toks.append(("num", i, j))
            i = j
        else:
            j = i
            while j < n and not raw[j:j + 1].isspace() \
                    and raw[j:j + 1] not in delim:
                j += 1
            if j == i:  # 孤立分隔符（如 BDC 字典的 >），单独成 token
                j = i + 1
            toks.append(("op", i, j))
            i = j
    return toks


def strip_page_content(doc, page):
    """清理一页内容流里的不可见文字，返回 (删除操作数, 删除字节数)。

    判据只有文字渲染模式：`N Tr` 为 3/4/6/7 即不填充不描边（不可见）。
    不按颜色判断，避免误删文档里合法的白字/灰字。
    """
    ops = 0
    nbytes = 0
    for cx in page.get_contents():
        raw = doc.xref_stream(cx)
        toks = tokenize(raw)
        dead = set()
        tr_mode = 0
        stack = []
        for idx, (kind, s, e) in enumerate(toks):
            if kind != "op":
                continue
            op = raw[s:e].decode("latin-1")
            if op == "q":
                stack.append(tr_mode)
            elif op == "Q":
                if stack:
                    tr_mode = stack.pop()
            elif op == "Tr":
                if idx >= 1 and toks[idx - 1][0] == "num":
                    tr_mode = int(float(raw[toks[idx - 1][1]:toks[idx - 1][2]]))
            elif op in ("Tj", "TJ", "'", '"'):
                p = idx - 1
                if p >= 0 and toks[p][0] in ("str", "hex", "arr") \
                        and tr_mode in (3, 4, 6, 7):
                    dead.add((toks[p][1], toks[p][2]))
                    dead.add((s, e))
                    ops += 1
        if dead:
            # 重建：只保留 token 自身字节（跳过死 token 及其间空白）；
            # 两个“裸” token（num/name/op）相邻会合并成单个 token，补空格
            parts = []
            prev_kind = None
            for kind, s, e in toks:
                if (s, e) in dead:
                    continue
                if prev_kind in ("num", "name", "op") \
                        and kind in ("num", "name", "op"):
                    parts.append(b" ")
                parts.append(raw[s:e])
                prev_kind = kind
            new = b"".join(parts)
            nbytes += len(raw) - len(new)
            doc.update_stream(cx, new)
    return ops, nbytes


def main():
    ap = argparse.ArgumentParser(
        description="清理 PDF 里的不可见元素（嵌入字体、结构树、死 XObject、"
                    "书签、图层、附件、XMP、隐藏文字）")
    ap.add_argument("input", help="输入 PDF")
    ap.add_argument("output", nargs="?", help="输出 PDF（默认 <输入名>-unembed.pdf）")
    args = ap.parse_args()

    PDF = args.input
    OUT = args.output or os.path.splitext(PDF)[0] + "-unembed.pdf"

    t0 = time.time()
    size_before = os.path.getsize(PDF)
    doc = pymupdf.open(PDF)
    cat = doc.pdf_catalog()
    report = []

    # 1) 嵌入字体
    f_ops = 0
    f_bytes = 0
    n_fonts = 0
    for x in range(1, doc.xref_length()):
        if not doc.xref_is_font(x):
            continue
        n_fonts += 1
        t, v = doc.xref_get_key(x, "BaseFont")
        if t == "name" and PREFIX_RE.match(v):
            doc.xref_set_key(x, "BaseFont", strip_prefix(v))
        d_t, d_v = doc.xref_get_key(x, "FontDescriptor")
        if d_t != "xref":
            continue
        dx = ref_int(d_v)
        t3, v3 = doc.xref_get_key(dx, "FontName")
        if t3 == "name" and PREFIX_RE.match(v3):
            doc.xref_set_key(dx, "FontName", strip_prefix(v3))
        for k in ("FontFile", "FontFile2", "FontFile3"):
            t2, v2 = doc.xref_get_key(dx, k)
            if t2 == "xref":
                f_bytes += len(doc.xref_stream_raw(ref_int(v2)))
                doc.xref_set_key(dx, k, "null")
                f_ops += 1
    if n_fonts:
        report.append(f"嵌入字体: {n_fonts} 个字体"
                      + (f"，移除 {f_ops} 个字体文件（{f_bytes/1024:.0f} KB）"
                         if f_ops else "，无字体文件流"))

    # 2) 结构树
    if has_key(doc, cat, "StructTreeRoot"):
        n_elem = sum(1 for x in range(1, doc.xref_length())
                     if doc.xref_get_key(x, "Type")[1] == "/StructElem")
        doc.xref_set_key(cat, "StructTreeRoot", "null")
        doc.xref_set_key(cat, "MarkInfo", "null")
        for pno in range(doc.page_count):
            px = doc[pno].xref
            obj = doc.xref_object(px)
            if "/StructParent" in obj:
                doc.update_object(px, re.sub(
                    r"\s?/StructParent[ ]?\d+", " ", obj))
        report.append(f"结构树: 移除 {n_elem} 个 StructElem")

    # 3) 未使用的 XObject
    xo_ops = 0
    xo_bytes = 0
    for pno in range(doc.page_count):
        page = doc[pno]
        px = page.xref
        r_t, r_v = doc.xref_get_key(px, "Resources")
        if r_t != "xref":
            continue
        rx = ref_int(r_v)
        x_t, x_v = doc.xref_get_key(rx, "XObject")
        if x_t != "xref":
            continue
        xx = ref_int(x_v)
        content = b" ".join(doc.xref_stream(c) for c in page.get_contents())
        for k, v in doc.xref_get_keys(xx):
            if v[0] != "xref":
                continue  # 只处理间接引用（真正的 XObject）
            name = k.lstrip("/")
            if re.search(rb"/" + name.encode() + rb"(?![A-Za-z0-9_])",
                         content):
                continue  # 内容流里有引用
            t2, v2 = doc.xref_get_key(xx, k)
            if t2 == "xref":
                xo_bytes += len(doc.xref_stream_raw(ref_int(v2)))
            doc.xref_set_key(xx, k, "null")
            xo_ops += 1
    if xo_ops:
        report.append(f"死 XObject: 移除 {xo_ops} 个"
                      f"（{xo_bytes/1024:.0f} KB）")

    # 4) 书签 / 图层 / 附件 / XMP
    if has_key(doc, cat, "Outlines"):
        doc.xref_set_key(cat, "Outlines", "null")
        report.append("书签: 已移除")
    if has_key(doc, cat, "OCGs") or has_key(doc, cat, "OCProperties"):
        doc.xref_set_key(cat, "OCGs", "null")
        doc.xref_set_key(cat, "OCProperties", "null")
        n_refs = 0
        for pno in range(doc.page_count):
            px = doc[pno].xref
            r_t, r_v = doc.xref_get_key(px, "Resources")
            if r_t == "xref" and has_key(doc, ref_int(r_v), "Properties"):
                doc.xref_set_key(ref_int(r_v), "Properties", "null")
                n_refs += 1
            if has_key(doc, px, "OC"):
                doc.xref_set_key(px, "OC", "null")
        report.append(f"图层: 已移除（{n_refs} 处页面引用）")
    if has_key(doc, cat, "Names"):
        doc.xref_set_key(cat, "Names", "null")
        report.append("嵌入附件: 已移除")
    if has_key(doc, cat, "Metadata"):
        doc.xref_set_key(cat, "Metadata", "null")
        report.append("XMP 元数据: 已移除")

    # 5) 隐藏文字
    h_ops = 0
    h_bytes = 0
    for pno in range(doc.page_count):
        o, b_ = strip_page_content(doc, doc[pno])
        h_ops += o
        h_bytes += b_
    if h_ops:
        report.append(f"隐藏文字: 移除 {h_ops} 个绘制操作"
                      f"（{h_bytes/1024:.0f} KB）")

    doc.save(OUT, garbage=4, deflate=True)
    doc.close()

    # 验证
    chk = pymupdf.open(OUT)
    left_font = 0
    left_elem = 0
    bad = 0
    for x in range(1, chk.xref_length()):
        if chk.xref_get_key(x, "Type")[1] == "/FontDescriptor":
            for k in ("FontFile", "FontFile2", "FontFile3"):
                if chk.xref_get_key(x, k)[0] == "xref":
                    left_font += 1
        if chk.xref_get_key(x, "Type")[1] == "/StructElem":
            left_elem += 1
    for pno in range(chk.page_count):
        try:
            chk[pno].get_text()
            chk[pno].get_images()
        except Exception:
            bad += 1
    chk.close()

    size_after = os.path.getsize(OUT)
    saved = size_before - size_after
    print(f"输入: {size_before/1024/1024:.2f} MB")
    for line in report:
        print(f"  {line}")
    print(f"输出: {size_after/1024/1024:.2f} MB"
          f"（省 {saved/1024/1024:.2f} MB，{saved/size_before*100:.0f}%）")
    print(f"用时 {time.time()-t0:.1f} s")
    if left_font or left_elem or bad:
        sys.exit(f"验证未通过: 残留字体文件 {left_font}、StructElem {left_elem}、"
                 f"异常页 {bad}")


if __name__ == "__main__":
    main()
