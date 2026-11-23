#!/usr/bin/env python3
"""
pdfcompress.py —— 按目标 DPI 重编码 PDF 里的图像，缩小文件体积。

做法（三个阶段）：
  阶段 1（PyMuPDF）：抽取每张唯一图像，缩到目标 DPI，生成两个候选：
    * pngquant 量化 PNG
    * 指定质量的 JPEG
    取比原文件流小的最小者替换（保证只小不大）。PNG 胜者用
    replace_image() 在内存中替换，文档以 garbage=4, deflate=True,
    use_objstms=True 保存。
  阶段 2（原始字节手术）：MuPDF 保存时会把会话内动过的图像流重编码为
    裸像素，JPEG 无法通过 PyMuPDF 写入。因此 JPEG 胜者直接拼接进已保存
    文件的图像对象（字典 + 流），批量平移所有编号更高的对象偏移，最后
    重建 XRef 表。
  阶段 3（可选，--compress-pages）：在图片压缩后的文件上，把每页渲染成
    PNG 走 pngquant 压缩，胜出的页替换为整页图片。

整个流程尺寸安全：只有替换后确实更小才替换。
安全护栏：带 /ImageMask、/Mask、专色（Separation/DeviceN）的图跳过；
带 /SMask 的图把掩码与图像按同一比例缩放；灰度/CMYK 保留原通道与
ICC 色彩空间，不会统一扩成 RGB。PNG 胜者就地改流，不产生孤儿对象。
另默认清理不可见内容（隐藏文字、结构树、书签/附件/XMP、未绘制图像）；
不动嵌入字体流。--no-clean 可关闭清理。

用法：
  python pdfcompress.py 输入.pdf [输出.pdf] [选项]

  输出.pdf 可省略，默认为 <输入名>-<dpi>DPI.pdf
  （如 原文.pdf --dpi 96 → 原文-96DPI.pdf）。

选项：
  --dpi N             目标 DPI，只处理超过它的图像（默认 96）
  --pngquant PATH     pngquant 可执行文件路径
                      （默认：脚本旁的 pngquant/pngquant.exe）
  --png-quality S     图片压缩的 pngquant 质量区间（默认 "0-10"）
  --page-quality S    --compress-pages 的 pngquant 质量区间（默认 "0-20"）
  --jpeg-quality Q    JPEG 质量（默认 65）
  --jpeg-subsample S  JPEG 色度下采样 0/1/2（默认 2 = 4:2:0）
  --compress-pages    先做普通图片压缩，再把每页按 --dpi 渲染成 PNG
                      （pngquant，--page-quality），仅当 PNG 比该页在文档中
                      的边际存储（内容流 + 分摊后的共享图像，字体不计）小
                      才替换整页（默认关）
  --workdir DIR       候选图像临时目录（默认：输入旁的 <输入>.imgwork）

依赖：Python 3 + pymupdf（PyMuPDF）+ Pillow。
"""
import argparse
import hashlib
import io
import mmap
import os
import re
import subprocess
import sys
import time
import zlib

import pymupdf
from PIL import Image


def _colorspace_info(doc, xref):
    """解析图像的颜色空间。

    返回 (原样记号, ICC 通道数, 是否 ICCBased, 是否 Separation/DeviceN)。
    记号用于替换后写回 /ColorSpace，避免把 ICCBased 印刷色降级成设备色。
    """
    kind, val = doc.xref_get_key(xref, "ColorSpace")
    token = val if kind in ("name", "array", "xref") else None
    text = val or ""

    def _obj_text(n):
        try:
            return doc.xref_object(n, compressed=True)
        except Exception:
            return ""

    m = re.search(r"/ICCBased\s+(\d+)\s+0\s+R", text)
    if m is None and kind == "xref" and token:
        mm = re.fullmatch(r"(\d+)\s+0\s+R", token.strip())
        if mm:
            t = _obj_text(int(mm.group(1)))
            text += " " + t
            m = re.search(r"/ICCBased\s+(\d+)\s+0\s+R", t)
    if m:
        # ICCBased 的 /N 在它指向的 ICC 流字典里
        text += " " + _obj_text(int(m.group(1)))

    nm = re.search(r"/N\s+(\d+)", text)
    n = int(nm.group(1)) if nm else None
    return (token, n, "/ICCBased" in text,
            bool(re.search(r"/(Separation|DeviceN)\b", text)))


def _make_smask(doc, mask):
    """新建一个 DeviceGray 软掩码对象，返回 "N 0 R" 引用字符串。"""
    x = doc.get_new_xref()
    w, h = mask.size
    raw = mask.tobytes("raw")
    doc.update_object(x, f"<</Type/XObject/Subtype/Image/Width {w}"
                         f"/Height {h}/ColorSpace/DeviceGray"
                         f"/BitsPerComponent 8/Length {len(raw)}>>")
    doc.update_stream(x, raw, compress=True)
    return f"{x} 0 R"


# ============================================================
# 非绘制内容清理 —— 默认开启，只删视觉上不可见的对象/操作。
# 做法参考 pdfunembed.py（不含字体流清理）。
# ============================================================
def _clean_tokenize(raw):
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
            if j == i:  # 孤立分隔符（如 BDC 内联字典的 >），单独成 token
                j = i + 1
            toks.append(("op", i, j))
            i = j
    return toks


def _strip_hidden_text(doc):
    """删除内容流里 Tr 为 3/4/6/7（不填充不描边）的不可见文字绘制操作。

    判据只有渲染模式，不按颜色判断，避免误删合法的白字/灰字。
    返回 (删除操作数, 删除字节数)。
    """
    ops = nbytes = 0
    for pno in range(doc.page_count):
        for cx in doc[pno].get_contents():
            raw = doc.xref_stream(cx)
            toks = _clean_tokenize(raw)
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
                        tr_mode = int(float(raw[toks[idx - 1][1]:
                                                 toks[idx - 1][2]]))
                elif op in ("Tj", "TJ", "'", '"'):
                    p = idx - 1
                    if p >= 0 and toks[p][0] in ("str", "hex", "arr") \
                            and tr_mode in (3, 4, 6, 7):
                        dead.add((toks[p][1], toks[p][2]))
                        dead.add((s, e))
                        ops += 1
            if dead:
                # 只拼 token 自身字节，跨间隙取会把死 token 又带回来；
                # 两个裸 token 相邻时补空格，避免合并成一个
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


def _remove_struct_tree(doc):
    """移除 tagged PDF 的结构树（页面上完全不可见）。返回 StructElem 数。"""
    cat = doc.pdf_catalog()
    if doc.xref_get_key(cat, "StructTreeRoot")[0] == "null":
        return 0
    n = sum(1 for x in range(1, doc.xref_length())
            if doc.xref_get_key(x, "Type")[1] == "/StructElem")
    doc.xref_set_key(cat, "StructTreeRoot", "null")
    doc.xref_set_key(cat, "MarkInfo", "null")
    for pno in range(doc.page_count):
        px = doc[pno].xref
        obj = doc.xref_object(px)
        if "/StructParent" in obj:
            doc.update_object(px, re.sub(
                r"\s?/StructParent[ ]?\d+", " ", obj))
    return n


def _remove_nonvisual_meta(doc):
    """移除书签、附件名册、XMP 元数据（都不参与绘制）。返回移除项数。"""
    cat = doc.pdf_catalog()
    n = 0
    for key in ("Outlines", "Names", "Metadata"):
        if doc.xref_get_key(cat, key)[0] != "null":
            doc.xref_set_key(cat, key, "null")
            n += 1
    return n


def _xobj_xref(doc, res_xref):
    """resources 字典里的 /XObject 字典 xref，没有则 None。"""
    t, v = doc.xref_get_key(res_xref, "XObject")
    return int(v.split()[0]) if t == "xref" else None


def _used_xobject_names(doc, page):
    """递归收集页面内容流（含嵌套 Form XObject）里实际 Do 的资源名。"""
    used = set()
    seen = set()

    def walk(res_xref, content):
        names = {m.group(1).decode("latin-1")
                 for m in re.finditer(rb"/([A-Za-z0-9_]+)\s+Do", content)}
        used.update(names)
        xx = _xobj_xref(doc, res_xref) if res_xref else None
        if not xx:
            return
        for k, v in doc.xref_get_keys(xx):
            name = k.lstrip("/")
            if name not in names or v[0] != "xref":
                continue
            ox = int(v[1].split()[0])
            if (doc.xref_get_key(ox, "Subtype")[1] != "/Form"
                    or ox in seen):
                continue
            seen.add(ox)
            rt, rv = doc.xref_get_key(ox, "Resources")
            fres = int(rv.split()[0]) if rt == "xref" else None
            walk(fres, doc.xref_stream(ox))

    rt, rv = doc.xref_get_key(page.xref, "Resources")
    res = int(rv.split()[0]) if rt == "xref" else None
    content = b" ".join(doc.xref_stream(c) for c in page.get_contents())
    walk(res, content)
    return used


def _drop_dead_image_resources(doc):
    """移除页面资源里声明了但从未绘制的图像 XObject。

    只处理图像，且递归统计 Form XObject 内部的 Do，避免 Issue 6 的误删；
    带注释的页面跳过（注释外观流有自己的资源）。
    """
    removed = 0
    for pno in range(doc.page_count):
        page = doc[pno]
        if list(page.annots() or []):
            continue
        rt, rv = doc.xref_get_key(page.xref, "Resources")
        if rt != "xref":
            continue
        xx = _xobj_xref(doc, int(rv.split()[0]))
        if not xx:
            continue
        used = _used_xobject_names(doc, page)
        for k, v in list(doc.xref_get_keys(xx)):
            if v[0] != "xref" or k.lstrip("/") in used:
                continue
            ox = int(v[1].split()[0])
            if doc.xref_get_key(ox, "Subtype")[1] != "/Image":
                continue
            doc.xref_set_key(xx, k, "null")
            removed += 1
    return removed


# ============================================================
# 阶段 3：整页压缩 —— 每页渲染成 PNG（pngquant，--page-quality），
# 仅当 PNG 比该页在文档中的边际存储（内容流 + 图像/共享使用次数；
# 字体不计）小时才替换整页。在图片压缩后的文件上运行。
# ============================================================
def run_page_compress(src, out, args, pngquant):
    import tempfile
    t_start = time.time()
    doc = pymupdf.open(src)
    n = doc.page_count

    # ---- 本文件中每页的“边际存储” ----------------
    # 单独提取的单页 PDF 是不合适的参照：它会把共享字体、共享图像
    # 完整复制进每页，文本页看起来远比实际贵。一页真正的边际成本是：
    #   * 它的内容流（页专属）
    #   * 它的图像，按引用该图像的页数分摊
    #   * 字体：不计（全文档共享，只存一份）
    img_count = {}
    for pno in range(n):
        for img in doc[pno].get_images(full=True):
            img_count[img[0]] = img_count.get(img[0], 0) + 1
    img_size = {}
    for x in img_count:
        try:
            sz = len(doc.xref_stream_raw(x))
        except Exception:
            try:
                sz = len(doc.xref_stream(x))
            except Exception:
                sz = 0
        # 软掩码（透明度）是独立对象，一并计入
        _, smask_t = doc.xref_get_key(x, "SMask")
        if smask_t != "null":
            try:
                smx = int(smask_t.split()[0])
                try:
                    sz += len(doc.xref_stream_raw(smx))
                except Exception:
                    sz += len(doc.xref_stream(smx))
            except Exception:
                pass
        img_size[x] = sz
    refs = []
    for pno in range(n):
        page = doc[pno]
        size = 0.0
        for c in page.get_contents():
            try:
                size += len(doc.xref_stream_raw(c))
            except Exception:
                size += len(doc.xref_stream(c))
        for img in page.get_images(full=True):
            x = img[0]
            size += img_size[x] / img_count[x]
        refs.append(size)
    wins = {}                     # pno -> (量化 PNG 路径, 入 PDF 后的体积估计)
    with tempfile.TemporaryDirectory(prefix="pdfpage_") as tmp:
        for pno in range(n):
            page = doc[pno]
            pix = page.get_pixmap(dpi=int(args.dpi), alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height),
                                  pix.samples)
            png = os.path.join(tmp, f"page_{pno}.png")
            img.save(png)
            try:
                r = subprocess.run(
                    [pngquant, f"--quality={args.page_quality}",
                     "--ext=.png", "--force", "--skip-if-larger",
                     "--speed=1", png],
                    capture_output=True)
                # 退出码 98 = 压缩后没变小：未量化的 PNG 留作候选
                if r.returncode not in (0, 98):
                    raise subprocess.CalledProcessError(r.returncode,
                                                        pngquant)
            except (subprocess.CalledProcessError, OSError):
                print(f"  第 {pno + 1} 页：pngquant 失败，保留原页")
                continue
            # PNG 入 PDF 后变成 Flate(裸像素)，实际体积取 PNG 文件与
            # 裸像素 zlib 流的较大者
            png_size = max(os.path.getsize(png),
                           len(zlib.compress(img.tobytes("raw"), 6)))
            # 参照：该页在本文件中的边际存储（内容流 + 图像/使用次数）
            if png_size < refs[pno]:
                wins[pno] = (png, png_size)
            if (pno + 1) % 50 == 0:
                print(f"  已检查 {pno + 1}/{n} 页 "
                      f"({time.time() - t_start:.0f}s)")

        # 全局守卫（剪枝式）：逐页的“边际存储”估计可能集体偏小
        # （共享对象开销、对象头与 XRef 体积等）。先用全部胜出页构建
        # 候选文件；整文件不更小就只保留胜算较强的一半重建；再不更小
        # 就保留前四分之一重建；仍不更小则保留图片压缩版。
        src_size = os.path.getsize(src)
        accepted = False
        n_kept = 0
        if wins:
            # 按“胜算强度”（边际存储 - PNG 体积）从大到小排序
            order = sorted(wins, key=lambda p: refs[p] - wins[p][1],
                           reverse=True)
            for frac in (1.0, 0.5, 0.25):
                k = int(round(len(order) * frac))
                if k <= 0:
                    break
                keep = set(order[:k])
                cand = os.path.join(tmp, f"candidate_{frac}.pdf")
                new = pymupdf.open()
                for pno in range(n):
                    if pno in keep:
                        page = doc[pno]
                        newpage = new.new_page(width=page.rect.width,
                                               height=page.rect.height)
                        newpage.insert_image(newpage.rect,
                                             filename=wins[pno][0])
                    else:
                        new.insert_pdf(doc, from_page=pno, to_page=pno)
                new.save(cand, garbage=4, deflate=True)
                new.close()
                if os.path.getsize(cand) < src_size:
                    os.replace(cand, out)
                    accepted = True
                    n_kept = k
                    break
                if frac == 0.25:
                    break
                print(f"  替换 {k} 页后整文件更大 "
                      f"({os.path.getsize(cand) / 1024 / 1024:.2f} MB vs "
                      f"{src_size / 1024 / 1024:.2f} MB)；"
                      f"只保留胜算较强的 {k // 2} 页重建")
        doc.close()

        if not accepted:
            if wins:
                print(f"  即使只替换最强的 "
                      f"{int(round(len(wins) * 0.25))} 页，整文件仍更大；"
                      f"保留图片压缩版")
            else:
                print("  没有页面能压成更小的 PNG；保留图片压缩版")
            os.replace(src, out)

    # 验证：每一页必须可读（内容 + 字体 + 图像）
    chk = pymupdf.open(out)
    bad = 0
    for p in range(chk.page_count):
        try:
            chk[p].get_text()
            chk[p].get_images()
        except Exception:
            bad += 1
    ok = chk.page_count == n and bad == 0
    print(f"页数: {chk.page_count}（应为 {n}），不可读: {bad}")
    print(f"替换为整页 PNG 的页数: {n_kept}")
    print(f"输出: {os.path.getsize(out) / 1024 / 1024:.2f} MB "
          f"(输入 {src_size / 1024 / 1024:.2f} MB)")
    print(f"整页压缩用时 {time.time() - t_start:.0f}s")
    chk.close()
    if not ok:
        sys.exit("整页压缩产生了无效输出")


def main():
    ap = argparse.ArgumentParser(
        description="按目标 DPI 重编码 PDF 里的图像，缩小文件体积。")
    ap.add_argument("input", help="输入 PDF")
    ap.add_argument("output", nargs="?", default=None,
                    help="输出 PDF（默认：<输入名>-<dpi>DPI.pdf）")
    ap.add_argument("--dpi", type=float, default=96,
                    help="目标 DPI，只处理超过它的图像（默认 96）")
    ap.add_argument("--pngquant", default=None,
                    help="pngquant 可执行文件路径"
                         "（默认：脚本旁的 pngquant/pngquant.exe）")
    ap.add_argument("--png-quality", default="0-10",
                    help='图片压缩的 pngquant 质量区间（默认 "0-10"）')
    ap.add_argument("--page-quality", default="0-20",
                    help='整页压缩（--compress-pages）的 pngquant 质量区间'
                         '（默认 "0-20"）')
    ap.add_argument("--jpeg-quality", type=int, default=65,
                    help="JPEG 质量（默认 65）")
    ap.add_argument("--jpeg-subsample", type=int, default=2,
                    choices=[0, 1, 2],
                    help="JPEG 色度下采样 0/1/2（默认 2 = 4:2:0）")
    ap.add_argument("--compress-pages", action="store_true",
                    help="在普通图片压缩之后，把每页按 --dpi 渲染成 PNG"
                         "（pngquant，--page-quality），仅当 PNG 小于该页"
                         "在文档中的边际存储（内容 + 共享图像，字体不计）"
                         "才替换整页；被替换页的文字不可选中（默认关）")
    ap.add_argument("--workdir", default=None,
                    help="候选图像临时目录（默认：<输入>.imgwork）")
    ap.add_argument("--no-clean", action="store_true",
                    help="关闭非绘制内容清理（默认开启：不可见文字、结构树、"
                         "书签/附件/XMP、未绘制图像；不动嵌入字体）")
    args = ap.parse_args()

    PDF = os.path.abspath(args.input)
    if args.output:
        OUT = os.path.abspath(args.output)
    else:
        OUT = os.path.join(os.path.dirname(PDF),
                           os.path.splitext(os.path.basename(PDF))[0]
                           + f"-{int(args.dpi)}DPI.pdf")
    if not os.path.isfile(PDF):
        sys.exit(f"找不到输入文件: {PDF}")
    if os.path.normcase(OUT) == os.path.normcase(PDF):
        sys.exit("输出路径与输入相同；请指定不同的输出文件，避免覆盖原件")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    pngquant = args.pngquant or os.path.join(script_dir, "pngquant", "pngquant.exe")
    if not os.path.isfile(pngquant):
        sys.exit(f"找不到 pngquant: {pngquant}")
    WORK = os.path.abspath(args.workdir or PDF + ".imgwork")
    os.makedirs(WORK, exist_ok=True)

    t_start = time.time()
    # ============================================================
    # 阶段 1：PyMuPDF —— 抽取、缩放、选候选、PNG 替换
    # ============================================================
    doc = pymupdf.open(PDF)
    # 用户密码非空时，页面对象不可访问；空密码能打开（仅 owner 加密）则
    # save 会输出解密版，无害。
    if doc.is_encrypted and not doc.authenticate(""):
        doc.close()
        sys.exit("输入 PDF 已加密（需要用户密码），暂不支持")
    images = {}
    for pno in range(doc.page_count):
        page = doc[pno]
        seen = set()
        for img in page.get_images(full=True):
            x = img[0]
            if x in seen:
                continue
            seen.add(x)
            if x in images:
                continue
            rects = page.get_image_rects(x)
            if not rects:
                continue      # 资源里声明了但本页没有绘制，跳过
            images[x] = dict(page=pno, w=img[2], h=img[3], rect=rects[0])
    n_pages = doc.page_count
    print(f"唯一图像: {len(images)}")

    # ---- 非绘制内容清理（默认开启；不动嵌入字体） ----
    clean = not args.no_clean
    clean_report = []
    if clean:
        n_ops, n_bytes = _strip_hidden_text(doc)
        if n_ops:
            clean_report.append(f"隐藏文字 {n_ops} 个绘制操作"
                                f"（{n_bytes/1024:.0f} KB）")
        n_struct = _remove_struct_tree(doc)
        if n_struct:
            clean_report.append(f"结构树 {n_struct} 个 StructElem")
        n_meta = _remove_nonvisual_meta(doc)
        if n_meta:
            clean_report.append(f"书签/附件/XMP {n_meta} 项")
        n_dead = _drop_dead_image_resources(doc)
        if n_dead:
            clean_report.append(f"未绘制图像 {n_dead} 个")
        if clean_report:
            print("清理: " + "，".join(clean_report))

    stats = {"png": 0, "jpg": 0, "skip": 0,
             "skip_mask": 0, "skip_cs": 0, "fail": 0}
    jpg_wins = []   # (旧 xref, 页码, w, h, jpg 路径, w2, h2, 原流哈希, 模式)
    for x, im in images.items():
        pno = im["page"]
        # 着色模板（/ImageMask）与色键掩码（/Mask）替换后会改变绘制语义
        # 或直接丢掉掩码，跳过并计入统计
        imk_kind, imk_val = doc.xref_get_key(x, "ImageMask")
        if (imk_kind == "bool" and imk_val == "true"
                or doc.xref_get_key(x, "Mask")[0] != "null"):
            stats["skip_mask"] += 1
            continue
        # 软掩码（透明度）是独立对象；记下引用与 xref
        _, smask_t = doc.xref_get_key(x, "SMask")
        smask_ref = smask_t if smask_t != "null" else None
        smask_x = None
        if smask_ref:
            m = re.fullmatch(r"(\d+)\s+0\s+R", smask_ref.strip())
            if m:
                smask_x = int(m.group(1))

        try:
            pix = pymupdf.Pixmap(doc, x)
        except Exception:
            stats["fail"] += 1
            continue
        w, h = pix.width, pix.height
        alpha = bool(pix.alpha)
        channels = pix.n - int(alpha)
        mode = {1: "L", 3: "RGB", 4: "CMYK"}.get(channels)
        if mode is None or (alpha and mode == "CMYK"):
            stats["skip_cs"] += 1
            continue

        # 按原颜色空间分路：灰度保持单通道，CMYK 保持四通道；
        # ICCBased 尽量沿用原色彩空间，避免印刷件偏色
        token, icc_n, is_icc, special = _colorspace_info(doc, x)
        if special:
            # Separation / DeviceN 转 RGB 会丢专色，跳过
            stats["skip_cs"] += 1
            continue
        default_cs = {"L": "/DeviceGray", "RGB": "/DeviceRGB",
                      "CMYK": "/DeviceCMYK"}[mode]
        if token in ("/DeviceRGB", "/DeviceGray", "/DeviceCMYK"):
            cs_out = token if token == default_cs else default_cs
        elif is_icc and (icc_n is None or icc_n == channels):
            cs_out = token
        else:
            cs_out = default_cs

        # 由显示矩形推算实际 DPI
        disp_w_in = im["rect"].width / 72.0
        dpi = w / disp_w_in if disp_w_in > 0 else args.dpi
        if dpi > args.dpi:
            scale = args.dpi / dpi
            w2, h2 = max(1, round(w * scale)), max(1, round(h * scale))
        else:
            w2, h2 = w, h

        # 取像素。带 alpha 的 pixmap 已被 MuPDF 反预乘，转 PNG 再读能
        # 拿到直通色；没有独立 SMask 时把 alpha 通道保留成新的软掩码。
        smask_src = None
        try:
            if alpha:
                rgba = Image.open(io.BytesIO(pix.tobytes("png")))
                if smask_x is None:
                    smask_src = rgba.getchannel("A").convert("L")
                img = rgba.convert(mode)
            else:
                img = Image.frombytes(mode, (w, h), pix.samples)
        except Exception:
            stats["fail"] += 1
            continue

        if smask_x is not None:
            try:
                spix = pymupdf.Pixmap(doc, smask_x)
                if spix.alpha:
                    spix = pymupdf.Pixmap(spix, 0)
                if spix.n != 1:
                    spix = pymupdf.Pixmap(pymupdf.csGRAY, spix)
                smask_src = Image.frombytes(
                    "L", (spix.width, spix.height), spix.samples)
            except Exception:
                # 掩码取不出来就跳过，避免写回尺寸不符的 /SMask
                stats["skip_mask"] += 1
                continue

        need_resize = (w2, h2) != (w, h)
        img_out = img.resize((w2, h2), Image.LANCZOS) if need_resize else img
        # 有软掩码且缩放了，掩码必须同步缩放；内嵌 alpha 则总是新建
        mask_out = None
        if smask_src is not None and (need_resize or smask_x is None):
            mask_out = smask_src.resize((w2, h2), Image.LANCZOS)

        base = os.path.join(WORK, f"{x}")
        cands = {}
        # 候选 1：pngquant PNG（原地量化，skip-if-larger）。
        # CMYK 不能存 PNG，只走 JPEG。
        png = base + ".png"
        if mode != "CMYK":
            img_out.save(png)
            try:
                r = subprocess.run(
                    [pngquant, f"--quality={args.png_quality}",
                     "--ext=.png", "--force", "--skip-if-larger",
                     "--speed=1", png],
                    capture_output=True)
                # 退出码 98 = “压缩后没变小”：原 PNG 留作有效候选
                if r.returncode not in (0, 98):
                    raise subprocess.CalledProcessError(r.returncode, pngquant)
            except (subprocess.CalledProcessError, OSError):
                stats["fail"] += 1
                continue
            # PNG 入 PDF 后是 Flate(裸像素)；用量化后的 PNG 实际像素
            # 估计体积，并留作下面就地替换的像素源
            qimg = Image.open(png)
            if qimg.mode != mode:
                qimg = qimg.convert(mode)
            png_raw = qimg.tobytes("raw")
            cands["png"] = max(os.path.getsize(png),
                               len(zlib.compress(png_raw, 6)))
        # 候选 2：JPEG。CMYK 的色度下采样会明显伤小字/细线，固定 0；
        # Pillow 写出的 CMYK JPEG 是反相存储，补丁字典里要加 /Decode。
        jpg = base + ".jpg"
        sub = 0 if mode == "CMYK" else args.jpeg_subsample
        img_out.save(jpg, format="JPEG", quality=args.jpeg_quality,
                     subsampling=sub)
        # JPEG 入 PDF 后是裸 DCTDecode 或 Flate 包裹，取小者
        jpg_bytes = open(jpg, "rb").read()
        cands["jpg"] = min(len(jpg_bytes),
                           len(zlib.compress(jpg_bytes, 6)))
        # 原流体积（判断“只小不大”的基准）；缩放时把掩码体积变化也算上
        try:
            orig_raw = doc.xref_stream_raw(x)
        except Exception:
            orig_raw = doc.xref_stream(x)
        orig_size = len(orig_raw)
        orig_hash = hashlib.md5(orig_raw).digest()
        mask_orig_size = 0
        if smask_x is not None:
            try:
                mask_orig_size = len(doc.xref_stream_raw(smask_x))
            except Exception:
                mask_orig_size = 0
        if mask_out is not None:
            mask_new_size = len(zlib.compress(mask_out.tobytes("raw"), 6))
        else:
            mask_new_size = mask_orig_size
        orig_total = orig_size + mask_orig_size
        key = min(cands, key=cands.get)
        # 只小才换
        if cands[key] + mask_new_size >= orig_total:
            stats["skip"] += 1
            continue

        smask_new_ref = smask_ref
        if mask_out is not None:
            smask_new_ref = _make_smask(doc, mask_out)
        if key == "png":
            # PNG 胜出：就地改流与字典，不新建 XObject——replace_image()
            # 内部 insert_image() 会留下未绘制的副本（实测每页多出孤儿图）
            doc.update_stream(x, png_raw, compress=True)
            doc.xref_set_key(x, "Width", str(w2))
            doc.xref_set_key(x, "Height", str(h2))
            doc.xref_set_key(x, "ColorSpace", cs_out)
            doc.xref_set_key(x, "BitsPerComponent", "8")
            # Pixmap 已应用 Decode/DecodeParms，替换后不能再套一次
            doc.xref_set_key(x, "Decode", "null")
            doc.xref_set_key(x, "DecodeParms", "null")
            doc.xref_set_key(x, "SMask", smask_new_ref or "null")
            stats["png"] += 1
        else:
            if smask_new_ref:
                doc.xref_set_key(x, "SMask", smask_new_ref)
            # JPEG 补丁阶段会重写字典，这里先把原颜色空间定好
            doc.xref_set_key(x, "ColorSpace", cs_out)
            jpg_wins.append((x, pno, w, h, jpg, w2, h2, orig_hash, mode))
            stats["jpg"] += 1
    print(f"阶段 1 候选（{time.time()-t_start:.1f}s）: {stats} "
          f"jpeg 胜出={len(jpg_wins)}")
    if stats["skip_mask"] or stats["skip_cs"]:
        print(f"  跳过：蒙版/着色模板 {stats['skip_mask']} 张，"
              f"色彩空间不支持 {stats['skip_cs']} 张")

    STAGE1 = os.path.join(WORK, "stage1.pdf")
    doc.save(STAGE1, garbage=4, deflate=True, use_objstms=True)
    doc.close()

    # ============================================================
    # 阶段 2：原始字节手术 —— 拼接 JPEG、重建 XRef
    # ============================================================
    if args.compress_pages:
        PRE = os.path.join(WORK, "prepages.pdf")
    else:
        PRE = OUT
    if jpg_wins:
        stage2(PDF, PRE, STAGE1, jpg_wins)
    else:
        os.replace(STAGE1, PRE)

    if args.compress_pages:
        # 阶段 3：在图片压缩后的文件上做整页 PNG 压缩
        # （PRE 被该调用消费）
        run_page_compress(PRE, OUT, args, pngquant)
    else:
        # 验证：页数一致 + 每页文字/图像都能读
        dv = pymupdf.open(OUT)
        bad = 0
        for p in range(dv.page_count):
            try:
                dv[p].get_text()
                dv[p].get_images()
            except Exception:
                bad += 1
        ok = dv.page_count == n_pages and bad == 0
        print(f"重开校验: 页数 = {dv.page_count}，不可读页 = {bad}"
              + ("" if ok else f"（不一致：输入有 {n_pages} 页）"))
        dv.close()
        if not ok:
            sys.exit("输出校验未通过")

    sz_orig = os.path.getsize(PDF)
    sz = os.path.getsize(OUT)
    print(f"\noriginal: {sz_orig/1024/1024:.2f} MB")
    print(f"输出: {sz/1024/1024:.2f} MB "
          f"({(sz_orig-sz)/1024/1024:.2f} MB saved, "
          f"{(sz_orig-sz)/sz_orig*100:.1f}%)")
    print(f"完成，用时 {time.time()-t_start:.1f}s；候选保留在 {WORK}")


def stage2(PDF, OUT, STAGE1, jpg_wins):
    # ---- 旧 xref -> 新 xref 的映射（保存会重新编号对象） ----
    d2 = pymupdf.open(STAGE1)
    remap = {}
    used_new = set()
    for old, pno, w, h, jpg, w2, h2, old_hash, mode in jpg_wins:
        # 在 stage1 的该页上找这张图的新 xref：JPEG 胜者在阶段 1 没被
        # 动过，流内容不变（哈希相同），尺寸仍是原始的 w×h
        found = None
        for img in d2[pno].get_images(full=True):
            xref = img[0]
            if xref in used_new:
                continue
            if img[2] == w and img[3] == h:
                h1 = hashlib.md5(d2.xref_stream_raw(xref)).digest()
                if h1 == old_hash:
                    found = xref
                    break
        if found is None:
            # 兜底：只按尺寸匹配
            for img in d2[pno].get_images(full=True):
                xref = img[0]
                if xref in used_new:
                    continue
                if img[2] == w and img[3] == h:
                    found = xref
                    break
        if found is None:
            print(f"  !! 保存后找不到 xref {old}（第 {pno+1} 页）")
            continue
        used_new.add(found)
        remap[old] = (found, jpg, w2, h2, mode)
    d2.close()
    print(f"映射 JPEG 胜者: {len(remap)}")

    # ---- 只读映射 stage1；输出流式写出，避免 GB 级文件反复整块复制 ----
    fsrc = open(STAGE1, "rb")
    data = mmap.mmap(fsrc.fileno(), 0, access=mmap.ACCESS_READ)
    FSIZE = len(data)

    def fallback(msg):
        """布局无法解析时退回纯 PyMuPDF 输出：丢掉 JPEG 收益，但文件有效。"""
        print(f"  !! {msg}；回退为纯 PyMuPDF 输出")
        try:
            # 映射还在时不能 os.replace(STAGE1, OUT)（Windows 会拒绝），
            # 直接把完整内容写到 OUT
            with open(OUT, "wb") as fout:
                fout.write(data)
        finally:
            try:
                data.close()
            except BufferError:
                # 仍有 re.Match 引用映射；留给进程退出时回收
                pass
            fsrc.close()

    def ofield(ent, i, widths):
        """按 /W 的字段宽度解析 XRef 条目字段（i=0 类型，1 偏移/对象流，2 代/索引）。"""
        start = 0 if i == 0 else sum(widths[:i])
        width = widths[i]
        if width == 0:
            return 0
        return int.from_bytes(ent[start:start + width], "big")

    # ---- 定位 XRef 流对象（总在文件尾部） ----
    tm = re.search(rb"startxref\s+(\d+)", data[-2048:])
    if not tm:
        fallback("找不到 startxref")
        return
    xref_off = int(tm.group(1))
    send = data.find(b"stream", xref_off)
    if send < 0 or send - xref_off > 65536:
        fallback("找不到 XRef 流对象")
        return
    m = re.match(rb"(\d+)\s+0\s+obj\s*<<(.*?)>>\s*stream",
                 data[xref_off:send + 6], re.S)
    if not m:
        fallback("XRef 流对象字典异常")
        return
    xrefnum = int(m.group(1))
    xdict = m.group(2)
    sp = send + 6
    if data[sp:sp + 2] == b"\r\n":
        sp += 2
    elif data[sp:sp + 1] in (b"\n", b"\r"):
        sp += 1

    # ---- 逐个图像对象打补丁 ----
    def png_unfilter(raw, columns):
        # XRef 流可能带 PNG 预测（/Predictor >= 10）：逐行反预测
        rowsz = columns + 1
        out = bytearray()
        prev = bytearray(columns)
        for r in range(len(raw) // rowsz):
            row = raw[r * rowsz:(r + 1) * rowsz]
            ft = row[0]
            cur = bytearray(row[1:])
            for i in range(columns):
                a = cur[i - 1] if i > 0 else 0
                b = prev[i]
                c = prev[i - 1] if i > 0 else 0
                if ft == 1:
                    cur[i] = (cur[i] + a) & 255
                elif ft == 2:
                    cur[i] = (cur[i] + b) & 255
                elif ft == 3:
                    cur[i] = (cur[i] + (a + b) // 2) & 255
                elif ft == 4:
                    p = a + b - c
                    pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                    pr = a if (pa <= pb and pa <= pc) else (b if pb <= pc else c)
                    cur[i] = (cur[i] + pr) & 255
            out += cur
            prev = cur
        return bytes(out)

    def obj_marker_ok(n, p):
        # 对象起点："N 0 obj" + 空白或 '<'（兼容标量对象与不换行的写入器）
        tail = data[p:p + 40]
        return re.match(rb"%d\s+0\s+obj[\s<]" % n, tail) is not None

    obj_at = {}
    for mm in re.finditer(rb"(?m)^(\d+)\s+0\s+obj\b", data):
        n = int(mm.group(1))
        if obj_marker_ok(n, mm.start()):
            obj_at.setdefault(n, mm.start())

    # 读 XRef 流拿到原始条目（扫不到的对象按原条目保留）
    orig_entries = {}
    Size = 0
    W = None
    sm = re.search(rb"/Size\s+(\d+)", xdict)
    wm = re.search(rb"/W\s*\[([^\]]*)\]", xdict)
    if sm and wm:
        Size = int(sm.group(1))
        W = [int(v) for v in wm.group(1).split()]
        wsum = sum(W)
        if wsum > 0:
            lm = re.search(rb"/Length\s+(\d+)(\s+0\s+R)?", xdict)
            ln = None
            if lm:
                ln = int(lm.group(1))
                if lm.group(2):
                    # 间接 /Length：解引用
                    o2 = obj_at.get(ln)
                    lm2 = None
                    if o2 is not None:
                        lm2 = re.search(rb"/Length\s+(\d+)",
                                        data[o2:o2 + 300])
                    ln = int(lm2.group(1)) if lm2 else None
            if ln is None:
                ln = data.find(b"endstream", sp) - sp
            rawx = data[sp:sp + ln]
            if b"/Filter" in xdict:
                try:
                    rawx = zlib.decompress(rawx)
                except zlib.error:
                    fallback("XRef 流解压失败")
                    return
                # PNG 预测（MuPDF 常写 /Predictor 12 /Columns N）
                pm = re.search(rb"/Predictor\s+(\d+)", xdict)
                cm = re.search(rb"/Columns\s+(\d+)", xdict)
                if pm and cm and int(pm.group(1)) >= 10:
                    rawx = png_unfilter(rawx, int(cm.group(1)))
            # /Index 可以是多段；不能假设一整段 0..Size
            idxm = re.search(rb"/Index\s*\[([^\]]*)\]", xdict)
            if idxm:
                vals = [int(v) for v in idxm.group(1).split()]
                pairs = [(vals[i], vals[i + 1])
                         for i in range(0, len(vals) - 1, 2)]
            else:
                pairs = [(0, Size)]
            pos = 0
            for start, count in pairs:
                for n in range(start, start + count):
                    if pos + wsum > len(rawx):
                        break
                    orig_entries[n] = bytes(rawx[pos:pos + wsum])
                    pos += wsum
            # XRef 里的 type-1 偏移优先于扫描猜测
            for n, ent in orig_entries.items():
                if len(ent) >= wsum and ofield(ent, 0, W) == 1:
                    off = ofield(ent, 1, W)
                    if off < FSIZE and obj_marker_ok(n, off):
                        obj_at[n] = off
    if W is None:
        W = [1, 4, 1]

    # ---- 对象流：哪些对象住在对象流里 ----
    in_objstm = {}  # n -> (流 xref, 索引)
    objstm_fail = 0
    for n, off in list(obj_at.items()):
        seg = data[off:off + 1024]
        if not re.search(rb"/Type\s*/ObjStm", seg):
            continue
        s = data.find(b"stream", off)
        if s < 0 or s - off > 2048:
            continue
        dicttext = data[off:s]
        lm = re.search(rb"/Length\s+(\d+)(\s+0\s+R)?", dicttext)
        if not lm:
            continue
        l = int(lm.group(1))
        if lm.group(2):  # 间接 /Length：解引用
            o2 = obj_at.get(l)
            if o2 is None:
                continue
            lm2 = re.search(rb"/Length\s+(\d+)", data[o2:o2 + 300])
            if not lm2:
                continue
            l = int(lm2.group(1))
        p = s + 6
        if data[p:p + 2] == b"\r\n":
            p += 2
        elif data[p:p + 1] in (b"\n", b"\r"):
            p += 1
        try:
            dec = zlib.decompress(data[p:p + l])
        except zlib.error:
            objstm_fail += 1
            continue
        pm = re.search(rb"/Predictor\s+(\d+)", dicttext)
        cm = re.search(rb"/Columns\s+(\d+)", dicttext)
        if pm and cm and int(pm.group(1)) >= 10:
            dec = png_unfilter(dec, int(cm.group(1)))
        mN = re.search(rb"/N\s+(\d+)", dicttext)
        mF = re.search(rb"/First\s+(\d+)", dicttext)
        if not mN or not mF:
            continue
        n_obj = int(mN.group(1))
        first = int(mF.group(1))
        ints = [int(x) for x in re.findall(rb"\d+", dec[:first])]
        if len(ints) == n_obj * 3:
            # 标准头：(编号, 代, 偏移) x N
            seq = [ints[i * 3] for i in range(n_obj)]
        elif len(ints) == 2 + (n_obj - 1) * 2:
            # MuPDF 紧凑头
            seq = [ints[i * 2] for i in range(n_obj)]
        else:
            continue
        for idx, num in enumerate(seq):
            in_objstm[num] = (n, idx)
    if objstm_fail:
        print(f"  !! {objstm_fail} 个对象流解压失败；"
              f"其中对象按原 XRef 条目保留")
    missing = [n for n in range(1, Size)
               if n not in obj_at and n not in in_objstm
               and n not in orig_entries]
    if missing:
        print(f"  !! 对象 {len(missing)} 个既未定位也不在原 XRef 中；"
              f"标记为 free")
    print(f"定位对象: 顶层 {len(obj_at)}，对象流内 {len(in_objstm)}")

    # ---- 构建图像补丁 ----
    def resolve_length(dicttext):
        lm = re.search(rb"/Length\s+(\d+)(\s+0\s+R)?", dicttext)
        if not lm:
            return None
        v = int(lm.group(1))
        if lm.group(2):
            # 间接 /Length：解引用
            o2 = obj_at.get(v)
            if o2 is None:
                return None
            lm2 = re.search(rb"/Length\s+(\d+)", data[o2:o2 + 300])
            return int(lm2.group(1)) if lm2 else None
        return v

    patches = []
    for old, (newx, jpg, w2, h2, mode) in remap.items():
        off = obj_at.get(newx)
        if off is None:
            print(f"  !! xref {newx} 缺失；跳过补丁")
            continue
        send = data.find(b"stream", off)
        if send < 0 or send - off > 65536:
            print(f"  !! xref {newx}: 找不到流；跳过补丁")
            continue
        dicttext = data[off:send]
        jpgb = open(jpg, "rb").read()
        # 裸存（DCTDecode）或 Flate 包裹（[/FlateDecode /DCTDecode]），取小者
        flb = zlib.compress(jpgb, 6)
        if len(flb) < len(jpgb):
            fkey = "/Filter[/FlateDecode /DCTDecode]"
            store = flb
        else:
            fkey = "/Filter /DCTDecode"
            store = jpgb
        # 原字典里的透明度与渲染意图要带上
        extra = ""
        sm = re.search(rb"/SMask\s+(\d+)\s+0\s+R", dicttext)
        if sm:
            extra += f" /SMask {sm.group(1).decode()} 0 R"
        it = re.search(rb"/Intent\s*(/\w+)", dicttext)
        if it:
            extra += f" /Intent {it.group(1).decode()}"
        # 复用阶段 1 定好的颜色空间（可能是 ICCBased 引用）
        csm = re.search(
            rb"/ColorSpace\s*(/\w+|\[[^\]]*\]|\d+\s+\d+\s+R)", dicttext)
        want_cs = {"L": "/DeviceGray", "RGB": "/DeviceRGB",
                   "CMYK": "/DeviceCMYK"}[mode]
        cs = csm.group(1).decode("latin-1") if csm else want_cs
        if (cs in ("/DeviceRGB", "/DeviceGray", "/DeviceCMYK")
                and cs != want_cs):
            cs = want_cs
        # Pillow 写出的 CMYK JPEG 是反相存储，必须补 /Decode
        dec = " /Decode [1 0 1 0 1 0 1 0]" if mode == "CMYK" else ""
        ndict = (f"<< /Type /XObject /Subtype /Image /Width {w2} "
                 f"/Height {h2} {fkey} /ColorSpace {cs} "
                 f"/BitsPerComponent 8{dec}{extra}/Length {len(store)} >>")
        newobj = (f"{newx} 0 obj\n{ndict}\nstream\n").encode() + store \
            + b"\nendstream\nendobj"
        p = send + 6
        if data[p:p + 2] == b"\r\n":
            p += 2
        elif data[p:p + 1] in (b"\n", b"\r"):
            p += 1
        si = p
        Lold = resolve_length(dicttext)
        if Lold is None:
            print(f"  !! xref {newx}: 取不到 /Length；跳过补丁")
            continue
        e = si + Lold
        m2 = re.match(rb"[\r\n]*endstream[\r\n]*endobj", data[e:e + 64])
        if not m2:
            print(f"  !! xref {newx}: 对象结尾异常；跳过补丁")
            continue
        patches.append((off, e + m2.end(), newobj))

    # 每个对象的最终偏移（补丁造成的位移）
    def shifted(off):
        o = off
        for poff, pend, pnew in patches:
            if poff < off:
                o += len(pnew) - (pend - poff)
        return o

    final = {n: shifted(off) for n, off in obj_at.items()}
    print(f"应用图像补丁: {len(patches)}")

    # ---- 重建 XRef 流 ----
    # 读旧条目一律按旧 /W 解析；写出时按新宽度编码，避免 /W 变化后错位
    max_gen = 0
    max_idx = max((idx for _, idx in in_objstm.values()), default=0)
    for ent0 in orig_entries.values():
        if len(ent0) < sum(W):
            continue
        t = ofield(ent0, 0, W)
        if t == 1:
            max_gen = max(max_gen, ofield(ent0, 2, W))
        elif t == 2:
            max_idx = max(max_idx, ofield(ent0, 2, W))
    third = max(1, (max(max_gen, max_idx).bit_length() + 7) // 8)
    Wnew = [1, 4, third]

    maxobj = max(Size, max(final, default=0) + 1, xrefnum + 1)
    ent = bytearray()

    def add_entry(t, f1=0, f2=0):
        ent.extend(t.to_bytes(1, "big"))
        ent.extend(f1.to_bytes(4, "big"))
        ent.extend(f2.to_bytes(third, "big"))

    for n in range(maxobj):
        if n == 0:
            add_entry(0)
        elif n in final:
            add_entry(1, final[n], 0)
        elif n in in_objstm:
            add_entry(2, in_objstm[n][0], in_objstm[n][1])
        elif n in orig_entries and len(orig_entries[n]) >= sum(W):
            oe = orig_entries[n]
            t = ofield(oe, 0, W)
            if t == 1:
                off = ofield(oe, 1, W)
                if off < FSIZE:
                    # 没定位到但原 XRef 有记录：按补丁位移保留，别标 free
                    add_entry(1, shifted(off), ofield(oe, 2, W))
                else:
                    add_entry(0)
            elif t == 2:
                add_entry(2, ofield(oe, 1, W), ofield(oe, 2, W))
            else:
                add_entry(0)
        else:
            add_entry(0)

    rm = re.search(rb"/Root\s+(\d+\s+0\s+R)", xdict)
    if not rm:
        fallback("XRef 流里没有 /Root")
        return
    root = rm.group(1)
    info = re.search(rb"/Info\s+(\d+\s+0\s+R)", xdict)
    info = info.group(1) if info else b"null"
    idm = re.search(rb"/ID\s*\[(.*?)\]", xdict, re.S)
    idtxt = idm.group(1) if idm else b"[]"
    # 条目表不压缩（字典里不声明 /Filter，数据即原样条目字节）
    new_raw = bytes(ent)
    new_dict = (b"<</Type/XRef/W[%d %d %d]/Index[0 %d]/Size %d"
                b"/Root %s/Info %s/ID[%s]/Length %d>>"
                % (Wnew[0], Wnew[1], Wnew[2], maxobj, maxobj,
                   root, info, idtxt, len(new_raw)))
    xref_src_off = obj_at[xrefnum]   # 原文件里的偏移（用于从 mmap 取字节）
    tail_start = final[xrefnum]      # 补丁后的新偏移（写进 startxref）
    new_tail = (b"%d 0 obj\n%s\nstream\n" % (xrefnum, new_dict)) + new_raw \
        + b"\nendstream\nendobj\n\nstartxref\n%d\n%%EOF\n" % tail_start

    # 先写同目录临时文件再原子替换，避免中途失败留下半个输出
    tmp = OUT + ".tmp"
    try:
        with open(tmp, "wb") as fout:
            prev = 0
            for off, obj_end, newobj in sorted(patches):
                fout.write(data[prev:off])
                fout.write(newobj)
                prev = obj_end
            fout.write(data[prev:xref_src_off])
            fout.write(new_tail)
        os.replace(tmp, OUT)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        try:
            data.close()
        except BufferError:
            # 仍有 re.Match 引用映射；留给进程退出时回收
            pass
        fsrc.close()


if __name__ == "__main__":
    main()
