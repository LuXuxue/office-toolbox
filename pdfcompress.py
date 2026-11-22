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
import os
import re
import subprocess
import sys
import time
import zlib

import pymupdf
from PIL import Image


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

    stats = {"png": 0, "jpg": 0, "skip": 0, "fail": 0}
    jpg_wins = []   # (旧 xref, 页码, w, h, jpg 路径, w2, h2, 原流哈希)
    touched = set()
    replaced = set()
    for x, im in images.items():
        pno, w, h = im["page"], im["w"], im["h"]
        # 软掩码（透明度）是独立对象；记下引用，替换后写回
        _, smask_t = doc.xref_get_key(x, "SMask")
        smask_ref = smask_t if smask_t != "null" else None
        # 由显示矩形推算实际 DPI
        disp_w_in = im["rect"].width / 72.0
        dpi = w / disp_w_in if disp_w_in > 0 else args.dpi
        if dpi > args.dpi:
            scale = args.dpi / dpi
            w2, h2 = max(1, round(w * scale)), max(1, round(h * scale))
        else:
            w2, h2 = w, h
        # 稳健取像素
        pix = pymupdf.Pixmap(doc, x)
        if pix.alpha or pix.n - pix.alpha > 3:
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        if (w2, h2) != (w, h):
            img = img.resize((w2, h2), Image.LANCZOS)
        base = os.path.join(WORK, f"{x}")
        # 候选 1：pngquant PNG（原地量化，skip-if-larger）
        png = base + ".png"
        img.save(png)
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
        cands = {"png": os.path.getsize(png)}
        # PNG 入 PDF 后是 Flate(裸像素)；PNG 文件有扫描线滤波，盘上
        # 通常比入 PDF 小，这里按裸像素 zlib 流估计入 PDF 体积
        cands["png"] = max(cands["png"],
                           len(zlib.compress(img.tobytes("raw"), 6)))
        # 候选 2：JPEG
        jpg = base + ".jpg"
        img.save(jpg, format="JPEG", quality=args.jpeg_quality,
                 subsampling=args.jpeg_subsample)
        # JPEG 入 PDF 后是裸 DCTDecode 或 Flate 包裹，取小者
        jpg_bytes = open(jpg, "rb").read()
        cands["jpg"] = min(len(jpg_bytes),
                           len(zlib.compress(jpg_bytes, 6)))
        # 原流体积（判断“只小不大”的基准）
        try:
            orig_raw = doc.xref_stream_raw(x)
        except Exception:
            orig_raw = doc.xref_stream(x)
        orig_size = len(orig_raw)
        orig_hash = hashlib.md5(orig_raw).digest()
        key = min(cands, key=cands.get)
        # 只小才换
        if cands[key] >= orig_size:
            stats["skip"] += 1
            continue
        if key == "png":
            # PNG 胜出：内存中替换（MuPDF 保存时会重编码为裸像素 Flate）
            doc[pno].replace_image(x, filename=png)
            # replace_image 会清掉 /SMask，写回原引用
            if smask_ref:
                doc.xref_set_key(x, "SMask", smask_ref)
            stats["png"] += 1
            touched.add(pno)
            replaced.add(x)
        else:
            jpg_wins.append((x, pno, w, h, base + ".jpg",
                             w2, h2, orig_hash))
            stats["jpg"] += 1
    print(f"阶段 1 候选（{time.time()-t_start:.1f}s）: {stats} "
          f"jpeg 胜出={len(jpg_wins)}")

    # 死重复清理（保险措施：只清理本次替换过、且内容流已不再引用的
    # 图像条目；replace_image 是就地替换，通常不会留下孤儿）
    removed = 0
    for pno in sorted(touched):
        page = doc[pno]
        a, b = doc.xref_get_key(page.xref, "Resources/XObject")
        if not isinstance(b, str) or not b.startswith("<<"):
            continue
        names = {m.group(1): int(m.group(2))
                 for m in re.finditer(rb"/(\w+)\s+(\d+)\s+0\s+R", b.encode())}
        cs = b""
        for c in page.get_contents():
            cs += doc.xref_stream(c)
        used = {m.group(1) for m in re.finditer(rb"/(\w+)\s+Do", cs)}
        for name, xref in list(names.items()):
            if name not in used and xref in replaced:
                page.delete_image(xref)
                removed += 1
    print(f"死重复清理: {removed}")

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
        # 验证
        dv = pymupdf.open(OUT)
        ok = dv.page_count == n_pages
        print(f"重开校验: 页数 = {dv.page_count}"
              + ("" if ok else f"（不一致：输入有 {n_pages} 页）"))
        dv.close()

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
    for old, pno, w, h, jpg, w2, h2, old_hash in jpg_wins:
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
        remap[old] = (found, jpg, w2, h2)
    d2.close()
    print(f"映射 JPEG 胜者: {len(remap)}")

    # ---- 读 stage1 保存出的文件（原始字节手术的对象） ----
    data = open(STAGE1, "rb").read()
    FSIZE = len(data)

    # ---- 定位 XRef 流对象（总在文件尾部） ----
    i = data.rfind(b"startxref")
    j = data.find(b"\n", i + 9) + 1          # “startxref” 行末
    k = data.find(b"\n", j)                  # 偏移行行末
    xref_off = int(data[j:k].strip())
    m = re.match(rb"(\d+) 0 obj\r?\n<<(.*?)>>\r?\nstream\r?\n",
                 data[xref_off:], re.S)
    if not m:
        sys.exit("stage2: 找不到 XRef 流对象（不支持的布局）")
    xrefnum = int(m.group(1))
    xdict = m.group(2)

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
        # 对象起点："N 0 obj" + 空白（含标量对象，如 46 0 obj\n26444\nendobj）
        tail = data[p:p + 40]
        mm = re.match(rb"%d 0 obj[\r\n ]" % n, tail)
        return mm is not None

    obj_at = {}
    for mm in re.finditer(rb"(\d+) 0 obj\r?\n", data):
        n = int(mm.group(1))
        if obj_marker_ok(n, mm.start()):
            obj_at.setdefault(n, mm.start())

    # 读 XRef 流拿到原始条目（扫不到的对象用原偏移）
    xent_off = obj_at.get(xrefnum)
    orig_entries = {}
    Size = 0
    if xent_off is not None:
        seg = data[xent_off:xent_off + 400]
        Size = int(re.search(rb"/Size\s+(\d+)", seg).group(1))
        sm = re.search(rb"stream\r?\n", data[xent_off:xent_off + 400])
        sp = xent_off + sm.end()
        lm = re.search(rb"/Length\s+(\d+)", seg)
        ln = int(lm.group(1))
        rawx = data[sp:sp + ln]
        if b"/Filter" in seg:
            rawx = zlib.decompress(rawx)
            # PNG 预测（MuPDF 常写 /Predictor 12 /Columns N）
            pm = re.search(rb"/Predictor\s+(\d+)", seg)
            cm = re.search(rb"/Columns\s+(\d+)", seg)
            if pm and cm and int(pm.group(1)) >= 10:
                rawx = png_unfilter(rawx, int(cm.group(1)))
        wm = re.search(rb"/W\s*\[([^\]]*)\]", seg)
        W = [int(v) for v in wm.group(1).split()]
        wsum = sum(W)
        for n in range(1, Size):
            ep = n * wsum
            if ep + wsum > len(rawx):
                break
            orig_entries[n] = bytes(rawx[ep:ep + wsum])
        # XRef 里的 type-1 偏移优先于扫描猜测
        for n, ent in orig_entries.items():
            if ent[0] == 1:
                off = int.from_bytes(ent[1:1 + W[1]], "big")
                if off < FSIZE and obj_marker_ok(n, off):
                    obj_at[n] = off

    # ---- 对象流：哪些对象住在对象流里 ----
    in_objstm = {}  # n -> (流 xref, 索引)
    for n, off in list(obj_at.items()):
        seg = data[off:off + 300]
        if not re.search(rb"/Type\s*/ObjStm", seg):
            continue
        s = data.find(b"stream", off)
        if s < 0 or s - off > 500:
            continue
        lm = re.search(rb"/Length\s+(\d+)(\s+0\s+R)?", data[off:s])
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
        eol = 1 if data[s + 6:s + 7] == b"\n" else 2
        try:
            dec = zlib.decompress(data[s + 6 + eol:s + 6 + eol + l])
            pm = re.search(rb"/Predictor\s+(\d+)", seg)
            cm = re.search(rb"/Columns\s+(\d+)", seg)
            if pm and cm and int(pm.group(1)) >= 10:
                dec = png_unfilter(dec, int(cm.group(1)))
        except zlib.error:
            continue
        mN = re.search(rb"/N (\d+)", seg)
        mF = re.search(rb"/First (\d+)", seg)
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
    for n in range(1, Size):
        if n not in obj_at and n not in in_objstm:
            print(f"  !! 对象 {n} 确实缺失；标记为 free")
    print(f"定位对象: 顶层 {len(obj_at)}，对象流内 {len(in_objstm)}")

    # ---- 构建图像补丁 ----
    def stream_length(off):
        lm = re.search(rb"/Length\s+(\d+)(\s+0\s+R)?", data[off:off + 600])
        v = int(lm.group(1))
        if lm.group(2):
            # 间接 /Length：解引用
            o2 = obj_at.get(v)
            if o2 is not None:
                lm2 = re.search(rb"/Length\s+(\d+)", data[o2:o2 + 300])
                if lm2:
                    return int(lm2.group(1))
        return v

    patches = []
    for old, (newx, jpg, w2, h2) in remap.items():
        off = obj_at.get(newx)
        if off is None:
            print(f"  !! xref {newx} 缺失；跳过补丁")
            continue
        jpgb = open(jpg, "rb").read()
        # 裸存（DCTDecode）或 Flate 包裹（[/FlateDecode /DCTDecode]），取小者
        flb = zlib.compress(jpgb, 6)
        if len(flb) < len(jpgb):
            fkey = "/Filter[/FlateDecode /DCTDecode]"
            store = flb
        else:
            fkey = "/Filter /DCTDecode"
            store = jpgb
        # 原字典里的透明度键要带上
        dicttext = data[off:off + 600]
        extra = ""
        sm = re.search(rb"/SMask (\d+) 0 R", dicttext)
        if sm:
            extra += f" /SMask {sm.group(1).decode()} 0 R"
        mt = re.search(rb"/Matte \[([^\]]*)\]", dicttext)
        if mt:
            extra += f" /Matte [{mt.group(1).decode()}]"
        ndict = (f"<< /Type /XObject /Subtype /Image /Width {w2} /Height {h2} "
                 f"{fkey} /ColorSpace /DeviceRGB "
                 f"/BitsPerComponent 8{extra}/Length {len(store)} >>")
        newobj = (f"{newx} 0 obj\n{ndict}\nstream\n").encode() + store \
            + b"\nendstream\nendobj"
        sm = re.search(rb"stream\r?\n", data[off:off + 600])
        si = off + sm.end()
        Lold = stream_length(off)
        e = si + Lold
        m2 = re.match(rb"[\r\n]*endstream[\r\n]*endobj", data[e:e + 40])
        if not m2:
            print(f"  !! xref {newx}: 对象结尾异常；跳过补丁")
            continue
        patches.append((off, e + m2.end(), newobj))

    # 每个对象的最终偏移（减去下方补丁造成的位移）
    final = {}
    for n, off in obj_at.items():
        o = off
        for poff, pend, pnew in patches:
            if poff < off:
                o += len(pnew) - (pend - poff)
        final[n] = o
    for off, obj_end, newobj in sorted(patches, reverse=True):
        data = data[:off] + newobj + data[obj_end:]
    print(f"应用图像补丁: {len(patches)}")

    # ---- 重建 XRef 流 ----
    maxobj = max(max(final, default=0), xrefnum, Size) + 1
    ent = bytearray()
    for n in range(maxobj):
        if n == 0:
            ent += b"\x00" + (0).to_bytes(4, "big") + (0).to_bytes(1, "big")
        elif n in final:
            ent += b"\x01" + final[n].to_bytes(4, "big") + (0).to_bytes(1, "big")
        elif n in in_objstm:
            sx, idx = in_objstm[n]
            if idx < 256:
                ent += b"\x02" + sx.to_bytes(4, "big") + idx.to_bytes(1, "big")
            elif n in orig_entries and orig_entries[n][0] == 2:
                ent += orig_entries[n]
            else:
                ent += b"\x00" + (0).to_bytes(4, "big") + (0).to_bytes(1, "big")
        elif n in orig_entries:
            t = orig_entries[n][0]
            if t == 2:
                ent += orig_entries[n]
            elif t == 1:
                off = int.from_bytes(orig_entries[n][1:5], "big")
                if off < FSIZE and obj_marker_ok(n, off):
                    ent += orig_entries[n]
                else:
                    ent += b"\x00" + (0).to_bytes(4, "big") + (0).to_bytes(1, "big")
            else:
                ent += b"\x00" + (0).to_bytes(4, "big") + (0).to_bytes(1, "big")
        else:
            ent += b"\x00" + (0).to_bytes(4, "big") + (0).to_bytes(1, "big")
    root = re.search(rb"/Root (\d+ 0 R)", xdict).group(1)
    info = re.search(rb"/Info (\d+ 0 R)", xdict)
    info = info.group(1) if info else b"null"
    idm = re.search(rb"/ID\[(.*?)\]", xdict)
    idtxt = idm.group(1) if idm else b"[]"
    # 条目表不压缩（字典里不声明 /Filter，数据即原样条目字节）
    new_raw = bytes(ent)
    new_dict = (b"<</Type/XRef/W[1 4 1]/Index[0 %d]/Size %d/Root %s/Info %s"
                b"/ID[%s]/Length %d>>"
                % (maxobj, maxobj, root, info, idtxt, len(new_raw)))
    tail_start = final[xrefnum]
    new_tail = (b"%d 0 obj\n%s\nstream\n" % (xrefnum, new_dict)) + new_raw \
        + b"\nendstream\nendobj\n\nstartxref\n%d\n%%EOF\n" % tail_start
    data = data[:tail_start] + new_tail
    open(OUT, "wb").write(data)


if __name__ == "__main__":
    main()
