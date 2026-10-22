# -*- coding: utf-8 -*-
"""
docx2pdf.py —— 使用本机 Microsoft Word 把 .doc/.docx 批量导出为 PDF

用法：
    python docx2pdf.py                  # 命令行无目录 -> 转换【当前工作目录】
    python docx2pdf.py "D:\某目录"       # 命令行跟目录 -> 转换【该目录】
    python docx2pdf.py -r               # 递归处理子文件夹
    python docx2pdf.py -f               # 覆盖已存在的 PDF
    python docx2pdf.py "D:\某目录" -r -f

注意：
    1. 输出 PDF 始终与源文档同目录同名。
    2. 目录参数支持相对路径与 ~（如 ~ 或 ~\\Documents\\图纸）。
"""

import argparse
import os
import sys
import time

import pythoncom
import win32com.client as win32

# ---- Word 常量 ----
WD_FORMAT_PDF = 17                 # wdFormatPDF
WD_DO_NOT_SAVE_CHANGES = 0         # wdDoNotSaveChanges
WD_ALERTS_NONE = 0                 # wdAlertsNone
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0   # wdExportOptimizeForPrint
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

DOC_EXTS = (".doc", ".docx", ".docm", ".dot", ".dotx")


# ------------------------------------------------------------------ 工具函数
def _init_console():
    """输出里出现无法编码的字符时不要抛异常。"""
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass


def collect_files(base_dir, recursive=False):
    """收集目录下待转换的 Word 文档（跳过 Word 临时文件 ~$xxx）。"""
    files = []
    if recursive:
        for root, dirs, names in os.walk(base_dir):
            dirs[:] = [d for d in dirs if not d.startswith("~$")]
            for n in names:
                if os.path.splitext(n)[1].lower() in DOC_EXTS and not n.startswith("~$"):
                    files.append(os.path.join(root, n))
    else:
        for n in os.listdir(base_dir):
            full = os.path.join(base_dir, n)
            if os.path.isfile(full) \
                    and os.path.splitext(n)[1].lower() in DOC_EXTS \
                    and not n.startswith("~$"):
                files.append(full)
    files.sort(key=lambda p: p.lower())
    return files


def start_word():
    """启动一个隐藏的 Word 实例。"""
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = WD_ALERTS_NONE
    # 禁用宏、关闭“打印前警告保存标记”等干扰性提示
    try:
        word.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
    except Exception:
        pass
    try:
        word.Options.WarnBeforeSavingPrintingSendingMarkup = False
    except Exception:
        pass
    try:
        word.Options.SaveNormalPrompt = False
    except Exception:
        pass
    return word


def export_pdf(word, docx_path, pdf_path):
    """打开 docx 并导出为 pdf，失败时回退到 SaveAs2。"""
    doc = None
    try:
        doc = word.Documents.Open(
            FileName=docx_path,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        try:
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,          # 必须是绝对路径！
                ExportFormat=WD_FORMAT_PDF,
                OpenAfterExport=False,
                OptimizeFor=WD_EXPORT_OPTIMIZE_FOR_PRINT,
                CreateBookmarks=WD_EXPORT_CREATE_HEADING_BOOKMARKS,
            )
        except Exception:
            # 部分旧版 Word / 特殊文档走 SaveAs 更稳
            doc.SaveAs2(FileName=pdf_path, FileFormat=WD_FORMAT_PDF)
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass


def describe_error(err):
    """把 COM 报错翻译成人话。"""
    text = str(err)
    low = text.lower()
    if "password" in low or "密码" in text:
        return "文档设有打开密码，请先去掉密码"
    if "could not open" in low or "无法打开" in text:
        return "文件损坏、被占用或被加密，Word 无法打开"
    if "read-only" in low or "只读" in text:
        return "文件为只读或无写入权限"
    return text


# ------------------------------------------------------------------ 主流程
def main(argv=None):
    _init_console()

    ap = argparse.ArgumentParser(
        description="使用 Microsoft Word 批量把 doc/docx 转成 PDF。"
                    "命令行跟目录则转换该目录，不跟目录则转换当前工作目录。",
        epilog="示例：\n"
               "  python docx2pdf.py              转换当前目录\n"
               "  python docx2pdf.py D:\\图纸       转换 D:\\图纸 目录\n"
               "  python docx2pdf.py D:\\图纸 -r    递归转换 D:\\图纸 及子目录\n"
               "  python docx2pdf.py -f            覆盖当前目录已存在的 PDF",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("folder", nargs="?", default=".",
                    help="要转换的目录；省略则使用当前工作目录")
    ap.add_argument("-r", "--recursive", action="store_true", help="递归处理子文件夹")
    ap.add_argument("-f", "--force", action="store_true", help="覆盖已存在的 PDF")
    args = ap.parse_args(argv)

    # 省略目录 -> 当前工作目录；给了目录 -> 展开为绝对路径（兼容相对路径、"~"）
    base_dir = os.path.abspath(os.path.expanduser(args.folder))
    print("=== Word 批量转 PDF ===")
    print("源目录：", base_dir)
    print("输出到：与源文件同目录（同名 .pdf）")

    if not os.path.isdir(base_dir):
        print("【错误】目录不存在：", base_dir)
        return 2

    files = collect_files(base_dir, args.recursive)
    if not files:
        print("【提示】未找到 .doc/.docx 文件。")
        return 0
    print("待处理：", len(files), "个文件\n")

    pythoncom.CoInitialize()
    word = None
    word_dead = False   # Word 重启失败后，后续文件不再徒劳重试
    ok_list, skip_list, fail_list = [], [], []

    try:
        try:
            word = start_word()
        except Exception as e:
            print("【错误】无法启动 Microsoft Word。请确认已安装 Word 且已安装 pywin32。")
            print("详细错误：", e)
            return 3

        for i, src in enumerate(files, 1):
            src = os.path.abspath(src)
            pdf = os.path.join(os.path.dirname(src),                       # ★ 绝对路径
                               os.path.splitext(os.path.basename(src))[0] + ".pdf")
            name = os.path.basename(src)

            if os.path.exists(pdf) and not args.force:
                print("[%2d/%2d] 跳过（已存在）：%s" % (i, len(files), name))
                skip_list.append(name)
                continue

            if word_dead:
                fail_list.append((name, "Word 重启失败，未处理"))
                continue

            print("[%2d/%2d] 转换：%s" % (i, len(files), name))
            for attempt in (1, 2):
                try:
                    if os.path.exists(pdf):
                        os.remove(pdf)          # 覆盖前先删，避免残留旧页
                    export_pdf(word, src, pdf)
                    if not (os.path.exists(pdf) and os.path.getsize(pdf) > 0):
                        raise IOError("Word 未生成有效的 PDF 文件")
                    ok_list.append(name)
                    print("        -> %s（%.1f KB）"
                          % (os.path.basename(pdf), os.path.getsize(pdf) / 1024.0))
                    break
                except Exception as e:
                    if attempt == 2:
                        # Word 可能已写出 0 字节或不完整的 PDF，留着会让下次运行
                        # 误判为「已存在」而跳过，所以失败后清理掉
                        try:
                            if os.path.exists(pdf):
                                os.remove(pdf)
                        except OSError:
                            pass
                        print("        【失败】", describe_error(e))
                        fail_list.append((name, describe_error(e)))
                    else:
                        time.sleep(1.0)        # 文件被占用时重试一次
                        try:
                            word.Quit()
                        except Exception:
                            pass
                        word = None
                        try:
                            word = start_word()
                        except Exception as restart_err:
                            why = describe_error(restart_err)
                            print("        【失败】无法重新启动 Word：", why)
                            try:
                                if os.path.exists(pdf):
                                    os.remove(pdf)
                            except OSError:
                                pass
                            fail_list.append((name, why))
                            word_dead = True
                            break
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    print("\n=== 转换完成 ===")
    print("成功：%d    跳过：%d    失败：%d" % (len(ok_list), len(skip_list), len(fail_list)))
    if fail_list:
        print("\n失败明细：")
        for name, why in fail_list:
            print("  - %s：%s" % (name, why))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())