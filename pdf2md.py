#!/usr/bin/env python3
"""本地 PDF OCR：用 transformers 直接推理 OXR-1.0，输出 Markdown / JSON。

不依赖 vLLM、Docker 或 Linux，Windows 上可直接运行。脚本依赖同目录下的
OXR/ 子目录（从 SoMarkAI/OXR 复制的最小运行子集）：

    OXR/oxr/                        官方流水线（版面结果 -> md/json）
    OXR/services/layout_analyze/    D-FINE 版面分析推理代码
    OXR/models/OXR-1.0/             模型权重（不入库，需自行下载）

模型下载（二选一，约 3 GB）：

    pip install modelscope
    modelscope download SoMark/OXR-1.0 --local_dir ./OXR/models/OXR-1.0

    # 或者用 huggingface-cli
    huggingface-cli download SoMarkAI/OXR-1.0 --local-dir ./OXR/models/OXR-1.0

用法：

    # 单个 PDF
    python pdf2md.py D:\\docs\\a.pdf

    # 目录下所有 PDF
    python pdf2md.py -d D:\\docs

    # 递归处理 + 指定输出目录
    python pdf2md.py -d D:\\docs --recursive -o D:\\out

    # 只跑第 1~5 页做效果/速度验证
    python pdf2md.py D:\\docs\\a.pdf --pages 1-5

结果默认写到源 PDF 所在目录，文件名与源文件相同，默认只输出 .md
（加 --formats json 会同时输出 .json）；
PDF 里的图片导出到同目录的 <文件名>_media/。

速度相关的几个旋钮（默认值按“单卡、可中断”调过）：

    --batch-size      文本模型的批大小。流水线原本面向 vLLM 的连续批处理，
                      本地串行时把同一类型的识别请求拼成一批，能把每 token
                      都要重读一遍模型权重的开销摊薄。单步耗时主要由权重
                      带宽决定、几乎不随批大小增长（实测 8 与 16 持平、
                      20/24 没有净收益，4 明显慢 158 s），故默认 16。
    --chunk-pages     每批送进流水线的页数。它是本机最不能放松的一个：
                      显存峰值基本随它线性上涨，一旦逼近 12 GB 就开始抖动，
                      实测 16 页 chunk 8→38 s、chunk 12→129 s、chunk 16→90 s。
    --pages           页码范围（1 起），只在选中的页上工作。
    --layout-concurrency  版面分析的在途批数（单卡串行，一般不用改）。
    --graph-decode    解码段是否捕获成 CUDA Graph（auto/on/off，仅 NVIDIA）。
                      图按「批大小 × 缓存长度」分桶复用，小批用小图，免得
                      1 行请求按 16 行做一遍预填充。
    --graph-cache-len 图解码预留的缓存上限（默认：2048）。实际每批按
                      「prompt + 预留」取 256 的整数倍，桶切换时重新捕获。
                      超出的请求续跑一段或回退到普通生成。

内部还有三处不对外暴露的并行：PDF 按 chunk 懒渲染（渲染与上一块推理重叠，
不再整本前置渲染）、版面分析预取下一块、视觉塔跑在独立 CUDA 流上与下一批的
解码重叠（视觉受限算力、解码受限带宽，互相当填空；按 patch 切块以免两条流
的显存池叠加顶满 12 GB 卡）。

依赖（GPU 版 torch 请按 https://pytorch.org 的说明安装）：

    pip install -r requirements.txt
"""

from __future__ import annotations

import argparse
import asyncio
import contextlib
import importlib
import json
import logging
import os
import sys
import threading
import time
import types
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

BASE_DIR = Path(__file__).resolve().parent
OXR_ROOT = BASE_DIR / "OXR"
LAYOUT_SERVICE_DIR = OXR_ROOT / "services" / "layout_analyze"
DEFAULT_MODEL_DIR = OXR_ROOT / "models" / "OXR-1.0"

# 与官方 vLLM processor 保持一致（见 models/OXR-1.0/plugins/vllm/oxr_vllm/oxr_processor.py）。
MIN_PIXELS = 200704   # 448 * 448
MAX_PIXELS = 921600   # 1280 * 720

LOGGER = logging.getLogger("pdf2md")

FORMAT_ALIASES = {
    "md": "markdown",
    "markdown": "markdown",
    "json": "json",
}

# ROCm 性能环境变量，必须在 torch 导入前设置（进程启动时生效）：
# TORCH_ROCM_AOTRITON_ENABLE_EXPERIMENTAL：启用 AOTriton flash attention。
# 实测带 mask 的注意力从 ~7 ms 降到 ~1 ms（8 倍），ROCm 上不开这个开关
# 会退化到 math 内核，是解码阶段最大的单点瓶颈。
os.environ.setdefault("TORCH_ROCM_AOTRITON_ENABLE_EXPERIMENTAL", "1")
# MIOpen：FAST 模式不做穷举寻优（默认 EXTENSIVE 在 gfx1103 上首次跑
# 版面模型的卷积要编译/搜索二十分钟以上），一次性推理收益最大。
os.environ.setdefault("MIOPEN_FIND_MODE", "2")
os.environ.setdefault("MIOPEN_SEARCH_CUTOFF", "1")

def _setup_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def _bootstrap_sys_path() -> None:
    for path in (OXR_ROOT, LAYOUT_SERVICE_DIR):
        if str(path) not in sys.path:
            sys.path.insert(0, str(path))


def _parse_formats(raw: str) -> List[str]:
    formats: List[str] = []
    for item in raw.replace(";", ",").split(","):
        item = item.strip().lower()
        if not item:
            continue
        if item not in FORMAT_ALIASES:
            raise argparse.ArgumentTypeError(
                f"不支持的输出格式 {item!r}，可选：md、json"
            )
        normalized = FORMAT_ALIASES[item]
        if normalized not in formats:
            formats.append(normalized)
    if not formats:
        raise argparse.ArgumentTypeError("至少需要一种输出格式")
    return formats


def _parse_pages(raw: str) -> List[Tuple[int, Optional[int]]]:
    """把 "1-5,8,12-" 解析成 [(1,5), (8,8), (12,None)]，页码从 1 开始。"""
    ranges: List[Tuple[int, Optional[int]]] = []
    for item in raw.replace("，", ",").replace("~", "-").split(","):
        item = item.strip()
        if not item:
            continue
        if "-" in item:
            left, _, right = item.partition("-")
            left, right = left.strip(), right.strip()
            try:
                start = int(left) if left else 1
                end = int(right) if right else None
            except ValueError as error:
                raise argparse.ArgumentTypeError(f"无法解析页码范围 {item!r}") from error
            if start < 1:
                raise argparse.ArgumentTypeError(f"页码从 1 开始：{item!r}")
            if end is not None and end < start:
                raise argparse.ArgumentTypeError(f"页码范围倒置：{item!r}")
        else:
            try:
                start = int(item)
            except ValueError as error:
                raise argparse.ArgumentTypeError(f"无法解析页码 {item!r}") from error
            if start < 1:
                raise argparse.ArgumentTypeError(f"页码从 1 开始：{item!r}")
            end = start
        ranges.append((start, end))
    if not ranges:
        raise argparse.ArgumentTypeError("页码范围不能为空")
    return ranges


def select_page_indices(
    total_pages: int, ranges: Sequence[Tuple[int, Optional[int]]]
) -> List[int]:
    """按用户给定顺序返回 0 起的页下标，越界页自动丢弃，重复页只保留一次。"""
    indices: List[int] = []
    seen = set()
    for start, end in ranges:
        stop = total_pages if end is None else min(end, total_pages)
        for page in range(start, stop + 1):
            index = page - 1
            if index not in seen:
                seen.add(index)
                indices.append(index)
    return indices


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="使用 OXR-1.0 把 PDF 转成 Markdown，默认输出同名 .md 文件。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python pdf2md.py a.pdf\n"
            "  python pdf2md.py a.pdf --pages 1-5\n"
            "  python pdf2md.py -d D:\\pdfs\n"
            "  python pdf2md.py -d D:\\pdfs --recursive -o D:\\out\n"
            "  python pdf2md.py -d D:\\pdfs --formats md\n"
        ),
    )
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("pdf", nargs="?", help="单个 PDF 文件路径")
    source.add_argument("-d", "--dir", dest="directory", help="包含多个 PDF 的目录")

    parser.add_argument(
        "-o", "--out", help="输出目录；默认写到每个 PDF 所在目录，文件名保持不变"
    )
    parser.add_argument(
        "--recursive", action="store_true", help="配合 -d 递归处理子目录"
    )
    parser.add_argument(
        "--formats", type=_parse_formats, default=["markdown"],
        help="输出格式，逗号分隔，可选 md、json（默认：md，只输出 Markdown）",
    )
    parser.add_argument(
        "--pages", type=_parse_pages, default=None,
        help="只处理指定页（1 起），如 1-5、1-3,8、12-（默认：全部）",
    )
    parser.add_argument(
        "--chunk-pages", type=int, default=8,
        help="每批送进流水线的页数（默认：8）。越小越省显存、进度越细；"
             "调大会抬高显存峰值，本机 12 GB 显存在 12 页以上明显变慢",
    )
    parser.add_argument(
        "--batch-size", type=int, default=16,
        help="文本模型的批大小，同类型请求合批推理（默认：16，1 表示不合批；"
             "本地串行推理下 4 明显慢（约 158 s，对照 16 约 37 s）；"
             "8 与 16 持平；20/24 没有净收益）",
    )
    parser.add_argument(
        "--reading-order", choices=("geo", "llm"), default="llm",
        help=(
            "块阅读序的确定方式：llm=上游原行为，把每页整页图片再送模型推理一轮"
            "（更贵，但阅读序质量最高，默认）；geo=XY-Cut++ 几何排序（零 GPU 开销，"
            "可提速约 1/6，但对多栏、复杂版式等场景有局限，输出顺序可能不准）"
        ),
    )
    parser.add_argument(
        "--model", default=str(DEFAULT_MODEL_DIR),
        help=f"OXR 模型目录（默认：{DEFAULT_MODEL_DIR}）",
    )
    parser.add_argument(
        "--device", default="auto", help="推理设备：auto / cuda / cpu（默认：auto）"
    )
    parser.add_argument(
        "--dpi", type=int, default=150, help="PDF 渲染 DPI（默认：150，与官方服务一致）"
    )
    parser.add_argument(
        "--max-tokens", type=int, default=8192, help="单次识别最大生成 token 数（默认：8192）"
    )
    parser.add_argument(
        "--layout-batch-size", type=int, default=8, help="版面分析批大小（默认：8）"
    )
    parser.add_argument(
        "--layout-concurrency", type=int, default=2,
        help="版面分析在途批数（默认：2）",
    )
    parser.add_argument(
        "--torch-threads", type=int, default=0,
        help="torch 的 CPU 线程数，0 表示保持默认（默认：0）",
    )
    parser.add_argument(
        "--graph-decode", choices=("auto", "on", "off"), default="auto",
        help=(
            "是否把逐 token 的解码捕获成 CUDA Graph（默认：auto，CUDA 上自动开启）。"
            "auto/on 在 CPU、采样模式或 prompt 超出预留长度时自动回退"
        ),
    )
    parser.add_argument(
        "--graph-cache-len", type=int, default=2048,
        help=(
            "图解码预留的 KV 缓存长度（默认：2048）。"
            "prompt 与生成长度之和超过它时回退到普通生成，"
            "显存占用约 112 KB/行/千 token"
        ),
    )
    parser.add_argument(
        "--keep-header-footer", action="store_true", help="保留页眉和页脚"
    )
    parser.add_argument(
        "--skip-existing", action="store_true", help="输出文件已存在时跳过"
    )
    parser.add_argument(
        "--verbose", action="store_true", help="输出调试日志"
    )
    return parser


# --------------------------------------------------------------------------- #
# 运行统计
# --------------------------------------------------------------------------- #


class RunStats:
    """记录各阶段耗时与调用次数，跑完后一次性打印，方便定位瓶颈。

    版面分析的推理在独立线程里跑，统计字段可能被多个线程同时累加，
    所以统一加锁，避免累计值被写坏。
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self.layout_batches = 0
        self.layout_pages = 0
        self.layout_seconds = 0.0
        self.generate_batches = 0
        self.generate_samples = 0
        self.prepare_seconds = 0.0
        self.generate_seconds = 0.0
        self.output_tokens = 0
        self.length_hits = 0
        self.graph_batches = 0
        self.graph_fallbacks = 0

    def add_layout(self, pages: int, seconds: float) -> None:
        with self._lock:
            self.layout_batches += 1
            self.layout_pages += pages
            self.layout_seconds += seconds

    def add_prepare(self, seconds: float) -> None:
        with self._lock:
            self.prepare_seconds += seconds

    def add_generate(
        self, samples: int, tokens: int, seconds: float, graph: bool = False
    ) -> None:
        with self._lock:
            self.generate_batches += 1
            self.generate_samples += samples
            self.generate_seconds += seconds
            self.output_tokens += tokens
            if graph:
                self.graph_batches += 1

    def add_graph_fallback(self) -> None:
        with self._lock:
            self.graph_fallbacks += 1

    def add_length_hit(self) -> None:
        with self._lock:
            self.length_hits += 1

    def summary(self) -> str:
        if not self.generate_batches and not self.layout_batches:
            return ""
        lines = ["运行统计："]
        if self.layout_batches:
            lines.append(
                f"  版面分析：{self.layout_batches} 批 / {self.layout_pages} 页，"
                f"{self.layout_seconds:.1f}s"
                f"（{self.layout_seconds / max(1, self.layout_pages) * 1000:.0f} ms/页）"
            )
        if self.generate_batches:
            per_sample = self.generate_seconds / max(1, self.generate_samples)
            per_token = (
                self.generate_seconds / max(1, self.output_tokens) * 1000
                if self.output_tokens else 0.0
            )
            lines.append(
                f"  文本识别：{self.generate_batches} 批 / {self.generate_samples} 个请求，"
                f"预处理 {self.prepare_seconds:.1f}s，生成 {self.generate_seconds:.1f}s"
            )
            lines.append(
                f"    平均每请求 {per_sample * 1000:.0f} ms，"
                f"每输出 token {per_token:.1f} ms，共 {self.output_tokens} token"
            )
            if self.graph_batches:
                lines.append(
                    f"    CUDA Graph 解码 {self.graph_batches} 批"
                    + (
                        f"，超长回退 {self.graph_fallbacks} 批"
                        if self.graph_fallbacks
                        else ""
                    )
                )
            if self.length_hits:
                lines.append(
                    f"    达到 token 上限后重跑的块：{self.length_hits} 个"
                )
        return "\n".join(lines)


# --------------------------------------------------------------------------- #
# 本地 OXR 文本模型
# --------------------------------------------------------------------------- #


def _load_model_module(model_dir: Path):
    """把模型目录当作一个 Python 包导入，避免 trust_remote_code 的缓存模块问题。

    modeling_oxr.py 使用相对导入（from .configuration_oxr import ...），而
    transformers 的 trust_remote_code 会把每个文件复制到不同的缓存子目录，
    导致配置类和模型类来自不同模块副本、注册关系失效。这里直接把模型目录
    挂到一个临时包名下导入，保证所有类都来自同一份源码。
    """
    package_name = "_oxr_local_model"
    if package_name not in sys.modules:
        package = types.ModuleType(package_name)
        package.__path__ = [str(model_dir)]  # type: ignore[attr-defined]
        sys.modules[package_name] = package
    return importlib.import_module(f"{package_name}.modeling_oxr")


# CUDA Graph 解码的分块大小：每跑这么多步同步一次，看是否已经有结论。
GRAPH_CHUNK_STEPS = 16

# 单次图解码最多跑多少步；超出（或缓存装不下）时把没结束的行续跑一段。
# 正常情况下一次就跑完，这个值只是给超长输出兜底。
GRAPH_STAGE_STEPS = 4096

# 一批里的行同步推进，解码步数等于最长那条的输出长度。按“预估输出长度”
# 把长度接近的请求排进同一批，能显著减少陪跑；关掉它等于退回“先到先得”。
LENGTH_AWARE_BATCH = True

# 图解码的缓存长度按 256 的整数倍分桶（每桶捕获一张图复用），并在 prompt
# 之外预留这么多 token 给生成；预留越多，单步解码越慢（整段缓存每步都要读），
# 预留越少，超出容量的批越可能退回普通 generate。
GRAPH_CACHE_STEP = 256
GRAPH_CACHE_RESERVE = 512

# 图解码的 KV 缓存：每行每 token = 28 层 × 2（K/V）× 8 个 KV 头 × 128 维
# × 2 字节（bf16）。显存预算和 --graph-cache-len 的说明都用它换算。
GRAPH_KV_BYTES_PER_TOKEN = 28 * 2 * 8 * 128 * 2

# 攒批的最长等待（秒）。攒满的批能把视觉编码器与预填充的固定开销摊薄，所以
# 愿意多等一小会儿；但本地生成本来就是瓶颈，密集到一定程度后继续等只会多
# 添延迟。0.5s 只是上限——真正的停止判据是「队列连着 30 ms 没再增长」
# 或「够两批」（见 _next_group），到点就走。
GROUP_WAIT_SECONDS = 0.5

# 视觉塔与解码重叠的总开关：视觉塔是算力受限、解码是带宽受限，两条流上
# 同时跑能互相填空（16 页实测视觉塔约占识别时间 31%）。视觉塔一失败就会
# 自动关掉重叠，退回串行路径。
VISION_OVERLAP = True

# 单次视觉前向最多吃这么多 patch。重叠时视觉塔跑在独立流上，两条流各有各
# 的显存池（显存块只能被同一条流复用），整批整页图一次前向要吃掉 2~5 GB，
# 和 KV 缓存挤在一起会把 reserved 顶到 12 GB 卡的边界、触发驱动换页 ——
# 实测不是慢一点，而是慢 2~3 倍（32 页 55s 变 124s）。按 patch 数切块后峰值
# 只和这一块有关，视觉总计算量不变。
#
# 上限也不是越大越好：16384 与 14000 在 16 页内看不出问题（32.3s、62.6s），
# 但连续跑几个 chunk 之后就开始抖动（32 页 135s、124s）。8192 与 6144 一样
# 稳定（32 页 54.7s、55.0s），全本 67 页识别 123.5s vs 123.6s，而切得少
# 意味着 bf16 累加顺序更接近整批前向 —— 输出与串行路径的差异从 3 处降到
# 1 处（“2.8 扶手”多一个空格）。所以取 8192。
VISION_CHUNK_PATCHES = 8192

# 图捕获必须在「没有别的流在提交 kernel」时进行，否则 capture 会报
# “operation not permitted when stream is capturing”。视觉塔的提交与捕获
# 共用这把锁：两边都只锁提交、不锁等待，所以重叠照样成立。
_CUDA_CAPTURE_LOCK = threading.Lock()


def _cuda_graph_supported(torch, device) -> bool:
    """CUDA Graph 只在 NVIDIA 的 CUDA 上启用。

    ROCm 版的 torch 也把设备类型报成 "cuda"，`torch.cuda.is_available()` 同样
    是 True，但图捕获依赖的若干行为在 HIP 上并不保证（StaticCache 的原地写入、
    内核句柄复用等）。两者的区分点是版本号：CUDA 版有 `torch.version.cuda`、
    没有 `hip`，ROCm 版反过来。非 NVIDIA 环境直接保留原来的 `generate`，
    只是慢一些，转换结果不受影响。
    """
    if device.type != "cuda":
        return False
    if getattr(torch.version, "hip", None) is not None:
        return False
    return getattr(torch.version, "cuda", None) is not None


@contextlib.contextmanager
def _sdpa_gqa_enabled():
    """临时让 SDPA 在带 mask 时也走 `enable_gqa`。

    transformers 的 sdpa 实现在 `attention_mask` 非空时会用 `repeat_kv` 把
    K/V 整段复制成和 Q 一样多的 head，每层两次 contiguous 拷贝，耗时随缓存
    长度线性增长（实测占单步 GPU 时间的 42%）。torch 的融合内核本身支持
    GQA + mask，只是 transformers 出于保守把它关掉了；这里临时打开，退出时
    复原，避免影响走 `generate` 的代码路径。
    """
    from transformers.integrations import sdpa_attention

    original = sdpa_attention.use_gqa_in_sdpa
    sdpa_attention.use_gqa_in_sdpa = lambda *args, **kwargs: True
    try:
        yield
    finally:
        sdpa_attention.use_gqa_in_sdpa = original


def _probe_sdpa_gqa(torch, config, device, dtype) -> bool:
    """先用模型真实的 head 数试一次“带 mask 的 GQA”，失败就不开这个优化。

    个别 torch / 显卡组合在这种组合下会直接触发 device-side assert，整张卡
    作废且无法恢复，所以先在几个小张量上试一遍，再决定要不要真的用。
    """
    try:
        heads = int(getattr(config, "num_attention_heads", 0))
        kv_heads = int(getattr(config, "num_key_value_heads", heads))
        head_dim = int(getattr(config, "head_dim", 128))
        if heads <= 0 or kv_heads >= heads:
            return False
        q = torch.randn((1, heads, 1, head_dim), device=device, dtype=dtype)
        k = torch.randn((1, kv_heads, 512, head_dim), device=device, dtype=dtype)
        v = torch.randn((1, kv_heads, 512, head_dim), device=device, dtype=dtype)
        mask = torch.ones((1, 1, 1, 512), dtype=torch.bool, device=device)
        with torch.inference_mode():
            out = torch.nn.functional.scaled_dot_product_attention(
                q, k, v, attn_mask=mask, enable_gqa=True
            )
        return bool(torch.isfinite(out).all().item())
    except Exception:  # noqa: BLE001 - 不支持就退回 repeat_kv 的写法
        return False


class GreedyGraphDecoderPool:
    """按 (批大小, 缓存长度) 复用图解码器，避免为每种组合重复捕获。

    显存和速度都由「批 × 缓存长度」决定：每行每 token 的 KV 是 28 层 ×
    2 × 8 头 × 128 维 × 2 字节 = 112 KB/千 token（见 --graph-cache-len 说明）。
    一步解码要读完整段预留缓存，所以缓存长度直接决定单步耗时（实测批 16：
    缓存 2048 → 38.8 ms/步，512 → 18.6 ms/步）。预填充也要按补齐后的行数算，
    所以 1 行请求塞进 16 行的图里，光预填充就要多花十几倍——小批必须用小图。

    因此这里按「实际行数向上取到 2 的幂」和「prompt 长度 + 预留量」选一张
    现成的图；池子只保留最近用到的几张，换出时释放其 KV 缓存。

    池子的字节预算不能按“显存还空着多少”给满：本机 12 GB 卡上留到 4 GiB
    就开始抖动（32 页实测：3 GiB→64 s，4 GiB→106 s，5 GiB→215 s），
    所以默认只给 3 GiB，宁可每个 chunk 重新捕获一次（0.2~0.5 s）。
    """

    # 批大小阶梯：补齐到不小于实际行数的 2 的幂
    _LADDER = (1, 2, 4, 8, 16, 32, 48, 64)

    def __init__(
        self,
        llm,
        max_batch: int,
        max_cache_len: int,
        device,
        max_bytes: int = 3 * 1024 ** 3,
        max_entries: int = 4,
    ) -> None:
        self.llm = llm
        self.max_batch = max(1, int(max_batch))
        self.max_cache_len = max(256, int(max_cache_len))
        self.device = device
        self.max_bytes = max(0, int(max_bytes))
        self.max_entries = max(1, int(max_entries))
        self._decoders: Dict[Tuple[int, int], GreedyGraphDecoder] = {}
        self._order: List[Tuple[int, int]] = []
        self._lock = threading.Lock()

    def _batch_bucket(self, rows: int) -> int:
        for size in self._LADDER:
            if rows <= size:
                return min(size, self.max_batch)
        return self.max_batch

    def cache_bucket(self, prompt_length: int, reserve: Optional[int] = None) -> int:
        """给摸清长度的请求选缓存长度：prompt + 预留，向上取整到 256。"""
        if reserve is None:
            reserve = GRAPH_CACHE_RESERVE
        wanted = prompt_length + max(1, int(reserve))
        bucket = -(-wanted // GRAPH_CACHE_STEP) * GRAPH_CACHE_STEP
        return max(GRAPH_CACHE_STEP, min(self.max_cache_len, bucket))

    def acquire(
        self, rows: int, prompt_length: int, cache_len: Optional[int] = None
    ) -> Optional[GreedyGraphDecoder]:
        """取一张能覆盖 (rows, prompt_length) 的图解码器，必要时现场捕获。"""
        if rows > self.max_batch:
            return None
        if cache_len is None:
            cache_len = self.cache_bucket(prompt_length)
        key = (self._batch_bucket(rows), int(cache_len))
        with self._lock:
            decoder = self._decoders.pop(key, None)
            if decoder is not None:
                self._decoders[key] = decoder
                self._order = [item for item in self._order if item != key] + [key]
                return decoder
        batch, cache_len = key
        if cache_len <= prompt_length + 2:
            return None
        try:
            decoder = GreedyGraphDecoder(
                self.llm, batch=batch, cache_len=cache_len, device=self.device
            )
        except Exception as error:  # noqa: BLE001 - 捕获失败就退回 generate
            LOGGER.warning(
                "CUDA Graph 解码器创建失败（批 %d，缓存 %d）：%s", batch, cache_len, error
            )
            return None
        with self._lock:
            self._decoders[key] = decoder
            self._order.append(key)
            # KV 缓存是「批 × 缓存长度 × 112 KB」量级，留太多张会把显存吃光，
            # 但留太少又会因为桶来回切换而反复捕获（每次 0.2~0.5 s）。按总字节
            # 数淘汰，块大小不同的几张图才能长期共存。
            while self._order and (
                len(self._order) > self.max_entries
                or self._decoder_bytes() > self.max_bytes and len(self._order) > 1
            ):
                stale = self._order.pop(0)
                self._decoders.pop(stale, None)
        LOGGER.info("图解码器就绪：批 %d，缓存 %d token", batch, cache_len)
        return decoder

    def _decoder_bytes(self) -> int:
        total = 0
        for key in self._order:
            batch, cache_len = key
            total += batch * cache_len * GRAPH_KV_BYTES_PER_TOKEN
        return total


class GreedyGraphDecoder:
    """把逐 token 的贪心解码捕获成一张 CUDA Graph。

    本地推理的瓶颈不在算力。实测单步 decode 的 GPU 时间只有约 11 ms，其余
    50 ms 全花在 Windows 上几百次 kernel 提交的 CPU 开销（也能解释“CPU 高、
    GPU 空转”）。解码段的形状本来是固定的（每步 1 个 token），只要把 KV
    缓存按固定长度预分配（StaticCache），整段就能捕获成一张图，之后每步只剩
    一次 graph.replay()。

    为了让“缓存位置”和 mask 都能在图内推进，这两者用固定缓冲区保存：图里
    每步把新 token 写回输入、把缓存位置 +1、把 mask 的下一位打开，因此一个
    图可以被任意多个请求复用，只要 prompt 与生成长度不超出预留长度。
    """

    def __init__(
        self,
        llm,
        batch: int,
        cache_len: int,
        device,
    ) -> None:
        import torch
        from transformers import StaticCache

        self.torch = torch
        self.llm = llm
        self.batch = int(batch)
        self.cache_len = int(cache_len)
        self.device = device

        self.cache = StaticCache(config=llm.config, max_cache_len=self.cache_len)
        # 带 mask 的 GQA 能省掉每层两次 K/V 拷贝，单步快约 40%；
        # 能不能用取决于 torch 与显卡，先探测再决定。
        dtype = next(llm.parameters()).dtype
        self.use_gqa = _probe_sdpa_gqa(torch, llm.config, device, dtype)
        self.input_ids = torch.zeros((self.batch, 1), dtype=torch.long, device=device)
        self.mask = torch.zeros(
            (self.batch, 1, 1, self.cache_len), dtype=torch.bool, device=device
        )
        # 能生成多少 token 只受缓存长度约束，history 直接按缓存长度开
        self.history = torch.zeros(
            (self.batch, self.cache_len), dtype=torch.long, device=device
        )
        self.step = torch.zeros((1,), dtype=torch.long, device=device)
        self.mask_pos = torch.zeros((1,), dtype=torch.long, device=device)
        # prompt 长度每个请求都不同，但图里不能出现 Python 常量，统一走张量
        self.pos_base = torch.zeros((1,), dtype=torch.long, device=device)

        self.graph = None
        self.first = None
        self.mask2d = None
        self.prompt_length = 0

    # -- 状态复位（每次请求开始前） ---------------------------------------- #

    def _set_cumulative(self, value: int) -> None:
        """把 KV 缓存的写入位置拨到 value（StaticCache 内部是张量，图里也用它）。"""
        with self.torch.inference_mode():
            for layer in self.cache.layers:
                if getattr(layer, "is_initialized", False):
                    layer.cumulative_length.zero_().add_(value)

    def _reset(self) -> None:
        """回到“预填充刚结束”的状态，供下一次 replay 从第一个新 token 开始。"""
        with self.torch.inference_mode():
            self._set_cumulative(self.prompt_length)
            self.step.zero_()
            self.pos_base.fill_(self.prompt_length)
            self.mask_pos.fill_(self.prompt_length)
            self.mask.zero_()
            self.mask[..., : self.prompt_length] = self.mask2d[:, None, None, :]
            # 第一个新 token 自己的位置也要放行（它的 key 在这一步刚写进缓存）
            self.mask[..., self.prompt_length] = True
            self.input_ids.copy_(self.first)
            self.history.zero_()
            self.history[:, 0:1].copy_(self.first)

    # -- 一步解码（会被捕获进图） ------------------------------------------ #

    def _step(self) -> None:
        out = self.llm(
            input_ids=self.input_ids,
            attention_mask=self.mask,
            past_key_values=self.cache,
            use_cache=True,
            logits_to_keep=1,
        )
        nxt = out.logits[:, -1].argmax(-1, keepdim=True)
        # history[0] 是预填充给出的第一个 token，本例输出写在下一位
        self.history.scatter_(
            1, (self.step + 1).view(1, 1).expand(self.batch, 1), nxt
        )
        self.input_ids.copy_(nxt)
        self.step.add_(1)
        self.mask_pos.copy_(self.pos_base + self.step)
        self.mask.index_fill_(3, self.mask_pos, True)

    # -- 对外接口 ---------------------------------------------------------- #

    def supports(self, prompt_length: int) -> bool:
        # 至少留出 1 个 token 的生成空间，并且 mask 的“下一位”不能越界
        return prompt_length + 2 <= self.cache_len

    def capacity(self, prompt_length: int) -> int:
        """这批请求最多能生成多少个 token（受预分配的缓存长度限制）。"""
        return max(0, self.cache_len - prompt_length - 1)

    def _fit_batch(self, tensor, expand_dims):
        """把不足一批的行用最后一行的副本补齐，避免整行被 mask 掉产生 NaN。"""
        rows = int(tensor.shape[0])
        if rows == self.batch:
            return tensor
        tail = tensor[-1:].expand(self.batch - rows, *expand_dims)
        return self.torch.cat([tensor, tail], dim=0)

    def decode(
        self,
        inputs_embeds,
        attention_mask,
        max_new_tokens: int,
        eos_lookup: set,
    ):
        """返回 (每行的 token 列表, 是否全部正常结束)。

        第二个返回值为 False 表示这批 token 超出了预留的生成长度，调用方应当
        回退到 transformers.generate 重跑，避免把结果截断。
        """
        torch = self.torch
        rows = int(inputs_embeds.shape[0])
        if rows > self.batch:
            raise ValueError(f"图解码的批大小是 {self.batch}，收到 {rows} 行")
        budget = min(int(max_new_tokens), self.capacity(int(attention_mask.shape[1])))
        embeds = self._fit_batch(inputs_embeds, (inputs_embeds.shape[1], inputs_embeds.shape[2]))
        mask2d = self._fit_batch(attention_mask, (attention_mask.shape[1],))

        gqa = _sdpa_gqa_enabled() if self.use_gqa else contextlib.nullcontext()
        with gqa, torch.inference_mode():
            self._set_cumulative(0)
            out = self.llm(
                inputs_embeds=embeds,
                attention_mask=mask2d,
                past_key_values=self.cache,
                use_cache=True,
                logits_to_keep=1,
            )
            self.first = out.logits[:, -1].argmax(-1, keepdim=True).clone()
            self.mask2d = mask2d
            self.prompt_length = int(mask2d.shape[1])
            self._reset()
            if self.graph is None:
                self._capture()

        produced = 1
        finished = [token in eos_lookup for token in self.first[:rows, 0].tolist()]
        stream = torch.cuda.current_stream(self.device)
        while produced < budget and not all(finished):
            steps = min(GRAPH_CHUNK_STEPS, budget - produced)
            with torch.inference_mode():
                for _ in range(steps):
                    self.graph.replay()
                # 只等自己这条流：视觉塔可能同时在另一条流上跑，等整卡会把
                # 那点重叠抹掉（实测重叠收益的大头就丢在这里）。
                stream.synchronize()
            produced += steps
            fresh = self.history[:rows, produced - steps : produced].tolist()
            for row in range(rows):
                if not finished[row]:
                    finished[row] = any(token in eos_lookup for token in fresh[row])

        tokens = self.history[:rows, :produced].tolist()
        complete = all(finished) or produced >= int(max_new_tokens)
        return tokens, complete

    def _capture(self) -> None:
        torch = self.torch
        started = time.perf_counter()
        # 捕获期间不能让另一条流（视觉塔）提交 kernel，否则 capture 会报
        # “operation not permitted when stream is capturing”。两边只锁“提交”
        # 这一段，不锁等待，所以不会把重叠吃掉。
        with _CUDA_CAPTURE_LOCK:
            stream = torch.cuda.Stream(device=self.device)
            stream.wait_stream(torch.cuda.current_stream(self.device))
            with torch.inference_mode(), torch.cuda.stream(stream):
                for _ in range(2):  # 预热：把 cuBLAS/算子句柄都建好再捕获
                    self._step()
            torch.cuda.current_stream(self.device).wait_stream(stream)
            torch.cuda.synchronize(self.device)
            self._reset()
            self.graph = torch.cuda.CUDAGraph()
            with torch.inference_mode(), torch.cuda.graph(self.graph):
                self._step()
            torch.cuda.synchronize(self.device)
            self._reset()
        LOGGER.info(
            "CUDA Graph 解码已捕获：批 %d，缓存 %d token，耗时 %.1fs",
            self.batch,
            self.cache_len,
            time.perf_counter() - started,
        )


class OXRTextBackend:
    """用 transformers 本地加载 OXR-1.0，按批执行视觉问答。

    预处理（图像缩放、模板渲染、张量搬运）和生成（GPU）被拆成两步：
    预处理在独立的单线程池里做，生成在另一个线程里做，两者可以重叠，
    这样 CPU 侧的图像处理不会白等 GPU 解码。
    """

    def __init__(
        self,
        model_dir: Path,
        device,
        stats: Optional[RunStats] = None,
        graph_decode: bool = False,
        graph_batch: int = 4,
        graph_cache_len: int = 2048,
    ) -> None:
        import torch
        from transformers import GenerationConfig

        self.torch = torch
        self.device = device
        self.dtype = torch.bfloat16 if device.type == "cuda" else torch.float32
        self.stats = stats

        module = _load_model_module(model_dir)
        LOGGER.info("正在加载 OXR 模型：%s（%s，%s）", model_dir, device, self.dtype)
        started = time.perf_counter()
        try:
            model = module.OXR.from_pretrained(str(model_dir), dtype=self.dtype)
        except TypeError:  # transformers 4.x 使用 torch_dtype
            model = module.OXR.from_pretrained(
                str(model_dir), torch_dtype=self.dtype
            )
        self.model = model.to(device).eval()

        # from_pretrained(low_cpu_mem_usage=True) 会在 meta/空张量上构造模型，
        # 非持久化 buffer 不会被 checkpoint 覆盖，可能是未初始化的内存。
        # VisionRotaryEmbedding.inv_freq 在 forward 中会重新计算，这里只需修复
        # 视觉词表的指示 token 下标，否则会触发 CUDA device-side assert。
        vocab_size = int(self.model.config.visual_vocab_size)
        indicator_ids = len(module.INDICATOR_IDS)
        self.model.indicator_token_indices = torch.arange(
            vocab_size - indicator_ids, vocab_size, dtype=torch.long, device=device
        )
        self.tokenizer = self.model.text_tokenizer

        # OXR.generate 内部委托给 llm.generate，而 llm 的 generation_config 只带
        # eos=151643；模型实际用 <|im_end|>(151645) 结束，必须显式传入。
        try:
            self.generation_config = GenerationConfig.from_pretrained(str(model_dir))
        except (OSError, ValueError):
            self.generation_config = None

        pad_token_id = self.tokenizer.pad_token_id
        if pad_token_id is None:
            pad_token_id = self.tokenizer.eos_token_id
        if pad_token_id is None:
            raise SystemExit("分词器缺少 pad/eos token，无法批量生成。")
        self.pad_token_id = int(pad_token_id)

        eos_token_id = (
            self.generation_config.eos_token_id
            if self.generation_config is not None
            else None
        )
        if eos_token_id is None:
            eos_token_id = self.tokenizer.eos_token_id
        if isinstance(eos_token_id, (list, tuple)):
            self.eos_token_ids = [int(value) for value in eos_token_id]
        elif eos_token_id is None:
            self.eos_token_ids = []
        else:
            self.eos_token_ids = [int(eos_token_id)]
        self._eos_lookup = set(self.eos_token_ids)
        LOGGER.info("模型加载完成，耗时 %.1fs", time.perf_counter() - started)

        # 图解码只在 NVIDIA 的 CUDA 上、且是贪心生成时可用；创建实例不占多少
        # 显存，StaticCache 由第一次使用时的一次真实预填充惰性分配。
        self.graph_pool: Optional[GreedyGraphDecoderPool] = None
        if graph_decode and _cuda_graph_supported(torch, device):
            try:
                self.graph_pool = GreedyGraphDecoderPool(
                    self.model.llm,
                    max_batch=max(1, int(graph_batch)),
                    max_cache_len=max(256, int(graph_cache_len)),
                    device=device,
                )
                LOGGER.info(
                    "CUDA Graph 解码可用（最大批 %d，缓存上限 %d token）",
                    self.graph_pool.max_batch,
                    self.graph_pool.max_cache_len,
                )
            except Exception as error:  # noqa: BLE001 - 退化为普通 generate
                LOGGER.warning("CUDA Graph 解码不可用，回退到 generate：%s", error)
                self.graph_pool = None
        elif graph_decode:
            LOGGER.info(
                "当前设备不是 NVIDIA CUDA（%s，torch %s），保留 transformers.generate；"
                "CUDA Graph 仅在 NVIDIA 显卡上启用",
                device.type,
                getattr(torch.version, "cuda", None) or "ROCm/CPU",
            )

    # -- 预处理（纯 CPU） -------------------------------------------------- #

    def prepare(self, prompt: str, image_bgr):
        """把一张裁剪图和提示词变成模型输入，全程不碰 GPU。"""
        import cv2
        from PIL import Image

        rgb = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2RGB)
        image = Image.fromarray(rgb)
        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "image", "image": image},
                    {"type": "text", "text": prompt},
                ],
            }
        ]
        input_ids, pixel_values, grid_thws = self.model.preprocess_inputs(
            messages, min_pixels=MIN_PIXELS, max_pixels=MAX_PIXELS
        )
        return input_ids[0], pixel_values, grid_thws

    def prepare_batch(self, items: Sequence[Tuple[str, Any]]):
        """逐张预处理；单张失败只影响它自己，不会毁掉整批。"""
        started = time.perf_counter()
        prepared = []
        failures: List[Tuple[int, Exception]] = []
        for position, (prompt, image) in enumerate(items):
            try:
                prepared.append(self.prepare(prompt, image))
            except Exception as error:  # noqa: BLE001 - 交给调用方重试
                failures.append((position, error))
        if self.stats is not None:
            self.stats.add_prepare(time.perf_counter() - started)
        return prepared, failures

    def can_overlap_vision(self) -> bool:
        """视觉塔能否挪到独立流上、与上一批的解码重叠（只有 CUDA 能）。"""
        return bool(VISION_OVERLAP) and self.device.type == "cuda"

    # -- 生成（GPU） ------------------------------------------------------- #

    def _generation_kwargs(
        self, max_new_tokens: int, model_options: Optional[Dict[str, Any]]
    ) -> Dict[str, Any]:
        from oxr.model.client import resolve_model_options

        options = resolve_model_options(model_options)
        limit = int(max_new_tokens or 8192)
        generation_kwargs: Dict[str, Any] = {
            "max_new_tokens": limit,
            "use_cache": True,
            "pad_token_id": self.pad_token_id,
        }
        if self.eos_token_ids:
            generation_kwargs["eos_token_id"] = self.eos_token_ids
        if options["repetition_penalty"] != 1.0:
            generation_kwargs["repetition_penalty"] = options["repetition_penalty"]
        if options["temperature"] > 0:
            generation_kwargs.update(
                do_sample=True,
                temperature=options["temperature"],
                top_p=options["top_p"],
            )
        else:
            generation_kwargs["do_sample"] = False
        return generation_kwargs

    def _graph_decode_rows(
        self,
        inputs_embeds,
        attention_mask,
        row_limits: Sequence[int],
    ) -> Optional[List[List[int]]]:
        """用图解码跑完整批；缓存装不下时长出来的行续跑，而不是整批重来。

        一批里的行是同步推进的：只要还有一条长请求没结束，整批都得陪着跑，
        所以理想情况是让已经结束的行早点退出。但实测「按步数分段 + 把结束行
        剔出去续跑」在本机并不划算——每次续跑都要按新 prompt 换一张图（重新
        捕获 0.2~0.5 s）并重做一遍预填充，省的步数全被吃掉（16 页实测慢 5%）。
        真正划得来的是只在**缓存装不下**时才续跑：那本来要退回普通 generate
        把整批重跑一遍，代价大得多。

        返回每行的 token 列表；图解码用不上（显存/长度不够）时返回 None，
        由调用方回退到普通 generate。
        """
        torch = self.torch
        pool = self.graph_pool
        if pool is None:
            return None

        eos_lookup = self._eos_lookup
        rows = int(inputs_embeds.shape[0])
        tokens: List[List[int]] = [[] for _ in range(rows)]
        active = list(range(rows))
        embeds = inputs_embeds
        mask = attention_mask
        embedding = self.model.llm.get_input_embeddings()
        # 缓存长度按「prompt + 预留」一次定好，正常情况下整批只用一张图。
        reserve = max(1, min(GRAPH_CACHE_RESERVE, max(row_limits) - 1))
        cache_len = pool.cache_bucket(int(mask.shape[1]), reserve)

        while active:
            prompt_length = int(mask.shape[1])
            decoder = pool.acquire(len(active), prompt_length, cache_len)
            if decoder is None or not decoder.supports(prompt_length):
                # 这一行长到了预留量之外：按当前长度重新规划缓存长度再取图。
                cache_len = pool.cache_bucket(prompt_length, reserve)
                decoder = pool.acquire(len(active), prompt_length, cache_len)
            if decoder is None or not decoder.supports(prompt_length):
                LOGGER.debug(
                    "图解码放弃：行数 %d，prompt %d，缓存 %d",
                    len(active), prompt_length, cache_len,
                )
                return None
            budget = max(row_limits[row] - len(tokens[row]) for row in active)
            budget = min(budget, GRAPH_STAGE_STEPS)
            if budget <= 0:
                break
            stage, _complete = decoder.decode(embeds, mask, budget, eos_lookup)

            pending: List[Tuple[int, int]] = []
            for slot, row in enumerate(active):
                produced = stage[slot]
                tokens[row].extend(produced)
                if len(tokens[row]) >= row_limits[row]:
                    continue
                if any(token in eos_lookup for token in produced):
                    continue
                pending.append((slot, row))
            if not pending:
                return tokens

            # 只把没结束的行拼回下一轮：原 prompt 的合法区间（mask 为真的部分）
            # 加上已经生成的前缀，右侧对齐后统一左填充。
            pieces = []
            for slot, row in pending:
                piece_embeds = embeds[slot:slot + 1]
                piece_mask = mask[slot:slot + 1]
                prefix = tokens[row]
                if prefix:
                    index = torch.tensor(
                        [prefix], dtype=torch.long, device=embeds.device
                    )
                    piece_embeds = torch.cat(
                        [piece_embeds, embedding(index)], dim=1
                    )
                    piece_mask = torch.cat(
                        [
                            piece_mask,
                            torch.ones(
                                (1, len(prefix)), dtype=torch.bool,
                                device=embeds.device,
                            ),
                        ],
                        dim=1,
                    )
                pieces.append((piece_embeds, piece_mask))

            width = max(int(piece.shape[1]) for piece, _ in pieces)
            next_embeds = torch.zeros(
                (len(pieces), width, embeds.shape[2]),
                dtype=embeds.dtype,
                device=embeds.device,
            )
            next_mask = torch.zeros(
                (len(pieces), width), dtype=torch.bool, device=embeds.device
            )
            for index, (piece_embeds, piece_mask) in enumerate(pieces):
                length = int(piece_embeds.shape[1])
                next_embeds[index:index + 1, width - length:] = piece_embeds
                next_mask[index, width - length:] = piece_mask[0]
            embeds, mask = next_embeds, next_mask
            active = [row for _slot, row in pending]

        return tokens

    # -- 视觉前向与解码拆开，方便两者在两条流上重叠 ------------------------ #

    def build_batch(self, prepared: Sequence[Any]) -> Dict[str, Any]:
        """把预处理结果拼成一个 batch（纯 CPU 侧拼装，暂不搬显存）。

        左填充到等长，靠 attention_mask 屏蔽填充位；视觉 token 用负数的
        指示 ID 占位，merge_multimodal 会按行序填回视觉特征，所以填充用的
        pad_token_id（正数）不会和它们冲突。搬显存的动作用 merge_batch 做，
        好让它和视觉前向落在同一条流上。
        """
        torch = self.torch
        lengths = [int(input_ids.shape[0]) for input_ids, _, _ in prepared]
        max_length = max(lengths)
        rows = []
        for input_ids, _, _ in prepared:
            padding = max_length - int(input_ids.shape[0])
            if padding > 0:
                input_ids = torch.cat([
                    torch.full((padding,), self.pad_token_id, dtype=torch.long),
                    input_ids,
                ])
            rows.append(input_ids)
        input_ids = torch.stack(rows)

        attention_mask = torch.zeros_like(input_ids, dtype=torch.bool)
        for row, length in enumerate(lengths):
            attention_mask[row, max_length - length:] = True

        pixel_values = [item[1] for item in prepared]
        grid_thws = [item[2] for item in prepared]
        return {
            "input_ids": input_ids,
            "attention_mask": attention_mask,
            "pixel_values": pixel_values,
            "grid_thws": grid_thws,
            "rows": len(prepared),
        }

    def _vision_groups(self, grid_thws: Sequence[Any], rows: int) -> List[List[int]]:
        """按 patch 数把行切成若干组，组内一次视觉前向。

        行与行在视觉塔里完全独立，切开再拼回来内容不变；之所以要切，是因为
        视觉塔跑在独立流上、显存池和 KV 缓存那条流不共享，一整批整页图
        （8 页阅读序就是 27648 patch）会顶到 12 GB 卡的容量上限。
        """
        groups: List[List[int]] = []
        current: List[int] = []
        patches = 0
        for row in range(rows):
            grid = grid_thws[row]
            if grid is None:  # 没有图的行单独成组，走不带像素的那条路
                if current:
                    groups.append(current)
                    current, patches = [], 0
                groups.append([row])
                continue
            count = int(grid.prod().item())
            if current and patches + count > VISION_CHUNK_PATCHES:
                groups.append(current)
                current, patches = [], 0
            current.append(row)
            patches += count
        if current:
            groups.append(current)
        return groups

    def _merge_rows(
        self,
        input_ids,
        rows: Sequence[int],
        pixel_values: Sequence[Any],
        grid_thws: Sequence[Any],
    ):
        """只给指定行做一次视觉前向 + 嵌入合并（行序、宽度都不变）。"""
        torch = self.torch
        pixels = [pixel_values[row] for row in rows if pixel_values[row] is not None]
        grids = [grid_thws[row] for row in rows if grid_thws[row] is not None]
        pv = None
        if pixels:
            pv = torch.cat(pixels, dim=0).to(self.device, dtype=self.dtype)
        gt = None
        if grids:
            gt = torch.cat(grids, dim=0).to(self.device)
        with torch.inference_mode():
            return self.model.merge_multimodal(
                input_ids=input_ids,
                pixel_values=pv,
                grid_thws=gt,
            )

    def merge_batch(self, batch: Dict[str, Any]):
        """视觉塔前向 + 文本嵌入合并，在调用方当前的 CUDA 流上执行。

        视觉塔是算力受限、解码是带宽受限，两者在 GPU 上能互相填空，所以
        这一步允许跑在独立流上，与上一批的解码重叠。patch 多的时候按组
        分批前向，把这次前向的显存峰值压住（见 VISION_CHUNK_PATCHES）。
        """
        torch = self.torch
        input_ids = batch["input_ids"].to(self.device)
        groups = self._vision_groups(batch["grid_thws"], int(batch["rows"]))
        if len(groups) <= 1:
            return self._merge_rows(
                input_ids, groups[0] if groups else [],
                batch["pixel_values"], batch["grid_thws"],
            )
        parts = [
            self._merge_rows(
                input_ids[group], group, batch["pixel_values"], batch["grid_thws"]
            )
            for group in groups
        ]
        # 各组都是原批里连续的一段，直接按顺序拼回整批。
        return torch.cat(parts, dim=0)

    def run_batch(
        self,
        batch: Dict[str, Any],
        inputs_embeds,
        max_new_tokens: int,
        model_options: Optional[Dict[str, Any]] = None,
        limits: Optional[Sequence[int]] = None,
        started: Optional[float] = None,
        vision_event=None,
    ):
        """拿已合并好的嵌入跑预填充 + 解码。

        vision_event 是视觉塔在另一条流上跑完后记下的事件：解码前等它一次
        即可，不能再去等整卡（那会把重叠整个抹掉）。
        """
        from oxr.model.client import ModelOutput

        torch = self.torch
        if started is None:
            started = time.perf_counter()
        if vision_event is not None:
            stream = torch.cuda.current_stream(self.device)
            stream.wait_event(vision_event)
            # 嵌入是在视觉那条流上分配的，这里要告诉分配器“它也会在本流
            # 用过”，否则视觉流上的下一批可能提前复用这块显存。
            if isinstance(inputs_embeds, torch.Tensor):
                inputs_embeds.record_stream(stream)

        attention_mask = batch["attention_mask"].to(self.device)
        generation_kwargs = self._generation_kwargs(max_new_tokens, model_options)
        limit = int(generation_kwargs["max_new_tokens"])
        row_limits = list(limits or [limit] * int(batch["rows"]))
        greedy = (
            generation_kwargs.get("do_sample") is False
            and "repetition_penalty" not in generation_kwargs
        )
        graph_used = False
        tokens_by_row = None

        if self.graph_pool is not None and greedy:
            try:
                with torch.inference_mode():
                    decoded = self._graph_decode_rows(
                        inputs_embeds, attention_mask, row_limits
                    )
            except Exception as error:  # noqa: BLE001 - 图解码失败就退回 generate
                LOGGER.warning("CUDA Graph 解码失败，回退到 generate：%s", error)
                decoded = None
            if decoded is not None:
                tokens_by_row = decoded
                graph_used = True
            elif self.stats is not None:
                self.stats.add_graph_fallback()

        if tokens_by_row is None:
            with torch.inference_mode():
                generated = self.model.llm.generate(
                    inputs=None,
                    inputs_embeds=inputs_embeds,
                    attention_mask=attention_mask,
                    **generation_kwargs,
                )
            tokens_by_row = [row.tolist() for row in generated]

        results = []
        produced_tokens = 0
        for row, row_limit in enumerate(row_limits):
            tokens = tokens_by_row[row]
            # 合批后整批按最大的行上限生成，本行要按自己的上限截断，
            # 否则一条短请求会白等最长的那条。
            cut = min(len(tokens), row_limit)
            for position, token in enumerate(tokens[:cut]):
                if token in self._eos_lookup:
                    cut = position
                    break
            text = self.tokenizer.decode(tokens[:cut], skip_special_tokens=True)
            if cut >= row_limit:
                finish_reason = "length"
            else:
                finish_reason = "stop"
            if finish_reason == "length" and self.stats is not None:
                self.stats.add_length_hit()
            produced_tokens += cut
            results.append(ModelOutput(text, finish_reason=finish_reason))

        if self.stats is not None:
            self.stats.add_generate(
                int(batch["rows"]),
                produced_tokens,
                time.perf_counter() - started,
                graph_used,
            )
        return results

    def generate_batch(
        self,
        prepared: Sequence[Any],
        max_new_tokens: int,
        model_options: Optional[Dict[str, Any]] = None,
        limits: Optional[Sequence[int]] = None,
    ):
        """不分流地做完视觉前向与解码（对照用，也是重叠失败时的退路）。"""
        if not prepared:
            return []
        started = time.perf_counter()
        batch = self.build_batch(prepared)
        inputs_embeds = self.merge_batch(batch)
        return self.run_batch(
            batch, inputs_embeds, max_new_tokens, model_options, limits, started
        )


class _PendingRequest:
    """排队中的一次视觉问答请求。key 相同的请求才会被合进同一批。"""

    __slots__ = ("prompt", "image", "limit", "options", "key", "future", "hint")

    def __init__(
        self,
        prompt: str,
        image,
        limit: int,
        options: Dict[str, float],
        kind: Optional[str] = None,
        hint: float = 0.0,
    ) -> None:
        self.prompt = prompt
        self.image = image
        self.limit = limit
        self.options = options
        # 同类任务的输出长度分布接近，适合合批；提示词本身只是“同类”的粗判据，
        # 阅读序的提示词里带本页 bbox（每页都不同），所以允许调用方给一个
        # 更粗的类别名，否则每页都会各自成批。
        self.key = (kind or prompt, tuple(sorted(options.items())))
        # 输出长度的粗估（裁切像素面积 / 提示词长度）。同一批里行是同步推进的，
        # 短请求会被最长的那条拖着一起解码，按 hint 把长度接近的排到一起能显著
        # 减少“陪跑”的步数。只是排序提示，不参与缓存键，也不影响单行结果。
        self.hint = hint
        self.future: Optional[asyncio.Future] = None


def _length_hint(prompt: str, image, kind: Optional[str]) -> float:
    """粗估一次请求会生成多少 token，只用来把长度接近的请求排进同一批。

    识别请求的输出长度几乎正比于裁切面积（本机实测 89 个块：面积与输出字符数
    的 Spearman 相关 0.96），阅读序的输出长度则正比于本页块数（提示词里逐个
    列了 bbox）。估错没有正确性影响，只是合批不够紧凑。
    """
    if kind == "reading_order":
        return float(len(prompt))
    shape = getattr(image, "shape", None)
    if shape is not None and len(shape) >= 2:
        return float(shape[0]) * float(shape[1])
    return float(len(prompt))


class LocalOXRClient:
    """替换 OXRModelClient：把本机模型调用聚合成批，单卡流水执行。

    流水线本身是按“HTTP 服务端有连续批处理”设计的，每个 block 发一次请求。
    本地推理如果照单全收就退化成几千次 batch=1 的生成，而 batch=1 时每生成
    一个 token 都要把整份权重读一遍显存。这里用三个常驻协程把请求攒起来：

        _prepare_loop  : 按 key 取一组请求 -> 线程池里做 CPU 预处理
        _vision_loop   : 取预处理结果 -> 视觉塔前向（独立 CUDA 流）
        _generate_loop : 取已合并的嵌入 -> 预填充 + 解码

    _prepare_loop 与 _vision_loop 的队列深度都是 2，所以第 N 批在 GPU 上解码
    的时候，第 N+1 批的视觉塔已经在另一条流上跑完了 —— 视觉塔是算力受限、
    解码是带宽受限，两者在 GPU 上能互相填空（16 页实测视觉塔本来要占识别
    时间的 31%）。
    """

    def __init__(self, backend: OXRTextBackend, config, batch_size: int = 4) -> None:
        self.backend = backend
        self.config = config
        self.batch_size = max(1, int(batch_size))
        self._pending: List[_PendingRequest] = []
        self._condition = asyncio.Condition()
        self._prepared: Optional[asyncio.Queue] = None
        self._embedded: Optional[asyncio.Queue] = None
        self._workers: List[asyncio.Task] = []
        self._worker_loop: Optional[asyncio.AbstractEventLoop] = None
        self._prepare_pool = ThreadPoolExecutor(
            max_workers=1, thread_name_prefix="oxr-prepare"
        )
        self._prepare_pool_closed = False
        self._vision_stream = None
        self._semaphore: Optional[asyncio.Semaphore] = None
        self._semaphore_loop: Optional[asyncio.AbstractEventLoop] = None

    def get_request_semaphore(self) -> asyncio.Semaphore:
        """流水线用它限制在途请求数。

        本地实现里真正的串行点是对 GPU 的独占，用 1 会让请求永远无法攒成
        一批；这里放宽到批大小的 8 倍，只作为内存/背压上限。
        """
        loop = asyncio.get_running_loop()
        # 只在途请求数上限，同时也是识别阶段的内存背压。本地实现里独占 GPU 的
        # 是生成线程（asyncio.to_thread 派遣，不占事件循环），卡住请求数只会让
        # 攒批攒不满（16 页规范有 123 个块，太小的上限队列经常只剩一两个同类
        # 请求），所以放宽到批大小的 8 倍：排队的请求只占着已经裁好的图，不再
        # 额外占显存。
        limit = max(self.batch_size * 8, 16)
        if self._semaphore_loop is not loop or self._semaphore is None:
            self._semaphore = asyncio.Semaphore(limit)
            self._semaphore_loop = loop
        return self._semaphore

    # -- 常驻协程 ---------------------------------------------------------- #

    def _ensure_workers(self) -> None:
        loop = asyncio.get_running_loop()
        if (
            self._workers
            and self._worker_loop is loop
            and not any(task.done() for task in self._workers)
        ):
            return
        if self._workers:
            # 上一轮协程属于别的（已经关闭的）事件循环，或者中途退出了。
            # 队列和协程都绑在那个循环上，只能整套重建；否则新请求会静静地
            # 躺在队列里没人取，表现为调用方永远等下去（不是报错，是挂死）。
            LOGGER.debug("事件循环已更换，重建 OXR 客户端的工作协程")
            self._abandon_workers()
        self._condition = asyncio.Condition()
        self._prepared = asyncio.Queue(maxsize=2)
        self._embedded = asyncio.Queue(maxsize=2)
        if self._prepare_pool_closed:
            # aclose() 关掉的线程池不能再用，换一个新的。
            self._prepare_pool = ThreadPoolExecutor(
                max_workers=1, thread_name_prefix="oxr-prepare"
            )
            self._prepare_pool_closed = False
        if self.backend.can_overlap_vision() and self._vision_stream is None:
            self._vision_stream = self.backend.torch.cuda.Stream(
                device=self.backend.device
            )
        self._worker_loop = loop
        self._workers.append(loop.create_task(self._prepare_loop()))
        self._workers.append(loop.create_task(self._vision_loop()))
        self._workers.append(loop.create_task(self._generate_loop()))

    def _abandon_workers(self) -> None:
        """丢弃上一轮事件循环留下的协程与排队请求（这里不能 await）。"""
        for task in self._workers:
            if not task.done():
                task.cancel()
        self._workers.clear()
        self._worker_loop = None
        for request in self._pending:
            self._reject(request, RuntimeError("客户端已更换事件循环"))
        self._pending.clear()

    async def _next_group(self) -> List[_PendingRequest]:
        """取出下一批要推理的请求：同类、且预估输出长度接近。

        生成阶段是瓶颈、而且批大小几乎不影响单步耗时（CUDA Graph 下批 1 与
        批 4 只差几毫秒），所以这里愿意多等一小会儿、并优先挑同类里排队最多
        的，把一批攒满：一份规范类文档 16 页有 123 个块，按“先到先得”只能
        凑出 68 批（平均 1.8 行），预填充和补行的开销全白花。
        """
        async with self._condition:
            while not self._pending:
                await self._condition.wait()
            deadline = time.monotonic() + GROUP_WAIT_SECONDS
            # 一个 chunk 的识别请求几乎同时到达，先等这一波到齐才有得挑。
            # 判据是「队列连着 30 ms 没再增长」或「已够两批」，避免 GPU 空转；
            # 只看是否凑满一批的话，第一批永远是从最先到的 16 条里硬凑。
            target = self.batch_size * 2
            last_count = len(self._pending)
            while True:
                key, count = self._busiest_key()
                if count >= target:
                    break
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    break
                try:
                    await asyncio.wait_for(
                        self._condition.wait(), timeout=min(0.03, remaining)
                    )
                except asyncio.TimeoutError:
                    if len(self._pending) == last_count:
                        break
                last_count = len(self._pending)
            key, _count = self._busiest_key()
            candidates = [request for request in self._pending if request.key == key]
            group = self._best_length_group(candidates)
            chosen = {id(request) for request in group}
            self._pending = [
                request for request in self._pending if id(request) not in chosen
            ]
            return group

    def _best_length_group(
        self, candidates: Sequence[_PendingRequest]
    ) -> List[_PendingRequest]:
        """在同一类请求里挑出「长度接近、且值得先做」的一组。

        一批里的行是同步推进的：批的解码步数等于最长那条的输出长度，短请求
        全程陪跑。按 hint 排序后，把序列切成若干**连续**的组，目标是各组
        最大 hint 之和最小（每个组的代价就是这个最大值），组大小不超过
        `batch_size`、也不小于它的一半（太小的批预填充开销划不来）。
        用动态规划求这个划分，每次只交出第一组，剩下的留到下一次再划。
        """
        if len(candidates) <= self.batch_size:
            return list(candidates)
        if not LENGTH_AWARE_BATCH:
            return list(candidates[: self.batch_size])
        ordered = sorted(candidates, key=lambda request: request.hint)
        total = len(ordered)
        maximum = self.batch_size
        minimum = max(1, self.batch_size // 2)
        best = [float("inf")] * (total + 1)
        step = [0] * (total + 1)
        best[0] = 0.0
        for end in range(1, total + 1):
            # 排好序后组内最大值就是最后一个元素的 hint
            group_max = ordered[end - 1].hint
            for size in range(1, min(maximum, end) + 1):
                start = end - size
                # 尾组也不例外：总长都能划成若干 [minimum, maximum] 的组合，
                # 用不着为"不满一组"开后门（穷举过 17~64 都能覆盖）。
                if size < minimum:
                    continue
                cost = best[start] + group_max
                if cost < best[end]:
                    best[end] = cost
                    step[end] = size
        # 回溯出划分；从尾部往前推，最后记下的那组就是最前面（hint 最小）的一组
        sizes = []
        end = total
        while end > 0:
            # 防御性兜底：DP 下 end 必然可达，万一没覆盖就拿剩下的全部当一组
            size = step[end] or min(maximum, end)
            sizes.append(size)
            end -= size
        return ordered[: sizes[-1]]

    def _busiest_key(self) -> Tuple[Any, int]:
        """排队最多的请求类型（并列时取最早到达的，保证先进先出）。"""
        counts: Dict[Any, int] = {}
        for request in self._pending:
            counts[request.key] = counts.get(request.key, 0) + 1
        best_key = self._pending[0].key if self._pending else None
        best = counts.get(best_key, 0)
        for request in self._pending:
            count = counts[request.key]
            if count > best:
                best_key, best = request.key, count
        return best_key, best

    async def _prepare_loop(self) -> None:
        loop = asyncio.get_running_loop()
        while True:
            try:
                group = await self._next_group()
                items = [(request.prompt, request.image) for request in group]
                try:
                    prepared, failures = await loop.run_in_executor(
                        self._prepare_pool, self.backend.prepare_batch, items
                    )
                except Exception as error:  # noqa: BLE001 - 单批失败不影响后续
                    LOGGER.error("请求预处理失败：%s", error)
                    for request in group:
                        self._reject(request, error)
                    continue
                failed_positions = set()
                for position, error in failures:
                    failed_positions.add(position)
                    self._reject(group[position], error)
                # prepared 与 succeeded 都按输入顺序排列，两者一一对应。
                succeeded = [
                    group[position]
                    for position in range(len(group))
                    if position not in failed_positions
                ]
                if not succeeded:
                    continue
                await self._prepared.put((succeeded, prepared))
            except asyncio.CancelledError:
                raise
            except Exception as error:  # noqa: BLE001 - 工作协程不能退出
                LOGGER.exception("预处理协程异常：%s", error)

    def _encode_group(self, prepared: Sequence[Any]):
        """视觉塔前向 + 文本嵌入合并，返回 (batch, embeds, 完成事件)。

        整段提交到视觉那条流上并记一个事件，解码前等这个事件即可。视觉流
        已经因为报错被关掉时，退回当前流（等于串行），保证不出错。
        """
        backend = self.backend
        batch = backend.build_batch(prepared)
        stream = self._vision_stream
        if stream is None:
            return batch, backend.merge_batch(batch), None
        torch = backend.torch
        with _CUDA_CAPTURE_LOCK:
            with torch.cuda.stream(stream):
                embeds = backend.merge_batch(batch)
                event = torch.cuda.Event()
                event.record(stream)
        return batch, embeds, event

    async def _vision_loop(self) -> None:
        loop = asyncio.get_running_loop()
        while True:
            try:
                group, prepared = await self._prepared.get()
                batch = embeds = event = None
                if self._vision_stream is not None:
                    try:
                        batch, embeds, event = await loop.run_in_executor(
                            None, self._encode_group, prepared
                        )
                    except Exception as error:  # noqa: BLE001 - 之后退回串行路径
                        LOGGER.warning("视觉前向失败，之后退回串行路径：%s", error)
                        self._vision_stream = None
                        batch = embeds = event = None
                await self._embedded.put((group, prepared, batch, embeds, event))
            except asyncio.CancelledError:
                raise
            except Exception as error:  # noqa: BLE001 - 工作协程不能退出
                LOGGER.exception("视觉协程异常：%s", error)

    async def _generate_loop(self) -> None:
        while True:
            try:
                group, prepared, batch, embeds, event = await self._embedded.get()
                limits = [request.limit for request in group]
                options = group[0].options
                try:
                    if batch is None:  # 视觉塔没跑成，整批走串行路径
                        outputs = await asyncio.to_thread(
                            self.backend.generate_batch,
                            prepared,
                            max(limits),
                            options,
                            limits,
                        )
                    else:
                        outputs = await asyncio.to_thread(
                            self.backend.run_batch,
                            batch,
                            embeds,
                            max(limits),
                            options,
                            limits,
                            time.perf_counter(),
                            event,
                        )
                except Exception as error:  # noqa: BLE001 - 单批失败不影响后续
                    LOGGER.error("批量生成失败（%d 条）：%s", len(group), error)
                    for request in group:
                        self._reject(request, error)
                    continue
                for request, output in zip(group, outputs):
                    self._resolve(request, output)
            except asyncio.CancelledError:
                raise
            except Exception as error:  # noqa: BLE001 - 工作协程不能退出
                LOGGER.exception("生成协程异常：%s", error)

    @staticmethod
    def _resolve(request: _PendingRequest, value) -> None:
        future = request.future
        if future is not None and not future.done():
            try:
                future.set_result(value)
            except RuntimeError:  # 事件循环已经关了，回调没人收
                pass

    @staticmethod
    def _reject(request: _PendingRequest, error: BaseException) -> None:
        future = request.future
        if future is not None and not future.done():
            try:
                future.set_exception(error)
            except RuntimeError:  # 事件循环已经关了，回调没人收
                pass

    # -- 与流水线对接 ------------------------------------------------------ #

    async def aclose(self) -> None:
        for task in self._workers:
            task.cancel()
        for task in self._workers:
            try:
                await task
            except (asyncio.CancelledError, Exception):  # noqa: BLE001
                pass
        self._workers.clear()
        self._worker_loop = None
        for request in self._pending:
            self._reject(request, RuntimeError("客户端已关闭"))
        self._pending.clear()
        self._prepare_pool.shutdown(wait=False)
        self._prepare_pool_closed = True

    async def _call_vision_model(
        self,
        prompt: str,
        image,
        timeout: int,
        model_options: Optional[Dict[str, Any]] = None,
        max_tokens: Optional[int] = None,
        kind: Optional[str] = None,
    ):
        from oxr.model.client import resolve_model_options

        del timeout  # 本地推理没有网络超时
        if image is None:
            raise ValueError("An image is required for vision model calls.")

        self._ensure_workers()
        options = resolve_model_options(model_options)
        limit = int(max_tokens or self.config.max_tokens or 8192)
        request = _PendingRequest(
            prompt, image, limit, dict(options), kind, _length_hint(prompt, image, kind)
        )
        request.future = asyncio.get_running_loop().create_future()

        async with self._condition:
            self._pending.append(request)
            self._condition.notify_all()
        return await request.future

    # 以下方法复用 oxr.model.client.OXRModelClient 的实现（prompt 与参数一致），
    # 通过继承在 _install 中绑定即可。
    async def request_reading_order(self, prompt, image, max_tokens):
        # 阅读序的提示词里带本页 bbox，逐页不同；按类别合批才能把整章页面
        # 拼成一批，否则每页一条、每条都要按 16 行补齐做一遍预填充。
        return await self._call_vision_model(
            prompt=prompt, image=image, timeout=self.config.timeout,
            model_options=None, max_tokens=max_tokens, kind="reading_order",
        )

    async def recognize_text(self, image, model_options=None):
        from oxr.model.prompts import PROMPT_TEXT_RECOGNITION
        return await self._call_vision_model(
            PROMPT_TEXT_RECOGNITION, image, self.config.timeout, model_options,
        )

    async def recognize_formula(self, image, model_options=None):
        from oxr.model.prompts import PROMPT_FORMULA_RECOGNITION
        return await self._call_vision_model(
            PROMPT_FORMULA_RECOGNITION, image, self.config.timeout, model_options,
        )

    async def recognize_table(self, image, model_options=None):
        from oxr.model.prompts import PROMPT_TABLE_RECOGNITION
        return await self._call_vision_model(
            PROMPT_TABLE_RECOGNITION, image, self.config.timeout, model_options,
        )

    async def recognize_code(self, image, model_options=None):
        from oxr.model.prompts import PROMPT_CODE_RECOGNITION
        return await self._call_vision_model(
            PROMPT_CODE_RECOGNITION, image, self.config.timeout, model_options,
        )

    async def recognize_chemical_structure(self, image, model_options=None):
        from oxr.model.prompts import PROMPT_CHEMICAL_STRUCTURE_RECOGNITION
        return await self._call_vision_model(
            PROMPT_CHEMICAL_STRUCTURE_RECOGNITION, image, self.config.timeout,
            model_options,
        )


# --------------------------------------------------------------------------- #
# 本地版面分析
# --------------------------------------------------------------------------- #


class LocalLayoutRuntime:
    """用 services/layout_analyze 的 D-FINE 代码在本地推理。"""

    def __init__(
        self, model_path: Path, device_name: str, stats: Optional[RunStats] = None
    ) -> None:
        from dfine_la.config import Settings
        from dfine_la.runtime import ModelRuntime

        settings = Settings(
            model_path=model_path,
            device=device_name,
            cuda_graph_enabled=False,
            parallel_preprocess=False,
        ).validate_checkpoint()
        LOGGER.info("正在加载版面分析模型：%s（%s）", model_path, device_name)
        self._runtime = ModelRuntime(settings)
        self.stats = stats

    def infer(self, images: Sequence[Any]) -> Dict[str, Any]:
        """直接吃 ndarray，省掉 PNG 编码再解码的往返。

        原实现把页面编成 PNG 传给“服务端”，本地推理根本不需要这层序列化，
        一张 1240x1754 的页面编解码各要几十毫秒。
        """
        import cv2

        from dfine_la.confidence import format_predictions
        from dfine_la.preprocess import preprocess_images

        started = time.perf_counter()
        rgb = [cv2.cvtColor(image, cv2.COLOR_BGR2RGB) for image in images]
        ready = preprocess_images(rgb, self._runtime.settings.image_size)
        outputs, _timings = self._runtime.infer(ready)
        # 阈值过滤和 NMS 由 oxr.pipeline.postprocess.layout_postprocess 统一完成。
        if self.stats is not None:
            self.stats.add_layout(len(images), time.perf_counter() - started)
        return format_predictions(outputs)


class LocalLayoutClient:
    """替换 LayoutAnalyzeClient：本地调用 D-FINE，而不是 HTTP 服务。"""

    def __init__(self, runtime: LocalLayoutRuntime, config) -> None:
        self.runtime = runtime
        self.config = config
        self._lock = asyncio.Lock()

    def encode_images(self, images: Sequence[Any]) -> List[Any]:
        # 名字沿用上游接口，但本地实现不再做 PNG 编码：直接传 ndarray。
        return list(images)

    async def request_batch(self, encoded_images: Sequence[bytes]):
        async with self._lock:
            return await asyncio.to_thread(self.runtime.infer, encoded_images)

    async def aclose(self) -> None:
        return None


# --------------------------------------------------------------------------- #
# 引擎
# --------------------------------------------------------------------------- #


def _resolve_device(name: str):
    import torch

    if name == "auto":
        return torch.device("cuda:0" if torch.cuda.is_available() else "cpu")
    if name == "cuda" and not torch.cuda.is_available():
        raise SystemExit("指定了 --device cuda，但当前环境没有可用的 CUDA。")
    return torch.device(name)


# 句子级终止标点：行尾是这些字符说明句子已结束，不与下一页续接。
# 逗号、分号、冒号、顿号都是句内停顿，句子仍在延续，允许跨页合并。
_TERMINAL_END_CHARS = "。？！….!?|>"


def _merge_adjacent_text_blocks(pages_blocks):
    """把同栏、纵向紧邻的纯文本块合并成更少、更高的裁切块。

    视觉塔按裁切面积计费，一行一块时每个块都要付一次视觉编码的固定开销。
    把相邻行合并成一块，请求数下降数倍，视觉开销随之下降；识别后按行自然
    成段，段落结构比逐行识别更合理。layout 返回的块是检测序而非垂直序，
    故先把纯文本块按 y 排序再贪心分组（同栏、组高不超上限），合并时拷贝
    新块并扩展 bbox，不改动 layout 原块对象（避免连带影响 JSON 输出的框）。
    """
    from dataclasses import replace

    from oxr.pipeline.models import BlockType

    # 组高上限（px）：约 10 行。太高则裁切缩到 512 后纵向拉伸失真，
    # 太低则请求数降得少。
    max_group_h = 250.0
    merged_pages = []
    for blocks in pages_blocks:
        text_blocks = [
            b for b in blocks
            if b.type == BlockType.TEXT and not b.generated_role
        ]
        non_text_blocks = [
            b for b in blocks
            if b.type != BlockType.TEXT or b.generated_role
        ]
        # 单栏文档按 y 即垂直序；多栏会交错，但 _mergeable 的水平重叠
        # 检查会阻止跨栏误并。
        text_blocks.sort(key=lambda b: (b.bbox[1] + b.bbox[3]) / 2)
        groups: list = []
        for b in text_blocks:
            if groups:
                g = groups[-1]
                g_h = max(x.bbox[3] for x in g) - min(x.bbox[1] for x in g)
                if (
                    _mergeable_text_blocks(g[-1], b)
                    and g_h + (b.bbox[3] - b.bbox[1]) <= max_group_h
                ):
                    g.append(b)
                    continue
            groups.append([b])
        out = []
        for g in groups:
            if len(g) == 1:
                out.append(replace(g[0], bbox=list(g[0].bbox)))
                continue
            x0 = min(b.bbox[0] for b in g)
            y0 = min(b.bbox[1] for b in g)
            x1 = max(b.bbox[2] for b in g)
            y1 = max(b.bbox[3] for b in g)
            out.append(replace(g[0], bbox=[x0, y0, x1, y1]))
        for b in non_text_blocks:
            out.append(replace(b, bbox=list(b.bbox)))
        # 按 y 中心排序，给识别/阅读序一个合理顺序
        out.sort(key=lambda b: (b.bbox[1] + b.bbox[3]) / 2)
        merged_pages.append(out)
    return merged_pages


def _mergeable_text_blocks(a, b) -> bool:
    """判断两个纯文本块可否纵向合并：同栏、间距小、合并后不过高（上下顺序无关）。"""
    ax0, ay0, ax1, ay1 = a.bbox
    bx0, by0, bx1, by1 = b.bbox
    # 同栏：水平重叠至少占较窄块宽度的一半
    overlap_x = min(ax1, bx1) - max(ax0, bx0)
    if overlap_x < 0.5 * min(ax1 - ax0, bx1 - bx0):
        return False
    # 纵向紧邻：间距不超过较短块高的一倍（也设 20px 绝对上限），上下顺序无关。
    # 同段行距小、段间空白大，一倍的行高可区分两者。
    gap = (by0 - ay1) if ay1 <= by0 else (ay0 - by1)
    line_h = min(ay1 - ay0, by1 - by0)
    if gap < 0 or gap > min(20.0, 1.0 * line_h):
        return False
    # 合并后总高度上限：避免裁切长边超过模型像素上限
    if max(ay1, by1) - min(ay0, by0) > 900.0:
        return False
    return True


def _last_text_width_pct(blocks) -> float:
    """阅读序中最后一个文本块的宽度占页宽的百分比（几何法）。

    末行若是标题，宽度占比明显偏小；正文行通常接近满宽。返回 0~1。
    """
    from oxr.pipeline.models import BlockType

    last_text = None
    for block in reversed(blocks):
        if block.type == BlockType.TEXT and not block.generated_role:
            last_text = block
            break
    if last_text is None:
        return 1.0
    page_width = max((b.bbox[2] for b in blocks), default=0.0)
    block_width = last_text.bbox[2] - last_text.bbox[0]
    return block_width / page_width if page_width > 0 else 1.0


def _estimate_width_pct(markdown: str) -> float:
    """用文本估算末行宽度占比：末行长度 / 该段最长行长度。

    跨 chunk 边界拿不到几何 bbox（默认不输出 json），用相对长度近似。
    """
    lines = [line for line in markdown.splitlines() if line.strip()]
    if not lines:
        return 1.0
    max_len = max(len(line) for line in lines)
    return len(lines[-1]) / max_len if max_len > 0 else 1.0


def _join_markdown_across_pages(parts):
    """跨页段落合并：页尾未以终止标点结束、末行不满宽且页首是延续内容时并段。

    parts 为 (markdown, 末行宽度占比) 元组序列。按页识别时，一个自然段落若跨页
    会被拆成两段，md 里出现不自然的中断。这里在页边界处做启发式合并，让跨页
    段落恢复为完整段落。同一函数既用于一个 chunk 内的页边界（包装 assemble_
    markdown），也用于 chunk 之间。
    """
    parts = [part for part in parts if part[0]]
    if len(parts) <= 1:
        return "".join(md for md, _ in parts)
    joined = [parts[0]]
    for part in parts[1:]:
        if _should_join_at_boundary(joined[-1][0], joined[-1][1], part[0]):
            # 去掉页尾的尾随空行与页首的前导空行，用单个换行连成同一段；
            # 合并后末行变为当前页的末行，宽度占比随之更新。
            joined[-1] = (
                joined[-1][0].rstrip("\n") + "\n" + part[0].lstrip("\n"),
                part[1],
            )
        else:
            joined.append(part)
    return "\n\n".join(md for md, _ in joined)


def _should_join_at_boundary(prev: str, prev_width_pct: float, cur: str) -> bool:
    """页尾行与页首行是否属于同一段落的延续。"""
    prev_lines = [line for line in prev.splitlines() if line.strip()]
    cur_lines = [line for line in cur.splitlines() if line.strip()]
    if not prev_lines or not cur_lines:
        return False
    last = prev_lines[-1].rstrip()
    first = cur_lines[0].lstrip()
    if not last or last[-1] in _TERMINAL_END_CHARS:
        return False
    # 末行宽度占比过小（偏标题/短片段）时不合并，避免把标题并进下一段
    if prev_width_pct < 0.5:
        return False
    # 页首是标题/表格/代码/公式/列表/图表等结构开头时不合并
    if first.startswith(("#", "|", "```", "$$", "<", "- ", "* ", "1. ", "1) ")):
        return False
    return True


class LocalEngine:
    """加载两个本地模型，并把它们接入仓库原有的流水线。"""

    def __init__(self, args: argparse.Namespace) -> None:
        self.args = args
        self.stats = RunStats()
        self._loaded = False

    def load(self) -> None:
        if self._loaded:
            return

        try:
            import cv2  # noqa: F401
            import sopdf  # noqa: F401
            import torch  # noqa: F401
            import transformers  # noqa: F401
            from oxr.config.settings import settings
            from oxr.pipeline.orchestrator import process_decoded_document
        except ImportError as exc:  # pragma: no cover - 依赖缺失时的提示
            raise SystemExit(
                "缺少依赖：%s\n请先安装：\n"
                "  pip install torch transformers sopdf openai pydantic-settings "
                "PyYAML httpx opencv-python\n"
                "GPU 版 torch 请参考 https://pytorch.org/get-started/locally/"
                % exc
            ) from exc

        import oxr.pipeline.layout as layout_module
        import oxr.pipeline.orchestrator as orchestrator_module
        import oxr.pipeline.reading_order as reading_order_module
        import oxr.pipeline.recognition as recognition_module

        model_dir = Path(self.args.model).expanduser().resolve()
        if not (model_dir / "config.json").is_file():
            raise SystemExit(
                f"模型目录无效：{model_dir}\n"
                "请先下载模型（约 3 GB）：\n"
                "  pip install modelscope\n"
                f"  modelscope download SoMark/OXR-1.0 --local_dir {DEFAULT_MODEL_DIR}"
            )
        layout_model = model_dir / "layout_analysis" / "last.pth"
        if not layout_model.is_file():
            raise SystemExit(f"版面分析权重不存在：{layout_model}")

        device = _resolve_device(self.args.device)
        if device.type == "cpu":
            LOGGER.warning("正在使用 CPU 推理，速度会非常慢。")

        if self.args.torch_threads > 0:
            torch.set_num_threads(int(self.args.torch_threads))
            cv2.setNumThreads(int(self.args.torch_threads))
            LOGGER.info("torch/cv2 的 CPU 线程数已设为 %d", self.args.torch_threads)

        # 先调整流水线参数，再加载模型。
        settings.oxr_model.max_tokens = int(self.args.max_tokens)
        settings.layout_analyze_model.batch_size = int(self.args.layout_batch_size)
        settings.layout_analyze_model.max_concurrency = max(
            1, int(self.args.layout_concurrency)
        )
        # 打开后 orchestrator 会在 metadata 里带上各阶段耗时，用于进度输出。
        settings.debug = True

        text_backend = OXRTextBackend(
            model_dir,
            device,
            self.stats,
            graph_decode=self.args.graph_decode != "off",
            graph_batch=max(1, int(self.args.batch_size)),
            graph_cache_len=max(256, int(self.args.graph_cache_len)),
        )
        oxr_client = LocalOXRClient(
            text_backend, settings.oxr_model, self.args.batch_size
        )
        layout_client = LocalLayoutClient(
            LocalLayoutRuntime(layout_model, device.type, self.stats),
            settings.layout_analyze_model,
        )

        # 各流水线模块在 import 时绑定了全局客户端，这里统一替换为本地实现。
        layout_module.layout_client = layout_client
        orchestrator_module.model_client = oxr_client
        reading_order_module.model_client = oxr_client
        recognition_module.model_client = oxr_client

        # 阅读序：版面模型不提供阅读序时，上游编排器会把每页整页图片再送
        # 模型推理一轮（每页一次整页请求，视觉开销约为文本块的 5~6 倍）。
        # 默认走上游原行为保证阅读序质量；--reading-order geo 改用 XY-Cut++
        # 几何排序（零 GPU 开销，可提速约 1/6，但多栏/复杂版式有局限）。
        if self.args.reading_order == "geo":
            from oxr.pipeline.xycut_order import xycut_plus_order

            async def _geometric_reading_orders(pages_blocks, pages_images, semaphore, **kwargs):
                orders = []
                for blocks in pages_blocks:
                    try:
                        orders.append(xycut_plus_order(blocks))
                    except Exception:
                        LOGGER.warning("XY-Cut++ 未能产生阅读序，该页回退为版面顺序")
                        orders.append(list(range(len(blocks))))
                return orders

            # orchestrator 在模块顶部直接 from-import 了该函数，
            # 两个命名空间都要替换才生效
            reading_order_module.resolve_reading_orders = _geometric_reading_orders
            orchestrator_module.resolve_reading_orders = _geometric_reading_orders

        # --- #1 合并相邻文本块：更少、更高的裁切，视觉开销下降数倍 ---
        # orchestrator 在模块顶部 from-import 了这两个函数，替换其命名空间即可。
        # 合并在识别/阅读序之前统一做，两边看到的是同一份合并结果。
        _orig_recognize_pages = orchestrator_module.recognize_pages
        _orig_resolve_reading_orders = orchestrator_module.resolve_reading_orders

        def _wrap_with_merge(async_fn):
            async def _wrapped(pages_blocks, *args, **kwargs):
                return await async_fn(
                    _merge_adjacent_text_blocks(pages_blocks), *args, **kwargs
                )

            return _wrapped

        orchestrator_module.recognize_pages = _wrap_with_merge(_orig_recognize_pages)
        orchestrator_module.resolve_reading_orders = _wrap_with_merge(
            _orig_resolve_reading_orders
        )

        # --- #3 裁切去重缓存：相同裁切（页眉/页脚/页码）只识别一次 ---
        import copy as _copy
        import hashlib as _hashlib
        import numpy as _np
        from oxr.utils.image import crop_image as _crop_image

        _orig_recognize_block = recognition_module.recognize_block
        # 键是「裁切像素 + 块类型」，值只存识别出的正文与格式。不能把整个
        # Block 存下来复用：Block 里的 bbox / text_before / text_after 属于被
        # 命中的那一个块，复用时会把这套上下文带到别的块上（实测同一份 PDF
        # 跑第二遍会多出一个重复的标题块）。缓存归 LocalEngine 管，每次
        # process_one 开头清空，避免跨文档、跨参数复用。
        self._crop_cache: Dict[bytes, Tuple[str, Any]] = {}

        async def _cached_recognize_block(
            block, page_image, sem, file_name, page_num,
            model_options=None, element_formats=None, table_images=None,
        ):
            # PICTURE 块要落盘图片文件、表格占位图裁切不固定，二者不缓存
            from oxr.pipeline.models import BlockType as _BlockType

            if block.type == _BlockType.PICTURE or table_images is not None:
                return await _orig_recognize_block(
                    block, page_image, sem, file_name, page_num,
                    model_options, element_formats, table_images,
                )
            crop = _crop_image(page_image, block.bbox)
            key = (
                _hashlib.md5(_np.ascontiguousarray(crop).tobytes()).digest()
                + block.type.value.encode("utf-8")
            )
            cached = self._crop_cache.get(key)
            if cached is not None:
                # 只借用识别结果，bbox / idx / 跨块上下文一律沿用当前块
                hit = _copy.deepcopy(block)
                hit.content, hit.format = cached
                return hit
            result = await _orig_recognize_block(
                block, page_image, sem, file_name, page_num,
                model_options, element_formats, table_images,
            )
            if len(self._crop_cache) < 2048:
                self._crop_cache[key] = (result.content, result.format)
            return result

        recognition_module.recognize_block = _cached_recognize_block

        # --- #4 跨页段落合并：逐页组装后在页边界处合并跨页段落 ---
        # 贴近原 assemble_markdown：逐页 blocks_to_markdown + 整体 space_markdown_math，
        # 只在页边界处插入段落合并。
        from oxr.pipeline.postprocess.assembly_postprocess import (
            blocks_to_markdown as _blocks_to_markdown,
        )
        from oxr.pipeline.postprocess.formula_postprocess import (
            space_markdown_math as _space_markdown_math,
        )

        def _merged_assemble_markdown(pages_blocks):
            page_parts = [
                (
                    _blocks_to_markdown(blocks),
                    _last_text_width_pct(blocks),
                )
                for blocks in pages_blocks
            ]
            return _space_markdown_math(_join_markdown_across_pages(page_parts))

        orchestrator_module.assemble_markdown = _merged_assemble_markdown

        # --- #2 chunk 级重叠：本 chunk 识别时预取下一 chunk 的版面 ---
        self._layout_prefetched: Dict[int, Any] = {}
        self._orig_analyze_layouts = orchestrator_module.analyze_layouts

        async def _prefetched_analyze_layouts(images, model_options):
            prefetched = self._layout_prefetched.pop(id(images), None)
            if prefetched is not None:
                return prefetched
            return await self._orig_analyze_layouts(images, model_options)

        orchestrator_module.analyze_layouts = _prefetched_analyze_layouts

        self.settings = settings
        self.process_decoded_document = process_decoded_document
        self._loaded = True


# --------------------------------------------------------------------------- #
# 文件处理
# --------------------------------------------------------------------------- #


def collect_inputs(args: argparse.Namespace) -> Tuple[List[Path], Optional[Path]]:
    if args.directory:
        root = Path(args.directory).expanduser().resolve()
        if not root.is_dir():
            raise SystemExit(f"目录不存在：{root}")
        pattern = "**/*" if args.recursive else "*"
        files = sorted(
            path
            for path in root.glob(pattern)
            if path.is_file() and path.suffix.lower() == ".pdf"
        )
        if not files:
            raise SystemExit(f"目录中没有找到 PDF：{root}")
        return files, root

    pdf = Path(args.pdf).expanduser().resolve()
    if not pdf.is_file():
        raise SystemExit(f"文件不存在：{pdf}")
    if pdf.suffix.lower() != ".pdf":
        raise SystemExit(f"只支持 PDF 文件：{pdf}")
    return [pdf], None


def output_base(
    pdf_path: Path, input_root: Optional[Path], out_dir: Optional[Path]
) -> Path:
    if out_dir is None:
        return pdf_path.with_suffix("")
    if input_root is not None:
        relative = pdf_path.relative_to(input_root)
    else:
        relative = Path(pdf_path.name)
    base = (out_dir / relative).with_suffix("")
    base.parent.mkdir(parents=True, exist_ok=True)
    return base


def expected_outputs(base: Path, formats: Iterable[str]) -> List[Path]:
    return [
        base.with_suffix(".md") if fmt == "markdown" else base.with_suffix(".json")
        for fmt in formats
    ]


def write_outputs(base: Path, outputs: Dict[str, Any], formats: Iterable[str]) -> List[Path]:
    written: List[Path] = []

    if "markdown" in formats:
        markdown_path = base.with_suffix(".md")
        markdown_path.write_text(
            str(outputs.get("markdown", "")), encoding="utf-8", newline="\n"
        )
        written.append(markdown_path)

    if "json" in formats:
        json_path = base.with_suffix(".json")
        # 与官方 API 的 outputs.json 结构一致：{"pages": [...]}
        payload = outputs.get("json", {"pages": []})
        json_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
            newline="\n",
        )
        written.append(json_path)

    return written


def _debug_timing(result: Dict[str, Any]) -> Tuple[Dict[str, float], int]:
    metadata = result.get("metadata") or {}
    timing = metadata.get("debug_timing") or {}
    stats = metadata.get("debug_stats") or {}
    return timing, int(stats.get("total_blocks", 0))


def _pdf_page_count(pdf_bytes: bytes) -> int:
    import sopdf

    with sopdf.open(stream=pdf_bytes) as doc:
        return len(doc.pages)


def _render_pdf_pages(pdf_bytes: bytes, dpi: int, indices: Sequence[int]) -> List[Any]:
    """只渲染指定页（0 起下标），逐页解码成 BGR ndarray。

    整本一次性渲染时，所有页的位图会一直占着内存（150 dpi 约 6.5 MB/页，
    一百页就是 650 MB），而且渲染全部页之后才算真正开始推理。按 chunk 渲染
    可以把渲染藏进上一个 chunk 的 GPU 时间里，峰值内存也只和 chunk 有关。
    """
    import cv2
    import numpy as np
    import sopdf

    images: List[Any] = []
    with sopdf.open(stream=pdf_bytes) as doc:
        pages = list(doc.pages)
        for index in indices:
            if index < 0 or index >= len(pages):
                continue
            buffer = np.frombuffer(pages[index].render(dpi=dpi), np.uint8)
            images.append(cv2.imdecode(buffer, cv2.IMREAD_COLOR))
    return images


async def process_one(
    engine: LocalEngine,
    pdf_path: Path,
    base: Path,
    formats: Sequence[str],
    args: argparse.Namespace,
    prefix: str = "",
) -> List[Path]:
    started = time.perf_counter()
    # 裁切去重缓存只在本次文件内有效：跨文件/跨次复用它，等于拿上一次的
    # 识别结果替代本次推理，既掩盖了真实开销，也会带进上一次的块上下文。
    cache = getattr(engine, "_crop_cache", None)
    if cache is not None:
        cache.clear()
    pdf_bytes = pdf_path.read_bytes()
    total_pages = await asyncio.to_thread(_pdf_page_count, pdf_bytes)
    if not total_pages:
        raise ValueError("PDF 没有可渲染的页面")

    selected = (
        select_page_indices(total_pages, args.pages)
        if args.pages
        else list(range(total_pages))
    )
    if not selected:
        raise ValueError(f"页码范围没有命中任何页（共 {total_pages} 页）")

    # PDF 里的图片按 pptx2md 的习惯导出到同级的 <文件名>_media/，
    # md/json 里写相对路径，搬走 md 时连同该目录一起搬。
    media_dir = base.parent / f"{base.name}_media"
    # 插图用 uuid 命名，上一次转换留下的文件本次不会再被引用，成功后清掉。
    stale_media = (
        {path.name for path in media_dir.glob("*.png")} if media_dir.is_dir() else set()
    )

    # 按页分块送进流水线：显存/内存峰值只和块大小有关，而且每块都能落盘，
    # 中断后已完成的页不会白跑。整篇一次性处理时，recognition 阶段会先把
    # 几千张裁剪图全部 copy 进内存，很容易把机器压到换页。
    chunk_size = max(1, int(args.chunk_pages))
    chunk_count = (len(selected) + chunk_size - 1) // chunk_size
    chunk_pages_list = [
        selected[offset:offset + chunk_size]
        for offset in range(0, len(selected), chunk_size)
    ]
    markdown_parts: List[Tuple[str, float]] = []
    json_pages: List[Dict[str, Any]] = []
    totals = {"layout": 0.0, "recognition": 0.0, "assembly": 0.0}
    block_total = 0
    written: List[Path] = []
    render_seconds = 0.0
    pending_chunk_task: Optional[asyncio.Task] = None

    async def _render_and_prefetch(pages: Sequence[int]) -> List[Any]:
        """渲染一个 chunk 的页面，紧接着把它的版面分析也做掉。

        渲染是纯 CPU（PDFium）、版面是 GPU，两者与上一个 chunk 的识别并行，
        所以这一段的墙钟时间基本被前一个 chunk 的推理吃掉。返回的列表对象
        既是流水线的输入，也是版面预取结果的 id() 键，两边必须是同一个对象。
        """
        images = await asyncio.to_thread(_render_pdf_pages, pdf_bytes, args.dpi, pages)
        try:
            engine._layout_prefetched[id(images)] = await engine._orig_analyze_layouts(
                images, None
            )
        except Exception as error:  # noqa: BLE001 - 预取失败就回退常规计算
            engine._layout_prefetched.pop(id(images), None)
            LOGGER.warning("版面预取失败，本块回退为常规分析：%s", error)
        return images

    for chunk_no, chunk in enumerate(chunk_pages_list, start=1):
        chunk_started = time.perf_counter()
        if pending_chunk_task is not None:
            chunk_images = await pending_chunk_task
            pending_chunk_task = None
        else:
            render_started = time.perf_counter()
            chunk_images = await asyncio.to_thread(
                _render_pdf_pages, pdf_bytes, args.dpi, chunk
            )
            render_seconds += time.perf_counter() - render_started
        # 与本 chunk 的识别并行：预渲染 + 预分析下一个 chunk。
        if chunk_no < len(chunk_pages_list):
            pending_chunk_task = asyncio.create_task(
                _render_and_prefetch(chunk_pages_list[chunk_no])
            )
        result = await engine.process_decoded_document(
            chunk_images,
            pdf_path.name,
            list(formats),
            keep_header_footer=args.keep_header_footer,
            asset_dir=media_dir,
            asset_url_prefix=f"./{base.name}_media",
        )
        outputs = result.get("outputs") or {}

        timing, blocks = _debug_timing(result)
        block_total += blocks
        chunk_timing = {
            key: float(timing.get(f"{key}_time_ms", 0)) / 1000.0 for key in totals
        }
        for key, seconds in chunk_timing.items():
            totals[key] += seconds

        # 分块后 page_num 从 0 重来，这里换回原文档的页号。
        if "json" in outputs:
            pages = outputs["json"].get("pages") or []
            for local_index, page in enumerate(pages):
                page["page_num"] = chunk[local_index]
            json_pages.extend(pages)
            outputs["json"] = {"pages": json_pages}
        if "markdown" in outputs:
            md = outputs["markdown"]
            # 跨 chunk 拿不到几何 bbox，用文本相对长度估算末行宽度占比。
            markdown_parts.append((md, _estimate_width_pct(md)))
            # 跨 chunk 的页边界同样做段落合并（与 chunk 内 assemble 一致）。
            outputs["markdown"] = _join_markdown_across_pages(markdown_parts)

        # 增量落盘：长文档跑到一半也能先看结果。
        written = write_outputs(base, outputs, formats)

        # 本块的页面位图只在这一轮用得到，及早释放（渲染下一块时会重新解码）。
        chunk_images.clear()

        elapsed = time.perf_counter() - chunk_started
        detail = "".join(
            f"，{name} {chunk_timing[key]:.1f}s"
            for name, key in (("版面", "layout"), ("识别", "recognition"), ("组装", "assembly"))
            if timing
        )
        print(
            f"{prefix} 页 {chunk[0] + 1}-{chunk[-1] + 1}/{total_pages}"
            f"（第 {chunk_no}/{chunk_count} 批）{elapsed:.1f}s"
            + (f"，块 {blocks}" if blocks else "")
            + detail,
            flush=True,
        )

    for name in stale_media:
        (media_dir / name).unlink(missing_ok=True)
    if media_dir.is_dir() and not any(media_dir.iterdir()):
        media_dir.rmdir()

    print(
        f"{prefix} 完成：{len(selected)}/{total_pages} 页，"
        f"渲染 {render_seconds:.1f}s，块 {block_total}，"
        f"合计 {time.perf_counter() - started:.1f}s",
        flush=True,
    )
    return written


async def run(args: argparse.Namespace) -> int:
    args.formats = list(dict.fromkeys(args.formats))
    files, input_root = collect_inputs(args)
    out_dir = Path(args.out).expanduser().resolve() if args.out else None
    if out_dir is not None:
        out_dir.mkdir(parents=True, exist_ok=True)

    jobs = [(pdf_path, output_base(pdf_path, input_root, out_dir)) for pdf_path in files]
    if args.skip_existing:
        pending = [
            job
            for job in jobs
            if not all(path.exists() for path in expected_outputs(job[1], args.formats))
        ]
    else:
        pending = jobs
    skipped = len(jobs) - len(pending)

    if skipped:
        print(f"共 {len(jobs)} 个 PDF，跳过已存在 {skipped} 个。", flush=True)
    if not pending:
        print("没有需要处理的 PDF。", flush=True)
        return 0

    engine = LocalEngine(args)
    engine.load()

    total = len(pending)
    failures: List[Tuple[Path, str]] = []
    started_all = time.perf_counter()

    for index, (pdf_path, base) in enumerate(pending, start=1):
        prefix = f"[{index}/{total}]"
        print(f"{prefix} 正在处理：{pdf_path}", flush=True)
        started = time.perf_counter()
        try:
            written = await process_one(
                engine, pdf_path, base, args.formats, args, prefix
            )
            elapsed = time.perf_counter() - started
            print(
                f"{prefix} 输出（{elapsed:.1f}s）："
                + "，".join(str(path) for path in written),
                flush=True,
            )
        except KeyboardInterrupt:
            raise
        except Exception as error:  # noqa: BLE001 - 单个文件失败不应中断批处理
            elapsed = time.perf_counter() - started
            LOGGER.error("处理失败（%.1fs）：%s", elapsed, error, exc_info=args.verbose)
            failures.append((pdf_path, f"{type(error).__name__}: {error}"))

    total_elapsed = time.perf_counter() - started_all
    print(
        f"\n全部完成：成功 {total - len(failures)}/{total}"
        + (f"，跳过 {skipped}" if skipped else "")
        + f"，总耗时 {total_elapsed:.1f}s",
        flush=True,
    )
    summary = engine.stats.summary()
    if summary:
        print("\n" + summary, flush=True)
    if failures:
        print("失败文件：", flush=True)
        for path, message in failures:
            print(f"  - {path}: {message}", flush=True)
        return 1
    return 0


def main(argv: Optional[Sequence[str]] = None) -> int:
    _setup_stdio()
    _bootstrap_sys_path()
    parser = build_parser()
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )
    logging.getLogger("transformers").setLevel(logging.ERROR)
    logging.getLogger("urllib3").setLevel(logging.ERROR)
    logging.getLogger("pikepdf").setLevel(logging.WARNING)
    try:
        from transformers.utils import logging as hf_logging

        hf_logging.set_verbosity_error()
        hf_logging.disable_progress_bar()
    except Exception:  # noqa: BLE001 - 旧版本 transformers 可能没有该模块
        pass

    try:
        return asyncio.run(run(args))
    except KeyboardInterrupt:
        print("\n已取消。", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
