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

结果默认写到源 PDF 所在目录，文件名与源文件相同，扩展名为 .md / .json；
PDF 里的图片导出到同目录的 <文件名>_media/。

速度相关的几个旋钮（默认值按“单卡、可中断”调过）：

    --batch-size      文本模型的批大小。流水线原本面向 vLLM 的连续批处理，
                      本地串行时把同一类型的多个识别请求拼成一批，能把
                      每 token 都要重读一遍模型权重的开销摊薄，4~8 收益明显。
    --chunk-pages     每批送进流水线的页数。越小越省显存、进度越细，
                      但单批的固定开销占比越高。
    --pages           页码范围（1 起），只在选中的页上工作。
    --layout-concurrency  版面分析的在途批数（单卡串行，一般不用改）。
    --graph-decode    解码段是否捕获成 CUDA Graph（auto/on/off，仅 NVIDIA）。
    --graph-cache-len 图解码预留的缓存长度，超出的请求回退到普通生成。

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
        description="使用 OXR-1.0 把 PDF 转成 Markdown，输出同名 .md / .json 文件。",
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
        "--formats", type=_parse_formats, default=["markdown", "json"],
        help="输出格式，逗号分隔，可选 md、json（默认：md,json）",
    )
    parser.add_argument(
        "--pages", type=_parse_pages, default=None,
        help="只处理指定页（1 起），如 1-5、1-3,8、12-（默认：全部）",
    )
    parser.add_argument(
        "--chunk-pages", type=int, default=8,
        help="每批送进流水线的页数（默认：8，越小越省内存、进度越细）",
    )
    parser.add_argument(
        "--batch-size", type=int, default=4,
        help="文本模型的批大小，同类型请求合批推理（默认：4，1 表示不合批）",
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

# 攒批时最多等这么久；本地生成是瓶颈，等一小会儿换更高的一批利用率更划算。
GROUP_WAIT_SECONDS = 0.05


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
        while produced < budget and not all(finished):
            steps = min(GRAPH_CHUNK_STEPS, budget - produced)
            with torch.inference_mode():
                for _ in range(steps):
                    self.graph.replay()
                torch.cuda.synchronize(self.device)
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
        self.graph_decoder: Optional[GreedyGraphDecoder] = None
        if graph_decode and _cuda_graph_supported(torch, device):
            try:
                self.graph_decoder = GreedyGraphDecoder(
                    self.model.llm,
                    batch=max(1, int(graph_batch)),
                    cache_len=int(graph_cache_len),
                    device=device,
                )
                LOGGER.info(
                    "CUDA Graph 解码可用（批 %d，缓存 %d token，SDPA %s）",
                    self.graph_decoder.batch,
                    self.graph_decoder.cache_len,
                    "GQA" if self.graph_decoder.use_gqa else "repeat_kv",
                )
            except Exception as error:  # noqa: BLE001 - 退化为普通 generate
                LOGGER.warning("CUDA Graph 解码不可用，回退到 generate：%s", error)
                self.graph_decoder = None
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

    def generate_batch(
        self,
        prepared: Sequence[Any],
        max_new_tokens: int,
        model_options: Optional[Dict[str, Any]] = None,
        limits: Optional[Sequence[int]] = None,
    ):
        """把若干条预处理结果拼成一个 batch 一次性生成。

        左填充到等长，靠 attention_mask 屏蔽填充位；视觉 token 用负数的
        指示 ID 占位，merge_multimodal 会按行序填回视觉特征，所以填充用的
        pad_token_id（正数）不会和它们冲突。
        """
        from oxr.model.client import ModelOutput

        torch = self.torch
        if not prepared:
            return []

        started = time.perf_counter()
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
        batch_input_ids = torch.stack(rows).to(self.device)

        attention_mask = torch.zeros_like(batch_input_ids, dtype=torch.bool)
        for row, length in enumerate(lengths):
            attention_mask[row, max_length - length:] = True

        pixel_values = None
        if prepared[0][1] is not None:
            pixel_values = torch.cat([item[1] for item in prepared], dim=0).to(
                self.device, dtype=self.dtype
            )
        grid_thws = None
        if prepared[0][2] is not None:
            grid_thws = torch.cat([item[2] for item in prepared], dim=0).to(self.device)

        generation_kwargs = self._generation_kwargs(max_new_tokens, model_options)
        limit = int(generation_kwargs["max_new_tokens"])
        row_limits = list(limits or [limit] * len(prepared))
        greedy = (
            generation_kwargs.get("do_sample") is False
            and "repetition_penalty" not in generation_kwargs
        )
        decoder = self.graph_decoder
        graph_used = False
        tokens_by_row = None

        with torch.inference_mode():
            inputs_embeds = self.model.merge_multimodal(
                input_ids=batch_input_ids,
                pixel_values=pixel_values,
                grid_thws=grid_thws,
            )

        if (
            decoder is not None
            and greedy
            and decoder.supports(max_length)
        ):
            try:
                tokens_by_row, complete = decoder.decode(
                    inputs_embeds, attention_mask, max(row_limits), self._eos_lookup
                )
                graph_used = complete
            except Exception as error:  # noqa: BLE001 - 图解码失败就退回 generate
                LOGGER.warning("CUDA Graph 解码失败，回退到 generate：%s", error)
                tokens_by_row = None
            if not graph_used:
                tokens_by_row = None
                if self.stats is not None:
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
            cut = len(tokens)
            for position, token in enumerate(tokens):
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
                len(prepared),
                produced_tokens,
                time.perf_counter() - started,
                graph_used,
            )
        return results


class _PendingRequest:
    """排队中的一次视觉问答请求。key 相同的请求才会被合进同一批。"""

    __slots__ = ("prompt", "image", "limit", "options", "key", "future")

    def __init__(
        self, prompt: str, image, limit: int, options: Dict[str, float]
    ) -> None:
        self.prompt = prompt
        self.image = image
        self.limit = limit
        self.options = options
        # 提示词相同、token 上限相同、采样参数相同 -> 输出长度分布接近，适合合批
        self.key = (prompt, limit, tuple(sorted(options.items())))
        self.future: Optional[asyncio.Future] = None


class LocalOXRClient:
    """替换 OXRModelClient：把本机模型调用聚合成批，单卡串行执行。

    流水线本身是按“HTTP 服务端有连续批处理”设计的，每个 block 发一次请求。
    本地推理如果照单全收就退化成几千次 batch=1 的生成，而 batch=1 时每生成
    一个 token 都要把整份权重读一遍显存。这里用两个常驻协程把请求攒起来：

        _prepare_loop  : 按 key 取一组请求 -> 线程池里做 CPU 预处理
        _generate_loop : 取预处理结果 -> 线程里做一次批量 GPU 生成

    预处理队列深度为 2，所以第 N 批在 GPU 上跑的时候，第 N+1 批已经在
    CPU 上准备好了。
    """

    def __init__(self, backend: OXRTextBackend, config, batch_size: int = 4) -> None:
        self.backend = backend
        self.config = config
        self.batch_size = max(1, int(batch_size))
        self._pending: List[_PendingRequest] = []
        self._condition = asyncio.Condition()
        self._prepared: Optional[asyncio.Queue] = None
        self._workers: List[asyncio.Task] = []
        self._prepare_pool = ThreadPoolExecutor(
            max_workers=1, thread_name_prefix="oxr-prepare"
        )
        self._semaphore: Optional[asyncio.Semaphore] = None
        self._semaphore_loop: Optional[asyncio.AbstractEventLoop] = None

    def get_request_semaphore(self) -> asyncio.Semaphore:
        """流水线用它限制在途请求数。

        本地实现里真正的串行点是对 GPU 的独占，用 1 会让请求永远无法攒成
        一批；这里放宽到批大小的两倍，只作为内存/背压上限。
        """
        loop = asyncio.get_running_loop()
        limit = max(self.batch_size * 2, 4)
        if self._semaphore_loop is not loop or self._semaphore is None:
            self._semaphore = asyncio.Semaphore(limit)
            self._semaphore_loop = loop
        return self._semaphore

    # -- 常驻协程 ---------------------------------------------------------- #

    def _ensure_workers(self) -> None:
        if self._workers:
            return
        self._prepared = asyncio.Queue(maxsize=2)
        loop = asyncio.get_running_loop()
        self._workers.append(loop.create_task(self._prepare_loop()))
        self._workers.append(loop.create_task(self._generate_loop()))

    async def _next_group(self) -> List[_PendingRequest]:
        """取出队首请求所在的一批同类型请求。

        生成阶段是瓶颈、而且批大小几乎不影响单步耗时（CUDA Graph 下批 1 与
        批 4 只差几毫秒），所以这里愿意多等一小会儿把一批攒满：36 个请求
        分成 24 批和分成 9 批，解码时间差一倍以上。
        """
        async with self._condition:
            while not self._pending:
                await self._condition.wait()
            first = self._pending[0]
            deadline = time.monotonic() + GROUP_WAIT_SECONDS
            while True:
                same = [r for r in self._pending if r.key == first.key]
                if len(same) >= self.batch_size:
                    break
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    break
                try:
                    await asyncio.wait_for(self._condition.wait(), timeout=remaining)
                except asyncio.TimeoutError:
                    break
            group: List[_PendingRequest] = []
            rest: List[_PendingRequest] = []
            for request in self._pending:
                if request.key == first.key and len(group) < self.batch_size:
                    group.append(request)
                else:
                    rest.append(request)
            self._pending = rest
            return group

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

    async def _generate_loop(self) -> None:
        while True:
            try:
                group, prepared = await self._prepared.get()
                limits = [request.limit for request in group]
                options = group[0].options
                try:
                    outputs = await asyncio.to_thread(
                        self.backend.generate_batch,
                        prepared,
                        max(limits),
                        options,
                        limits,
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
            future.set_result(value)

    @staticmethod
    def _reject(request: _PendingRequest, error: BaseException) -> None:
        future = request.future
        if future is not None and not future.done():
            future.set_exception(error)

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
        for request in self._pending:
            self._reject(request, RuntimeError("客户端已关闭"))
        self._pending.clear()
        self._prepare_pool.shutdown(wait=False)

    async def _call_vision_model(
        self,
        prompt: str,
        image,
        timeout: int,
        model_options: Optional[Dict[str, Any]] = None,
        max_tokens: Optional[int] = None,
    ):
        from oxr.model.client import resolve_model_options

        del timeout  # 本地推理没有网络超时
        if image is None:
            raise ValueError("An image is required for vision model calls.")

        self._ensure_workers()
        options = resolve_model_options(model_options)
        limit = int(max_tokens or self.config.max_tokens or 8192)
        request = _PendingRequest(prompt, image, limit, dict(options))
        request.future = asyncio.get_running_loop().create_future()

        async with self._condition:
            self._pending.append(request)
            self._condition.notify_all()
        return await request.future

    # 以下方法复用 oxr.model.client.OXRModelClient 的实现（prompt 与参数一致），
    # 通过继承在 _install 中绑定即可。
    async def request_reading_order(self, prompt, image, max_tokens):
        return await self._call_vision_model(
            prompt=prompt, image=image, timeout=self.config.timeout,
            model_options=None, max_tokens=max_tokens,
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
            import torch  # noqa: F401
            import transformers  # noqa: F401
            from oxr.config.settings import settings
            from oxr.pipeline.orchestrator import process_decoded_document
            from oxr.utils.pdf import pdf_to_images
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

        self.settings = settings
        self.pdf_to_images = pdf_to_images
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


async def process_one(
    engine: LocalEngine,
    pdf_path: Path,
    base: Path,
    formats: Sequence[str],
    args: argparse.Namespace,
    prefix: str = "",
) -> List[Path]:
    started = time.perf_counter()
    pdf_bytes = pdf_path.read_bytes()
    images = await asyncio.to_thread(engine.pdf_to_images, pdf_bytes, args.dpi)
    render_elapsed = time.perf_counter() - started
    if not images:
        raise ValueError("PDF 没有可渲染的页面")

    total_pages = len(images)
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
    markdown_parts: List[str] = []
    json_pages: List[Dict[str, Any]] = []
    totals = {"layout": 0.0, "recognition": 0.0, "assembly": 0.0}
    block_total = 0
    written: List[Path] = []

    for chunk_no, offset in enumerate(range(0, len(selected), chunk_size), start=1):
        chunk = selected[offset:offset + chunk_size]
        chunk_started = time.perf_counter()
        result = await engine.process_decoded_document(
            [images[index] for index in chunk],
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
            markdown_parts.append(outputs["markdown"])
            outputs["markdown"] = "\n\n".join(part for part in markdown_parts if part)

        # 增量落盘：长文档跑到一半也能先看结果。
        written = write_outputs(base, outputs, formats)

        for index in chunk:
            images[index] = None  # 尽早释放已处理页的位图

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
        f"渲染 {render_elapsed:.1f}s，块 {block_total}，"
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
