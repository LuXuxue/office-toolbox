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

结果默认写到源 PDF 所在目录，文件名与源文件相同，扩展名为 .md / .json；
PDF 里的图片导出到同目录的 <文件名>_media/。

依赖（GPU 版 torch 请按 https://pytorch.org 的说明安装）：

    pip install -r requirements.txt
"""

from __future__ import annotations

import argparse
import asyncio
import importlib
import json
import logging
import sys
import time
import types
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


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="使用 OXR-1.0 把 PDF 转成 Markdown，输出同名 .md / .json 文件。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python pdf2md.py a.pdf\n"
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


class OXRTextBackend:
    """用 transformers 本地加载 OXR-1.0 并执行视觉问答。"""

    def __init__(self, model_dir: Path, device) -> None:
        import torch
        from transformers import GenerationConfig

        self.torch = torch
        self.device = device
        self.dtype = torch.bfloat16 if device.type == "cuda" else torch.float32

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
        LOGGER.info("模型加载完成，耗时 %.1fs", time.perf_counter() - started)

    def generate(
        self,
        prompt: str,
        image_bgr,
        max_new_tokens: Optional[int],
        model_options: Optional[Dict[str, Any]] = None,
    ):
        import cv2
        from PIL import Image

        from oxr.model.client import ModelOutput, resolve_model_options

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
        input_ids = input_ids.to(self.device)
        if pixel_values is not None:
            pixel_values = pixel_values.to(self.device, dtype=self.dtype)
        if grid_thws is not None:
            grid_thws = grid_thws.to(self.device)

        options = resolve_model_options(model_options)
        limit = int(max_new_tokens or 8192)
        generation_kwargs: Dict[str, Any] = {
            "max_new_tokens": limit,
            "use_cache": True,
            "pad_token_id": self.tokenizer.pad_token_id,
        }
        eos_token_id = (
            self.generation_config.eos_token_id
            if self.generation_config is not None
            else None
        )
        if eos_token_id is None:
            eos_token_id = self.tokenizer.eos_token_id
        if eos_token_id is not None:
            generation_kwargs["eos_token_id"] = eos_token_id
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

        with self.torch.inference_mode():
            generated = self.model.generate(
                input_ids,
                pixel_values=pixel_values,
                grid_thws=grid_thws,
                **generation_kwargs,
            )
        generated_tokens = int(generated.shape[-1])
        text = self.tokenizer.decode(generated[0], skip_special_tokens=True)
        finish_reason = "length" if generated_tokens >= limit else "stop"
        return ModelOutput(text, finish_reason=finish_reason)


class LocalOXRClient:
    """替换 OXRModelClient：用本地模型代替 OpenAI 兼容的 HTTP 服务。"""

    def __init__(self, backend: OXRTextBackend, config) -> None:
        self.backend = backend
        self.config = config
        self._lock = asyncio.Lock()

    def get_request_semaphore(self) -> asyncio.Semaphore:
        # 本地单卡推理必须串行执行，见 _call_vision_model 中的锁。
        return asyncio.Semaphore(1)

    async def aclose(self) -> None:
        return None

    async def _call_vision_model(
        self,
        prompt: str,
        image,
        timeout: int,
        model_options: Optional[Dict[str, Any]] = None,
        max_tokens: Optional[int] = None,
    ):
        del timeout  # 本地推理没有网络超时
        if image is None:
            raise ValueError("An image is required for vision model calls.")
        async with self._lock:
            return await asyncio.to_thread(
                self.backend.generate, prompt, image, max_tokens, model_options
            )

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

    def __init__(self, model_path: Path, device_name: str) -> None:
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

    def infer(self, payloads: Sequence[bytes]) -> Dict[str, Any]:
        from dfine_la.confidence import format_predictions

        ready, _decode_ms, _preprocess_ms = self._runtime.prepare(list(payloads))
        outputs, _timings = self._runtime.infer(ready)
        # 阈值过滤和 NMS 由 oxr.pipeline.postprocess.layout_postprocess 统一完成。
        return format_predictions(outputs)


class LocalLayoutClient:
    """替换 LayoutAnalyzeClient：本地调用 D-FINE，而不是 HTTP 服务。"""

    def __init__(self, runtime: LocalLayoutRuntime, config) -> None:
        import cv2

        self._cv2 = cv2
        self.runtime = runtime
        self.config = config
        self._lock = asyncio.Lock()

    def encode_images(self, images: Sequence[Any]) -> List[bytes]:
        encoded_images: List[bytes] = []
        for image in images:
            success, encoded = self._cv2.imencode(".png", image)
            if not success:
                raise ValueError("Unable to encode image as PNG.")
            encoded_images.append(encoded.tobytes())
        return encoded_images

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

        # 先调整流水线参数，再加载模型。
        settings.oxr_model.max_tokens = int(self.args.max_tokens)
        settings.layout_analyze_model.batch_size = int(self.args.layout_batch_size)

        text_backend = OXRTextBackend(model_dir, device)
        oxr_client = LocalOXRClient(text_backend, settings.oxr_model)
        layout_client = LocalLayoutClient(
            LocalLayoutRuntime(layout_model, device.type), settings.layout_analyze_model
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


async def process_one(
    engine: LocalEngine,
    pdf_path: Path,
    base: Path,
    formats: Sequence[str],
    args: argparse.Namespace,
) -> List[Path]:
    pdf_bytes = pdf_path.read_bytes()
    images = await asyncio.to_thread(engine.pdf_to_images, pdf_bytes, args.dpi)
    if not images:
        raise ValueError("PDF 没有可渲染的页面")

    # PDF 里的图片按 pptx2md 的习惯导出到同级的 <文件名>_media/，
    # md/json 里写相对路径，搬走 md 时连同该目录一起搬。
    media_dir = base.parent / f"{base.name}_media"
    # 插图用 uuid 命名，上一次转换留下的文件本次不会再被引用，成功后再清掉。
    stale_media = (
        {path.name for path in media_dir.glob("*.png")} if media_dir.is_dir() else set()
    )
    result = await engine.process_decoded_document(
        images,
        pdf_path.name,
        list(formats),
        keep_header_footer=args.keep_header_footer,
        asset_dir=media_dir,
        asset_url_prefix=f"./{base.name}_media",
    )
    written = write_outputs(base, result["outputs"], formats)
    for name in stale_media:
        (media_dir / name).unlink(missing_ok=True)
    if media_dir.is_dir() and not any(media_dir.iterdir()):
        media_dir.rmdir()
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
            written = await process_one(engine, pdf_path, base, args.formats, args)
            elapsed = time.perf_counter() - started
            print(
                f"{prefix} 完成（{elapsed:.1f}s）："
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
