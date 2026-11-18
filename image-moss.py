# -*- coding: utf-8 -*-
"""image-moss.py — 调用本地 ComfyUI 的 MOSS 模型出图。

不启动 ComfyUI 服务，也不依赖任何自定义节点：把 ComfyUI 目录加进 sys.path，
直接使用 ComfyUI 自带节点类，按下面的接线依次执行：

    UNETLoader -> CLIPLoader / VAELoader
    CLIPTextEncode(正向/负向)
    SkipLayerGuidanceDiTSimple
    EmptyFlux2LatentImage -> KSampler -> VAEDecode -> 存 PNG

默认参数：seed=1、steps=6、cfg=1、er_sde/simple、1536x1024，
moss_image_sft_4step_v2.0_int8_convrot + Qwen3.5-9B-heretic-v2 + flux2-vae_small。

示例：
    python image-moss.py "a cat sitting on a windowsill, soft light"
    python image-moss.py "..." -o out --name cat --seed 42 --steps 8
    python image-moss.py "..." --width 1024 --height 1536 --cfg 1.5
"""

import argparse
import datetime
import os
import sys

from PIL import Image
from PIL.PngImagePlugin import PngInfo

# ComfyUI 源码目录，可用环境变量 COMFYUI_PATH 或 --comfyui 覆盖。
DEFAULT_COMFYUI = os.getenv("COMFYUI_PATH", r"E:\ComfyUI")

# 以下默认值：模型名沿用 folder_paths 的相对路径，
# 和 ComfyUI 里下拉框看到的写法一致。
DEFAULT_UNET = r"MOSS\moss_image_sft_4step_v2.0_int8_convrot.safetensors"
DEFAULT_CLIP = "Qwen3.5-9B-heretic-v2_int8_convrot.safetensors"
DEFAULT_VAE = "flux2-vae_small.safetensors"
DEFAULT_STEPS = 6
DEFAULT_CFG = 1.0
DEFAULT_SEED = 1
DEFAULT_SAMPLER = "er_sde"
DEFAULT_SCHEDULER = "simple"
# 1536x1024 是「3:2 + 1.5 MP + 16 的倍数」得到的尺寸。
DEFAULT_WIDTH = 1536
DEFAULT_HEIGHT = 1024
# SkipLayerGuidanceDiTSimple 的默认值，cfg=1 时不走 uncond，因此不起作用。
DEFAULT_SLG_SINGLE = "22, 24, 26"
DEFAULT_SLG_DOUBLE = ""
DEFAULT_SLG_START = 0.0
DEFAULT_SLG_END = 0.5


def round_to_multiple(value, multiple=16):
    """取最近的 multiple 的整数倍（四舍五入，至少一个 multiple）。"""
    return max(multiple, int(value / multiple + 0.5) * multiple)


def build_parser():
    parser = argparse.ArgumentParser(
        description="用本地 ComfyUI 的 MOSS 模型生成图像（正向提示词必填）。",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("prompt", help="正向提示词（必填）")
    parser.add_argument("-n", "--negative", default="",
                        help="负向提示词，空字符串表示不指定")
    parser.add_argument("--width", type=int, default=DEFAULT_WIDTH,
                        help="图像宽度，按 16 取最近值")
    parser.add_argument("--height", type=int, default=DEFAULT_HEIGHT,
                        help="图像高度，按 16 取最近值")
    parser.add_argument("--seed", type=int, default=DEFAULT_SEED, help="随机种子")
    parser.add_argument("--steps", type=int, default=DEFAULT_STEPS, help="采样步数")
    parser.add_argument("--cfg", type=float, default=DEFAULT_CFG,
                        help="CFG 强度")
    parser.add_argument("--sampler-name", default=DEFAULT_SAMPLER,
                        help="采样器名称")
    parser.add_argument("--scheduler", default=DEFAULT_SCHEDULER,
                        help="调度器名称")

    parser.add_argument("--unet-name", default=DEFAULT_UNET,
                        help="扩散模型（models/diffusion_models 下的相对路径）")
    parser.add_argument("--clip-name", default=DEFAULT_CLIP,
                        help="文本编码器（models/text_encoders 下的文件名）")
    parser.add_argument("--vae-name", default=DEFAULT_VAE,
                        help="VAE（models/vae 下的文件名）")
    parser.add_argument("--weight-dtype", default="default",
                        help="UNETLoader 的 weight_dtype")
    parser.add_argument("--clip-type", default="moss_image",
                        help="CLIPLoader 的 type")

    parser.add_argument("--slg-single-layers", default=DEFAULT_SLG_SINGLE,
                        help="SkipLayerGuidance 的 single_layers，空字符串表示跳过")
    parser.add_argument("--slg-double-layers", default=DEFAULT_SLG_DOUBLE,
                        help="SkipLayerGuidance 的 double_layers，空字符串表示跳过")
    parser.add_argument("--slg-start", type=float, default=DEFAULT_SLG_START,
                        help="SkipLayerGuidance 的起始百分比")
    parser.add_argument("--slg-end", type=float, default=DEFAULT_SLG_END,
                        help="SkipLayerGuidance 的结束百分比")

    parser.add_argument("--no-ck-attention", dest="ck_attention",
                        action="store_false",
                        help="关闭 ComfyUI 的 --use-ck-attention（默认开启）")
    parser.add_argument("--no-fast-disk", dest="fast_disk",
                        action="store_false",
                        help="关闭 ComfyUI 的 --fast-disk（默认开启）")

    parser.add_argument("-o", "--output-dir", default=None,
                        help="输出目录，默认是执行命令时的当前目录")
    parser.add_argument("--name", default=None,
                        help="输出文件名（可带 .png），默认用日期时间，精确到秒")
    parser.add_argument("--comfyui", default=DEFAULT_COMFYUI,
                        help="ComfyUI 源码目录")
    return parser


def check_comfyui(path):
    """确认给定目录是 ComfyUI 源码目录（要有 nodes.py）。"""
    path = os.path.abspath(os.path.expanduser(path))
    if not os.path.isfile(os.path.join(path, "nodes.py")):
        raise SystemExit(f"找不到 ComfyUI：{path}（该目录下应有 nodes.py）")
    return path


def load_comfyui(comfyui, ck_attention=True, fast_disk=True):
    """把 ComfyUI 加进 sys.path 并导入需要的原生节点。

    必须先把路径插进 sys.path 再 import torch/comfy，否则拿不到同一套环境。
    ComfyUI 的 `--use-ck-attention` / `--fast-disk` 原本由 main.py 解析命令行
    写入 comfy.cli_args.args；本脚本不经过 main.py，所以在这里自己写。
    attention 后端是在 import comfy.ldm.modules.attention 时绑定的，
    因此这几行必须放在 import nodes 之前，否则设置不会生效。
    """
    if comfyui not in sys.path:
        sys.path.insert(0, comfyui)

    import comfy.cli_args
    comfy.cli_args.args.use_ck_attention = ck_attention
    comfy.cli_args.args.fast_disk = fast_disk

    import torch
    import comfy.samplers
    import comfy.ldm.modules.attention as attention
    import nodes
    from comfy_extras.nodes_flux import EmptyFlux2LatentImage
    from comfy_extras.nodes_slg import SkipLayerGuidanceDiTSimple

    return (torch, comfy, nodes, EmptyFlux2LatentImage,
            SkipLayerGuidanceDiTSimple, attention.optimized_attention.__name__)


def output_path(args):
    """算出输出文件路径：默认当前目录 + 年月日-时分秒.png。"""
    out_dir = os.path.abspath(args.output_dir or os.getcwd())
    os.makedirs(out_dir, exist_ok=True)
    name = args.name or datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    if not name.lower().endswith(".png"):
        name += ".png"
    return os.path.join(out_dir, name)


def main(argv=None):
    args = build_parser().parse_args(argv)
    comfyui = check_comfyui(args.comfyui)

    width = round_to_multiple(args.width)
    height = round_to_multiple(args.height)

    torch, comfy, nodes, EmptyFlux2LatentImage, SkipLayerGuidanceDiTSimple, attn = \
        load_comfyui(comfyui, args.ck_attention, args.fast_disk)

    # 提前校验采样器/调度器名，避免加载完模型才发现拼错。
    if args.sampler_name not in comfy.samplers.KSampler.SAMPLERS:
        raise SystemExit(f"未知采样器：{args.sampler_name}")
    if args.scheduler not in comfy.samplers.KSampler.SCHEDULERS:
        raise SystemExit(f"未知调度器：{args.scheduler}")

    print(f"ComfyUI: {comfyui}")
    print(f"加速: ck-attention={attn} fast-disk={args.fast_disk}")
    print(f"模型: {args.unet_name} | {args.clip_name} | {args.vae_name}")
    print(f"参数: {width}x{height} seed={args.seed} steps={args.steps} "
          f"cfg={args.cfg} {args.sampler_name}/{args.scheduler}")

    # 1. 加载三个模型（对应 ComfyUI 的三个 Loader 节点）。
    print("[1/4] 加载模型 ...", flush=True)
    model = nodes.UNETLoader().load_unet(
        unet_name=args.unet_name, weight_dtype=args.weight_dtype)[0]
    clip = nodes.CLIPLoader().load_clip(
        clip_name=args.clip_name, type=args.clip_type, device="default")[0]
    vae = nodes.VAELoader().load_vae(vae_name=args.vae_name)[0]

    # 2. 编码提示词 + 挂 SkipLayerGuidance。
    print("[2/4] 编码提示词 ...", flush=True)
    positive = nodes.CLIPTextEncode().encode(clip=clip, text=args.prompt)[0]
    negative = nodes.CLIPTextEncode().encode(clip=clip, text=args.negative)[0]
    model = SkipLayerGuidanceDiTSimple.execute(
        model=model,
        double_layers=args.slg_double_layers,
        single_layers=args.slg_single_layers,
        start_percent=args.slg_start,
        end_percent=args.slg_end,
    )[0]

    # 3. 空 latent -> 采样 -> VAE 解码。
    latent = EmptyFlux2LatentImage.execute(width=width, height=height,
                                           batch_size=1)[0]
    print("[3/4] 采样并解码 ...", flush=True)
    with torch.inference_mode():
        samples = nodes.KSampler().sample(
            model=model, seed=args.seed, steps=args.steps, cfg=args.cfg,
            sampler_name=args.sampler_name, scheduler=args.scheduler,
            positive=positive, negative=negative, latent_image=latent,
            denoise=1.0,
        )[0]
        images = nodes.VAEDecode().decode(vae=vae, samples=samples)[0]

    # 4. 存图。IMAGE 是 [B,H,W,C]、浮点 0~1，PIL 只吃 3 维 uint8。
    path = output_path(args)
    print(f"[4/4] 保存到 {path}", flush=True)
    array = images[0].detach().float().cpu().clamp(0, 1).mul(255).round()
    array = array.to(torch.uint8).numpy()

    meta = PngInfo()
    meta.add_text("prompt", args.prompt)
    meta.add_text("negative", args.negative)
    meta.add_text("model", args.unet_name)
    meta.add_text("seed", str(args.seed))
    meta.add_text("steps", str(args.steps))
    meta.add_text("cfg", str(args.cfg))
    meta.add_text("sampler", f"{args.sampler_name}/{args.scheduler}")
    meta.add_text("resolution", f"{width}x{height}")
    Image.fromarray(array).save(path, pnginfo=meta)

    # 主动释放模型，避免解释器退出时 ModelPatcher.__del__ 在模块已清空后报错。
    del model, clip, vae
    comfy.model_management.unload_all_models()

    print(f"完成: {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
