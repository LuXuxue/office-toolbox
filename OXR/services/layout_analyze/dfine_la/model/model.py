"""Fixed D-FINE Stage-2 B2 architecture used by the public LA service."""

from __future__ import annotations

from pathlib import Path

import torch
import torch.nn as nn

from .backbone import HGNetv2
from .decoder import DFINETransformer
from .encoder import HybridEncoder
from .postprocessor import DFINEPostProcessor


class DFINE(nn.Module):
    """D-FINE Stage-2 B2 model for 960px layout analysis."""

    def __init__(self, num_classes: int = 12, eval_spatial_size=(960, 960)):
        super().__init__()
        self.backbone = HGNetv2(
            name="B2",
            use_lab=True,
            return_idx=[1, 2, 3],
            freeze_at=3,
            freeze_stem_only=False,
            freeze_norm=True,
            pretrained=False,
        )
        self.encoder = HybridEncoder(
            in_channels=[384, 768, 1536],
            feat_strides=[8, 16, 32],
            hidden_dim=256,
            use_encoder_idx=[2],
            num_encoder_layers=1,
            nhead=8,
            dim_feedforward=1024,
            dropout=0.0,
            enc_act="gelu",
            expansion=1.0,
            depth_mult=0.67,
            act="silu",
            eval_spatial_size=eval_spatial_size,
        )
        self.decoder = DFINETransformer(
            num_classes=num_classes,
            eval_spatial_size=eval_spatial_size,
            feat_channels=[256, 256, 256],
            feat_strides=[8, 16, 32],
            hidden_dim=256,
            num_levels=3,
            num_layers=4,
            eval_idx=-1,
            num_queries=300,
            num_denoising=100,
            label_noise_ratio=0.5,
            box_noise_scale=1.0,
            reg_max=32,
            reg_scale=4,
            layer_scale=1,
            num_points=[3, 6, 3],
            cross_attn_method="default",
            query_select_method="default",
        )

    def forward(self, images, targets=None):
        features = self.backbone(images)
        features = self.encoder(features)
        return self.decoder(features, targets)

    def deploy(self):
        self.eval()
        for module in self.modules():
            if hasattr(module, "convert_to_deploy"):
                module.convert_to_deploy()
        return self


class DfineWithPostProcessor(nn.Module):
    """Run model and postprocessor under the official FP16 AMP boundary."""

    def __init__(self, model, postprocessor, device: torch.device) -> None:
        super().__init__()
        self.model = model.deploy()
        self.postprocessor = postprocessor.deploy()
        self.device = device

    def forward(self, images, original_sizes):
        with torch.autocast(
            device_type=images.device.type,
            dtype=torch.float16,
            enabled=images.is_cuda,
        ):
            outputs = self.model(images)
            return self.postprocessor(outputs, original_sizes)


def load_model(
    checkpoint_path: str | Path,
    device: torch.device,
    num_classes: int = 12,
    image_size: int = 960,
) -> DfineWithPostProcessor:
    """Load a standard, unencrypted Stage-2 B2 checkpoint."""
    model = DFINE(
        num_classes=num_classes,
        eval_spatial_size=(image_size, image_size),
    )
    checkpoint = torch.load(
        str(checkpoint_path),
        map_location="cpu",
        weights_only=True,
    )
    state = checkpoint["ema"]["module"] if "ema" in checkpoint else checkpoint["model"]
    model.load_state_dict(state)

    postprocessor = DFINEPostProcessor(
        num_classes=num_classes,
        num_top_queries=300,
        letterbox=True,
        clip_boxes=True,
        target_size=image_size,
    )
    wrapped = DfineWithPostProcessor(model, postprocessor, device)
    return wrapped.eval().to(device)
