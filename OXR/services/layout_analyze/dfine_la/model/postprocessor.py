"""
Copied from RT-DETR (https://github.com/lyuwenyu/RT-DETR)
Copyright(c) 2023 lyuwenyu. All Rights Reserved.
"""

import torch
import torch.nn as nn
import torch.nn.functional as F
import torchvision


__all__ = ["DFINEPostProcessor"]


def mod(a, b):
    out = a - a // b * b
    return out


def postprocess_boxes(bbox_pred, orig_target_sizes, target_size=960, letterbox=False):
    if orig_target_sizes.dim() == 1:
        orig_target_sizes = orig_target_sizes.unsqueeze(0)
    orig_target_sizes = orig_target_sizes.to(bbox_pred.dtype)

    if not letterbox:
        scale_x = (orig_target_sizes[:, 0] / target_size).view(-1, 1)
        scale_y = (orig_target_sizes[:, 1] / target_size).view(-1, 1)
        bbox_pred[..., 0] = bbox_pred[..., 0] * scale_x
        bbox_pred[..., 2] = bbox_pred[..., 2] * scale_x
        bbox_pred[..., 1] = bbox_pred[..., 1] * scale_y
        bbox_pred[..., 3] = bbox_pred[..., 3] * scale_y
        return bbox_pred

    orig_w = orig_target_sizes[:, 0]
    orig_h = orig_target_sizes[:, 1]
    max_dim = torch.maximum(orig_w, orig_h)
    scale = target_size / max_dim
    new_w = (orig_w * scale).floor()
    new_h = (orig_h * scale).floor()
    pad_left = ((target_size - new_w) // 2).view(-1, 1)
    pad_top = ((target_size - new_h) // 2).view(-1, 1)
    scale = scale.view(-1, 1)

    bbox_pred[..., 0] = (bbox_pred[..., 0] - pad_left) / scale
    bbox_pred[..., 2] = (bbox_pred[..., 2] - pad_left) / scale
    bbox_pred[..., 1] = (bbox_pred[..., 1] - pad_top) / scale
    bbox_pred[..., 3] = (bbox_pred[..., 3] - pad_top) / scale
    return bbox_pred


class DFINEPostProcessor(nn.Module):
    __share__ = [
        "num_classes",
        "use_focal_loss",
        "num_top_queries",
        "remap_mscoco_category",
        "letterbox",
    ]

    def __init__(
        self,
        num_classes=80,
        use_focal_loss=True,
        num_top_queries=300,
        remap_mscoco_category=False,
        letterbox=False,
        clip_boxes=False,
        target_size=960,
    ) -> None:
        super().__init__()
        self.use_focal_loss = use_focal_loss
        self.num_top_queries = num_top_queries
        self.num_classes = int(num_classes)
        self.remap_mscoco_category = remap_mscoco_category
        self.letterbox = letterbox
        self.clip_boxes = clip_boxes
        self.target_size = int(target_size)
        self.deploy_mode = False

    def extra_repr(self) -> str:
        return (
            f"use_focal_loss={self.use_focal_loss}, num_classes={self.num_classes}, "
            f"num_top_queries={self.num_top_queries}, letterbox={self.letterbox}, "
            f"clip_boxes={self.clip_boxes}, target_size={self.target_size}"
        )

    # def forward(self, outputs, orig_target_sizes):
    def forward(self, outputs, orig_target_sizes: torch.Tensor):
        logits, boxes = outputs["pred_logits"], outputs["pred_boxes"]
        bbox_pred = torchvision.ops.box_convert(boxes, in_fmt="cxcywh", out_fmt="xyxy")
        # Both axes use the same square canvas size. Avoid constructing a CUDA
        # tensor during forward: host-to-device tensor creation invalidates a
        # CUDA Graph stream capture on newer CUDA runtimes.
        bbox_pred = bbox_pred * self.target_size

        if self.use_focal_loss:
            # ------ For ONNX Export ------
            # scores = F.sigmoid(logits).flatten(1)
            # topk_input = F.softmax(logits.flatten(1), dim=-1)
            # offset = torch.arange(topk_input.size()[-1]).to(torch.int32)
            # topk_input = (topk_input * 1000).to(torch.int32) * 10000 + offset
            # _, index = torch.topk(topk_input, self.num_top_queries, dim=-1)
            # scores = scores.gather(dim=-1, index=index)

            # ------ Original ------
            scores = F.sigmoid(logits)
            scores, index = torch.topk(scores.flatten(1), self.num_top_queries, dim=-1)

            # TODO for older tensorrt
            # labels = index % self.num_classes
            labels = mod(index, self.num_classes)
            index = index // self.num_classes
            boxes = bbox_pred.gather(
                dim=1, index=index.unsqueeze(-1).repeat(1, 1, bbox_pred.shape[-1])
            )

        else:
            scores = F.softmax(logits)[:, :, :-1]
            scores, labels = scores.max(dim=-1)
            if scores.shape[1] > self.num_top_queries:
                scores, index = torch.topk(scores, self.num_top_queries, dim=-1)
                labels = torch.gather(labels, dim=1, index=index)
                boxes = torch.gather(
                    boxes, dim=1, index=index.unsqueeze(-1).tile(1, 1, boxes.shape[-1])
                )

        if self.clip_boxes:
            boxes = boxes.clamp(min=0, max=self.target_size - 1)
        boxes = postprocess_boxes(
            boxes,
            orig_target_sizes,
            target_size=self.target_size,
            letterbox=self.letterbox,
        )

        # TODO for onnx export
        if self.deploy_mode:
            return labels, boxes, scores

        results = []
        for lab, box, sco in zip(labels, boxes, scores):
            result = dict(labels=lab, boxes=box, scores=sco)
            results.append(result)

        return results

    def deploy(
        self,
    ):
        self.eval()
        self.deploy_mode = True
        return self
