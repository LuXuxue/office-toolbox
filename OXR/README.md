# OXR（本地推理子集）

从 [SoMarkAI/OXR](https://github.com/SoMarkAI/OXR) 复制的最小运行子集，供上级目录的 `pdf2md.py` 使用。许可证见 `LICENSE` 与 `NOTICE`。

- `oxr/` — 官方流水线：版面结果 → 阅读顺序 → 分类识别 → md/json。已去掉 `cli/` 与 `server/`，其余保持原样。
- `services/layout_analyze/` — D-FINE 版面分析推理代码，以及第三方许可声明（`THIRD_PARTY_NOTICES.md`、`third_party/D-FINE/LICENSE`）。
- `models/OXR-1.0/` — 模型权重，不入库（见上级 `.gitignore`），需自行下载：

```bash
pip install modelscope
modelscope download SoMark/OXR-1.0 --local_dir ./OXR/models/OXR-1.0
```

升级上游时整块覆盖 `oxr/` 与 `services/layout_analyze/dfine_la/`，然后重跑 `python pdf2md.py` 的单文件用例。本地改动的接线（替换 `model_client` / `layout_client`、生成参数）都在 `pdf2md.py` 里，不要写进这个目录。
