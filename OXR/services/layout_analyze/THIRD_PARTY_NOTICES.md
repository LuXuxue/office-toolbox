# Third-party notices

This service contains a minimal inference-only adaptation of
[D-FINE](https://github.com/Peterande/D-FINE).

D-FINE is distributed under the Apache License 2.0. A copy of its license is
included at `third_party/D-FINE/LICENSE`.

The files under `dfine_la/model/` were adapted for a fixed Stage-2 B2 layout
analysis architecture. Changes include packaging, fixed inference
configuration, FP16 inference integration, CUDA Graph-compatible
postprocessing, and removal of training-only entry points.

The model checkpoint is not included. Its distribution terms must be provided
separately by the publisher of the model artifact.
