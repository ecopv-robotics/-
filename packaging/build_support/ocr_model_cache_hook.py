"""Use the bundled OCR model cache without importing ModelScope's model zoo.

PaddleX imports ModelScope eagerly, even though it only calls it to download
models. The operator release ships all five OCR models, so the download API
must never be needed during normal use.
"""

import sys
import types


modelscope = types.ModuleType("modelscope")


def snapshot_download(*_args, **_kwargs):
    raise RuntimeError("Required OCR model is missing from the bundled model_cache")


modelscope.snapshot_download = snapshot_download
sys.modules["modelscope"] = modelscope
