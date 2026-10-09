"""Hide optional, locally installed ML runtimes from the packaging analysis."""

import os
import sys


if os.environ.get("ECOPV_PACKAGING_BUILD") == "1":
    # PaddleOCR's default OCR path does not use either runtime. ModelScope
    # probes for torch during import, and importing this workstation's Torch
    # build from PyInstaller's DLL scanner can crash the build process.
    sys.modules["torch"] = None
    sys.modules["tensorflow"] = None
