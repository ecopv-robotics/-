"""Slim onedir variant: same app, without dependencies the OCR stack never loads.

Baseline measurement (2026-10-15):
  * `import paddleocr` pulls scipy / matplotlib / shapely / langchain / cv2 /
    pandas / numpy / PIL into sys.modules -> those MUST stay.
  * polars / numba / llvmlite / jieba / pymupdf(fitz) / asyncmy / grpc / hf_xet
    are NOT loaded by paddleocr and are not referenced anywhere in the project
    source -> ~377 MB removable.

Built to its own directory so the 2026-10-10-r2 release stays untouched and the
two can be compared side by side.
"""

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, copy_metadata


PROJECT_ROOT = Path(SPECPATH).resolve().parent
PACKAGE_NAME = "ECOPV_Mail_Audit_Workbench_2026-10-15-slim"

datas = [(str(PROJECT_ROOT / "workbench.html"), ".")]
web_assets = PROJECT_ROOT / "web_console"
for asset in web_assets.rglob("*"):
    if asset.is_file():
        relative_parent = asset.relative_to(web_assets).parent
        destination = Path("web_console") / relative_parent
        datas.append((str(asset), destination.as_posix()))

binaries = []
hiddenimports = []
# PaddleOCR's public package eagerly imports its OCR classes, and PyInstaller's
# normal analysis follows those imports. Collecting every PaddleOCR submodule
# also pulls optional frameworks (for example TensorFlow) that this app never
# calls. Paddle's Python imports are followed by analysis. Include its native
# runtime libraries without sweeping optional modules into this release.
binaries.extend(collect_dynamic_libs("paddle"))

# The app uses only PaddleX's standard OCR pipeline. Collect its runtime config,
# while leaving unrelated recognition pipelines and optional serving backends out.
datas.extend(
    collect_data_files(
        "paddlex",
        includes=[
            "configs/pipelines/OCR.yaml",
            "configs/pipelines/doc_preprocessor.yaml",
        ],
    )
)

# The OCR engine resolves package versions at runtime.
for distribution in ("paddleocr", "paddlex", "paddlepaddle"):
    datas.extend(copy_metadata(distribution))

analysis = Analysis(
    [str(PROJECT_ROOT / "web_app.py")],
    pathex=[str(PROJECT_ROOT)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[str(PROJECT_ROOT / "packaging" / "pyinstaller_hooks")],
    hooksconfig={},
    runtime_hooks=[str(PROJECT_ROOT / "packaging" / "build_support" / "ocr_model_cache_hook.py")],
    excludes=[
        "pytest", "IPython", "jupyter", "notebook",
        "torch", "torch.*", "torchvision", "torchvision.*",
        "tensorflow", "tensorflow.*",
        "modelscope", "modelscope.*",
        "onnxruntime", "onnxruntime.*",
        "paddlex.inference.genai", "paddlex.inference.genai.*",
        "paddlex.inference.serving", "paddlex.inference.serving.*",
        # --- slim-only: not loaded by paddleocr, not referenced by project code ---
        "polars", "polars.*",
        "numba", "numba.*",
        "llvmlite", "llvmlite.*",
        "jieba", "jieba.*",
        "fitz", "fitz.*", "pymupdf", "pymupdf.*",
        "asyncmy", "asyncmy.*",
        "grpc", "grpc.*",
        "hf_xet", "hf_xet.*",
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(analysis.pure)

exe = EXE(
    pyz,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name="ECOPV_Mail_Audit_Workbench",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
)

release = COLLECT(
    exe,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=False,
    name=PACKAGE_NAME,
)
