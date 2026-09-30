"""运行时路径：源码运行和 PyInstaller 便携版共用。"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from typing import Optional


def app_root() -> Path:
    """返回用户数据根目录，而不是 PyInstaller 的临时/内部目录。"""
    configured = os.environ.get("ECOPV_APP_ROOT", "").strip()
    if configured:
        return Path(configured).expanduser().resolve()
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


APP_ROOT = app_root()


def _candidate_paths(value: object, fallback: Optional[object] = None) -> list[Path]:
    """生成跨电脑可用的路径候选。

    旧版会把 ``C:\\Users\\某人\\Desktop\\...`` 写入会话文件。换电脑后该路径
    必然失效；这里按“当前程序目录优先、旧路径文件名兜底”的顺序查找，避免
    启动后继续引用开发机路径。
    """
    raw = str(value or "").strip()
    fallback_raw = str(fallback or "").strip()
    candidates: list[Path] = []
    if raw:
        original = Path(raw).expanduser()
        if original.is_absolute():
            candidates.append(original)
            name = original.name
            if name:
                candidates.extend([
                    APP_ROOT / "data" / name,
                    APP_ROOT / "storage" / "imported" / name,
                    APP_ROOT / "output" / name,
                    APP_ROOT / name,
                ])
        else:
            candidates.append(APP_ROOT / original)
            candidates.append(Path.cwd() / original)
    if fallback_raw:
        fallback_path = Path(fallback_raw).expanduser()
        if fallback_path.is_absolute():
            candidates.append(fallback_path)
        else:
            candidates.append(APP_ROOT / fallback_path)
    result: list[Path] = []
    seen: set[str] = set()
    for path in candidates:
        key = os.path.normcase(os.path.abspath(str(path)))
        if key not in seen:
            seen.add(key)
            result.append(path)
    return result


def resolve_runtime_path(value: object, fallback: Optional[object] = None) -> Path:
    """把配置/会话里的路径解析到当前应用根目录。

    存在的外部文件仍可用于当前次运行；外部文件不存在时，会自动尝试当前
    包内的同名文件，最后落到当前包的 fallback，不再返回旧电脑路径。
    """
    candidates = _candidate_paths(value, fallback)
    for path in candidates:
        try:
            if path.exists():
                return path.resolve()
        except OSError:
            continue
    if candidates:
        target = candidates[-1]
        if not target.is_absolute():
            target = APP_ROOT / target
        return target.resolve()
    return APP_ROOT.resolve()


def app_relative_path(value: object, fallback: Optional[object] = None) -> str:
    """将路径保存为相对应用根目录的路径，供 session/config 跨电脑使用。"""
    path = resolve_runtime_path(value, fallback)
    try:
        relative = path.relative_to(APP_ROOT.resolve())
        return relative.as_posix()
    except ValueError:
        return str(path)


def copy_reference_file(source: object, target_name: str) -> Path:
    """把用户选中的参考表复制到包内 data，避免保存开发机绝对路径。"""
    source_path = Path(str(source)).expanduser().resolve()
    if not source_path.is_file():
        raise FileNotFoundError(source_path)
    safe_name = Path(str(target_name or source_path.name)).name
    target = APP_ROOT / "data" / safe_name
    target.parent.mkdir(parents=True, exist_ok=True)
    if source_path != target:
        shutil.copy2(source_path, target)
    return target.resolve()


def copy_imported_file(source: object, folder: str = "storage/imported") -> Path:
    """复制外部阶段二输入到当前包，供换电脑后继续使用。"""
    source_path = Path(str(source)).expanduser().resolve()
    if not source_path.is_file():
        raise FileNotFoundError(source_path)
    target = APP_ROOT / folder / source_path.name
    target.parent.mkdir(parents=True, exist_ok=True)
    if source_path != target:
        shutil.copy2(source_path, target)
    return target.resolve()
