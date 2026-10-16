"""开放平台凭据的安全存取。

设计约束（来自业务方）：clientKey / clientSecret **不得明文落盘、不得进日志、
不得进仓库、不得经任何接口返回**。

实现：
  * 落盘一律走 Windows DPAPI（CryptProtectData），密文绑定当前 Windows 账户，
    拷到别的机器或同一台机器的别的账户都解不开；
  * 明文只在内存里存活，模块不提供任何"导出明文"的便捷方法；
  * 取值优先级：环境变量（仅用于一次性导入/CI） > DPAPI 密文文件；
  * 对外只暴露 mask()，任何日志/接口要展示凭据都必须先过它。

Windows 专属：非 Windows 平台直接抛错，不做降级（降级就是泄密）。
"""

from __future__ import annotations

import base64
import ctypes
import json
import os
from ctypes import wintypes
from pathlib import Path
from typing import Optional

# --------------------------------------------------------------------------- #
# DPAPI（ctypes 直调；argtypes 必须显式声明，否则 64 位句柄按 32 位传会静默失败）
# --------------------------------------------------------------------------- #

class _DATA_BLOB(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


_crypt32 = None
_kernel32 = None

if os.name == "nt":  # pragma: no cover - 仅在 Windows 生效
    _crypt32 = ctypes.WinDLL("crypt32.dll", use_last_error=True)
    _kernel32 = ctypes.WinDLL("kernel32.dll", use_last_error=True)

    _PROTECT_ARGS = [
        ctypes.POINTER(_DATA_BLOB),   # pDataIn
        wintypes.LPCWSTR,             # szDataDescr
        ctypes.POINTER(_DATA_BLOB),   # pOptionalEntropy
        ctypes.c_void_p,              # pvReserved
        ctypes.c_void_p,              # pPromptStruct
        wintypes.DWORD,               # dwFlags
        ctypes.POINTER(_DATA_BLOB),   # pDataOut
    ]
    _crypt32.CryptProtectData.argtypes = _PROTECT_ARGS
    _crypt32.CryptProtectData.restype = wintypes.BOOL
    _crypt32.CryptUnprotectData.argtypes = _PROTECT_ARGS
    _crypt32.CryptUnprotectData.restype = wintypes.BOOL
    _kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    _kernel32.LocalFree.restype = ctypes.c_void_p


def _dpapi_protect(raw: bytes, description: str = "ecopv-openapi-credential") -> bytes:
    if _crypt32 is None:
        raise RuntimeError("DPAPI 仅在 Windows 上可用")
    inp = ctypes.create_string_buffer(raw, len(raw))
    blob_in = _DATA_BLOB(len(raw), inp)
    blob_out = _DATA_BLOB()
    ok = _crypt32.CryptProtectData(
        ctypes.byref(blob_in), description, None, None, None, 0, ctypes.byref(blob_out)
    )
    if not ok:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        if blob_out.pbData:
            _kernel32.LocalFree(blob_out.pbData)


def _dpapi_unprotect(blob: bytes) -> bytes:
    if _crypt32 is None:
        raise RuntimeError("DPAPI 仅在 Windows 上可用")
    inp = ctypes.create_string_buffer(blob, len(blob))
    blob_in = _DATA_BLOB(len(blob), inp)
    blob_out = _DATA_BLOB()
    ok = _crypt32.CryptUnprotectData(
        ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)
    )
    if not ok:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        if blob_out.pbData:
            _kernel32.LocalFree(blob_out.pbData)


# --------------------------------------------------------------------------- #
# 脱敏
# --------------------------------------------------------------------------- #

def mask(value: Optional[str]) -> str:
    """任何要写进日志/接口的凭据都必须先过这一层。"""
    if not value:
        return "<未配置>"
    text = str(value)
    if len(text) <= 6:
        return "*" * len(text)
    return f"{text[:3]}{'*' * 6}{text[-2:]}（{len(text)} 字符）"


# --------------------------------------------------------------------------- #
# 凭据存储
# --------------------------------------------------------------------------- #

_MAGIC = "ECOPV-OPENAPI-CRED/1"


class OpenApiCredential:
    """一份凭据（client_key + client_secret），明文只活在实例属性里。"""

    __slots__ = ("client_key", "client_secret", "source")

    def __init__(self, client_key: str, client_secret: str, source: str = "") -> None:
        self.client_key = client_key
        self.client_secret = client_secret
        self.source = source

    def __repr__(self) -> str:  # 防止 print / traceback 意外吐明文
        return f"<OpenApiCredential key={mask(self.client_key)} secret={mask(self.client_secret)} source={self.source}>"

    __str__ = __repr__

    def as_dict(self) -> dict:  # 仅供内部签名使用，调用方不得落盘/打印
        return {"client_key": self.client_key, "client_secret": self.client_secret}


def save_credential(path, env_name: str, client_key: str, client_secret: str) -> Path:
    """把某套环境（prod/test）的凭据加密写入文件，已存在的其他环境保持不变。"""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)

    payload = {}
    if target.exists():
        try:
            raw = _dpapi_unprotect(target.read_bytes())
            if raw.startswith(_MAGIC.encode()):
                payload = json.loads(raw[len(_MAGIC):].decode("utf-8"))
        except Exception:
            payload = {}

    payload[env_name] = {"client_key": client_key, "client_secret": client_secret}
    blob = _MAGIC.encode() + json.dumps(payload, ensure_ascii=False).encode("utf-8")
    target.write_bytes(_dpapi_protect(blob))
    try:
        os.chmod(target, 0o600)
    except Exception:
        pass
    return target


def load_credential(path, env_name: str) -> Optional[OpenApiCredential]:
    """读取并解密某套环境的凭据；失败一律返回 None，绝不抛明文。"""
    target = Path(path)
    if not target.exists():
        return None
    try:
        raw = _dpapi_unprotect(target.read_bytes())
    except Exception:
        return None
    if not raw.startswith(_MAGIC.encode()):
        return None
    try:
        payload = json.loads(raw[len(_MAGIC):].decode("utf-8"))
        item = payload.get(env_name) or {}
        key = item.get("client_key") or ""
        secret = item.get("client_secret") or ""
    except Exception:
        return None
    if not key or not secret:
        return None
    return OpenApiCredential(key, secret, source=f"dpapi:{target.name}:{env_name}")


def load_from_env(prefix: str = "CRM") -> Optional[OpenApiCredential]:
    """从环境变量取（仅用于本机一次性导入或 CI），生产运行不依赖它。"""
    key = os.environ.get(f"{prefix}_CLIENT_KEY", "").strip()
    secret = os.environ.get(f"{prefix}_CLIENT_SECRET", "").strip()
    if key and secret:
        return OpenApiCredential(key, secret, source="env")
    return None


def resolve_credential(path, env_name: str, prefix: str = "CRM") -> Optional[OpenApiCredential]:
    """环境变量优先，其次 DPAPI 密文文件。"""
    return load_from_env(prefix) or load_credential(path, env_name)


def configured_envs(path) -> list:
    """列出已导入了凭据的环境名（不含任何值）。"""
    target = Path(path)
    if not target.exists():
        return []
    try:
        raw = _dpapi_unprotect(target.read_bytes())
    except Exception:
        return []
    if not raw.startswith(_MAGIC.encode()):
        return []
    try:
        return sorted(json.loads(raw[len(_MAGIC):].decode("utf-8")).keys())
    except Exception:
        return []


__all__ = [
    "OpenApiCredential",
    "save_credential",
    "load_credential",
    "load_from_env",
    "resolve_credential",
    "configured_envs",
    "mask",
    "base64",
]
