"""ECOPV-EU 业务系统开放平台客户端（替代阶段二的浏览器 RPA）。

认证：POST /open/auth/token，header 带 clientId / timestamp / sign
      sign = Base64(SHA256(clientId + timestamp + clientKey + clientSecret))
      取得 token 后所有请求带 header X-Access-Token，权限等同页面登录。

纪律：clientKey / clientSecret 只从 OpenApiCredential 内存对象取，
      任何日志/异常里出现凭据值都必须先过 mask()。
"""

from __future__ import annotations

import base64
import hashlib
import json
import time
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Optional

from utils.openapi_credential import OpenApiCredential, mask


class OpenApiError(RuntimeError):
    pass


class OpenApiClient:
    """最小可用的开放平台客户端：取 token + 查注册工单。"""

    def __init__(
        self,
        base_url: str,
        client_id: str,
        credential: OpenApiCredential,
        timeout: int = 30,
        logger=None,
        sign_mode: str = "raw",
    ) -> None:
        self.base_url = str(base_url or "").rstrip("/")
        self.client_id = str(client_id or "")
        self._cred = credential
        self.timeout = timeout
        self._log = logger
        self.sign_mode = sign_mode  # raw=Base64(摘要原始字节) / hex=Base64(十六进制串)
        self._token = ""
        self._token_expire_at = 0.0

    # ---------------- 内部 ---------------- #

    def _emit(self, message: str, level: str = "info") -> None:
        if callable(self._log):
            try:
                self._log(message, level)
            except TypeError:
                self._log(message)
        elif hasattr(self._log, "info"):
            getattr(self._log, level if hasattr(self._log, level) else "info")(message)

    def _sign(self, timestamp_ms: str) -> str:
        raw = f"{self.client_id}{timestamp_ms}{self._cred.client_key}{self._cred.client_secret}"
        digest = hashlib.sha256(raw.encode("utf-8")).digest()
        if self.sign_mode == "hex":
            digest = hashlib.sha256(raw.encode("utf-8")).hexdigest().encode("utf-8")
        return base64.b64encode(digest).decode("ascii")

    def _request(self, method: str, path: str, params=None, headers=None, body=None):
        url = f"{self.base_url}{path}"
        if params:
            cleaned = {k: v for k, v in params.items() if v not in (None, "", [])}
            if cleaned:
                url = f"{url}?{urllib.parse.urlencode(cleaned, doseq=True)}"
        data = None
        if body is not None:
            data = json.dumps(body).encode("utf-8")
        req = urllib.request.Request(url, data=data, method=method)
        for key, value in (headers or {}).items():
            req.add_header(key, value)
        if data is not None:
            req.add_header("Content-Type", "application/json")
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                text = resp.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as exc:  # 4xx/5xx 也可能带 JSON body
            text = exc.read().decode("utf-8", errors="replace") if hasattr(exc, "read") else ""
            raise OpenApiError(f"HTTP {exc.code}: {text[:300]}") from None
        except Exception as exc:
            raise OpenApiError(f"请求失败: {type(exc).__name__}") from None
        try:
            return json.loads(text)
        except Exception:
            return {"_raw": text[:2000]}

    # ---------------- 认证 ---------------- #

    def token(self, force: bool = False) -> str:
        now = time.time()
        if self._token and not force and now < self._token_expire_at:
            return self._token
        ts = str(int(now * 1000))
        payload = self._request(
            "POST",
            "/open/auth/token",
            headers={
                "clientId": self.client_id,
                "timestamp": ts,
                "sign": self._sign(ts),
            },
        )
        result = payload.get("result") if isinstance(payload, dict) else None
        token = ""
        expires = 0
        if isinstance(result, dict):
            token = result.get("token") or result.get("accessToken") or result.get("value") or ""
            expires = int(result.get("expiresIn") or result.get("expire") or result.get("expires") or 0)
        elif isinstance(result, str):
            token = result
        if not token:
            keys = list(result.keys()) if isinstance(result, dict) else type(result).__name__
            self._emit(f"取 token 失败: success={payload.get('success')} message={payload.get('message')} result键={keys}")
            raise OpenApiError(f"未取得 token: {payload.get('message') or '未知原因'}")
        self._token = token
        self._token_expire_at = now + max(60, expires - 60 if expires else 1800)
        self._emit(f"已取得开放平台 token（有效期 {expires or '未知'} 秒，clientId={mask(self.client_id)}）")
        return token

    def _auth_headers(self) -> dict:
        return {"X-Access-Token": self.token()}

    # ---------------- 业务查询 ---------------- #

    def query_register_list(
        self,
        customer_code: str = "",
        country_code: str = "",
        biz_type: str = "",
        agent_user_ids=None,
        page_no: int = 1,
        page_size: int = 20,
        **extra,
    ) -> dict:
        """注册工单分页查询，参数与 RPA 页面填写条件一一对应。"""
        params: dict = {
            "pageNo": page_no,
            "pageSize": page_size,
            "customerCode": customer_code,
            "countryCode": country_code,
            "bizType": biz_type,
        }
        if agent_user_ids:
            params["agentUserIds"] = agent_user_ids
        params.update(extra)
        return self._request("GET", "/work/orderRegister/queryPageList", params=params, headers=self._auth_headers())

    def query_register_by_id(self, register_id: str) -> dict:
        return self._request(
            "GET", "/work/orderRegister/queryById", params={"id": register_id}, headers=self._auth_headers()
        )

    def agent_list(self, keyword: str = "") -> dict:
        return self._request(
            "GET", "/base/partner/agent/list",
            params={"keyword": keyword} if keyword else None,
            headers=self._auth_headers(),
        )

    # ---------------- 敏感字段解密（enc: 前缀） ---------------- #

    def encrypted_key_iv(self) -> tuple:
        payload = self._request("GET", "/sys/getEncryptedString")
        result = payload.get("result") if isinstance(payload, dict) else None
        if isinstance(result, dict):
            return result.get("key", ""), result.get("iv", "")
        return "", ""

    @staticmethod
    def decrypt_field(value: str, key: str, iv: str) -> str:
        """解 enc: 前缀字段：AES-128-CBC/NoPadding，末尾去填充。

        实测（2026-10-15）：服务端填充的是 b'\\x00' 空字节，而接口文档写的是
        "去除末尾 ASCII 1~16"。两者都覆盖：一并剥离 ord <= 16 的尾字节。
        """
        if not isinstance(value, str) or not value.startswith("enc:"):
            return value
        if not key or not iv:
            return value
        try:
            from Crypto.Cipher import AES  # 惰性导入，无该库时原样返回
        except Exception:
            return value
        try:
            cipher = AES.new(key.encode("utf-8")[:16], AES.MODE_CBC, iv.encode("utf-8")[:16])
            raw = cipher.decrypt(base64.b64decode(value[4:]))
            text = raw.decode("utf-8", errors="ignore")
            # 剥离尾部填充：文档口径 ASCII 1~16 + 实测的 \x00
            end = len(text)
            while end > 0 and ord(text[end - 1]) <= 16:
                end -= 1
            return text[:end]
        except Exception:
            return value


def build_from_config(config: dict, credential: OpenApiCredential, logger=None) -> Optional[OpenApiClient]:
    """从 config.yaml 的 crm_api 段构造客户端；未启用返回 None。"""
    section = (config or {}).get("crm_api") or {}
    if not section.get("enabled", False):
        return None
    env_name = str(section.get("env") or "prod")
    profiles = section.get("profiles") or {}
    profile = profiles.get(env_name) or {}
    base_url = str(section.get("base_url") or profile.get("base_url") or "").strip()
    client_id = str(section.get("client_id") or profile.get("client_id") or "").strip()
    if not base_url or not client_id:
        if logger:
            logger(f"开放平台未配置完整（base_url/client_id 缺失，env={env_name}）", "warning")
        return None
    return OpenApiClient(
        base_url=base_url,
        client_id=client_id,
        credential=credential,
        timeout=int(section.get("timeout", 30)),
        logger=logger,
        sign_mode=str(section.get("sign_mode", "raw")),
    )
