"""M1 邮件读取模块 — 连接阿里企业邮箱 IMAP，按日期范围拉取收件箱邮件。

先由 IMAP 服务端按所选日期范围筛选候选 UID，再拉取候选邮件日期头并
下载最终命中的邮件；附件解析在断开 IMAP 后离线执行。
"""
import os
import email
import imaplib
import json
import re
import tempfile
import time
from email.header import decode_header
from email.utils import getaddresses, parsedate_to_datetime
from datetime import datetime, timedelta
from typing import List, Dict

from bs4 import BeautifulSoup
from utils.runtime_paths import APP_ROOT


def _decode_str(s):
    """解码邮件头部字符串"""
    if s is None:
        return ""
    parts = decode_header(s)
    result = []
    for data, charset in parts:
        if isinstance(data, bytes):
            try:
                result.append(data.decode(charset or "utf-8", errors="replace"))
            except (LookupError, Exception):
                result.append(data.decode("utf-8", errors="replace"))
        else:
            result.append(data)
    return "".join(result)


def _html_to_text(html_str):
    """HTML转纯文本"""
    soup = BeautifulSoup(html_str, "lxml")
    for tag in soup.find_all(["style", "script"]):
        tag.decompose()
    return soup.get_text(separator="\n", strip=True)


def _simplify_body(body_text):
    """精简正文 — 去除常见签名/寒暄"""
    lines = body_text.split("\n")
    kept = []
    skip_patterns = [
        "请审核资料",
        "有问题请联系",
        "谢谢",
        "祝好",
        "best regards",
        "regards",
        "--",
        "发件人",
        "发件时间",
        "收件人",
        "主题",
        "本邮件及其附件",
        "confidential",
    ]
    for line in lines:
        line = line.strip()
        if not line:
            continue
        lower = line.lower()
        if any(p in lower for p in skip_patterns):
            continue
        kept.append(line)
    return "\n".join(kept) if kept else body_text[:500]


def _parse_imap_internal_date(value):
    """将 IMAP FETCH 返回的 INTERNALDATE 转成 datetime。

    邮件自身的 Date 头可能缺失或损坏，但 IMAP 服务器仍会返回该邮件
    的内部到达时间。这个时间只作为展示/筛选兜底，不覆盖合法的 Date 头。
    """
    if not value:
        return None
    if isinstance(value, bytes):
        value = value.decode("ascii", errors="replace")
    try:
        text = str(value).strip()
        # 缓存侧车文件使用 datetime.isoformat() 保存；它不是 RFC 2822
        # 日期，必须先走 fromisoformat，否则缓存命中时会丢掉兜底日期。
        try:
            return datetime.fromisoformat(text)
        except ValueError:
            return parsedate_to_datetime(text)
    except (TypeError, ValueError, OverflowError):
        return None


def _extract_fetch_payload(msg_data):
    """从 IMAP FETCH 响应中提取 RFC822 正文和 INTERNALDATE。

    不同 IMAP 服务端对响应列表的尾部格式略有差异，因此不能固定取
    ``msg_data[0][1]`` 后再丢掉元数据。保留旧响应兼容性的同时解析
    ``INTERNALDATE \"...\"``，供缺失 Date 头的邮件使用。
    """
    raw_bytes = None
    metadata = b""
    for item in msg_data or []:
        if isinstance(item, tuple) and len(item) >= 2:
            header, payload = item[0], item[1]
            if isinstance(header, bytes):
                metadata += b" " + header
            elif header:
                metadata += b" " + str(header).encode("ascii", errors="replace")
            if isinstance(payload, bytes) and payload:
                # RFC822 正文在标准响应中位于 tuple 的第二项；取最长的
                # 一项可兼容少数服务端把额外字节拆成多个响应项的情况。
                if raw_bytes is None or len(payload) > len(raw_bytes):
                    raw_bytes = payload
        elif isinstance(item, bytes):
            metadata += b" " + item

    internal_date = None
    match = re.search(rb"INTERNALDATE\s+\"([^\"]+)\"", metadata, re.IGNORECASE)
    if match:
        internal_date = _parse_imap_internal_date(match.group(1))
    return raw_bytes, internal_date


def _parse_date_header(raw_bytes):
    """从完整邮件或仅邮件头的 FETCH 载荷解析 Date 头。"""
    if not raw_bytes:
        return None
    try:
        message = email.message_from_bytes(raw_bytes)
        value = message.get("Date", "")
        return parsedate_to_datetime(value) if value else None
    except (TypeError, ValueError, OverflowError):
        return None


def _extract_date_fetch_records(msg_data, fallback_uid=None):
    """解析 UID FETCH (UID INTERNALDATE BODY.PEEK[HEADER.FIELDS (DATE)]) 响应。"""
    records = {}
    for item in msg_data or []:
        if not isinstance(item, tuple) or len(item) < 2:
            continue
        metadata = item[0]
        if isinstance(metadata, bytes):
            header = metadata
        else:
            header = str(metadata or "").encode("ascii", errors="replace")
        payload = item[1] if isinstance(item[1], bytes) else b""

        uid_match = re.search(rb"\bUID\s+(\d+)\b", header, re.IGNORECASE)
        uid = uid_match.group(1).decode("ascii") if uid_match else fallback_uid
        if not uid:
            continue

        internal_match = re.search(
            rb'INTERNALDATE\s+"([^"]+)"', header, re.IGNORECASE
        )
        internal_date = (
            _parse_imap_internal_date(internal_match.group(1))
            if internal_match else None
        )
        records[str(uid)] = (_parse_date_header(payload), internal_date)
    return records


def _date_is_inclusive_range(value, date_from, date_to):
    """按 Date 头/兜底日期所在时区的日历日做闭区间筛选。"""
    if value is None:
        return False
    start_day = date_from.date() if isinstance(date_from, datetime) else date_from
    end_day = date_to.date() if isinstance(date_to, datetime) else date_to
    return start_day <= value.date() <= end_day


def _format_imap_search_date(value):
    """格式化 IMAP 日期参数，使用协议要求的英文月份缩写。"""
    day = value.date() if isinstance(value, datetime) else value
    months = (
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    )
    return f"{day.day:02d}-{months[day.month - 1]}-{day.year:04d}"


class MailReader:
    def __init__(self, config: dict, logger=None):
        self.server_addr = config["imap_server"]
        self.port = config["imap_port"]
        self.address = config["address"]
        self.password = config["password"]
        self.mailbox = config.get("mailbox", "INBOX")
        self.connect_timeout = max(5, int(config.get("connect_timeout", config.get("timeout", 30))))
        self.connect_retries = max(1, int(config.get("connect_retries", 3)))
        self.connect_retry_delay = max(0.5, float(config.get("connect_retry_delay", 2)))
        # 正式阶段一配置关闭缓存；保留 True 默认值兼容需要离线断点续拉的旧调用方。
        self.cache_enabled = bool(config.get("cache_enabled", True))
        # 原始邮件本地缓存（按 mailbox+UIDVALIDITY+UID 落盘，重跑同范围直接命中，
        # 断连不丢已拉数据）。UIDVALIDITY 变化意味着服务器 UID 命名空间已重置，
        # 旧缓存绝不能复用。
        self.cache_dir = config.get("cache_dir") or os.path.join(
            str(APP_ROOT), "cache", "mails"
        )
        # 每拉取 N 封（真实网络拉取，不含缓存命中）重连一次
        self.reconnect_batch_size = config.get("reconnect_batch_size", 20)
        self.logger = logger
        # 日期范围内的邮件数（按 Date 头优先、本地筛选后的结果），阶段一会把它
        # 单独写入工作台摘要。last_range_candidate_total 只统计 IMAP 日期查询候选，
        # 不是整个邮箱的邮件数；last_mailbox_total 保留为兼容别名。
        self.last_search_total = 0
        self.last_search_date_from = ""
        self.last_search_date_to = ""
        self.last_range_candidate_total = 0
        self.last_mailbox_total = 0
        self.last_header_date_match_total = 0
        self.last_internal_date_fallback_total = 0
        self.last_unknown_date_total = 0

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    def _connect(self, ctx):
        """创建新的 IMAP 连接；网络瞬断时有限重试，避免一次超时直接终止阶段一。"""
        last_error = None
        for attempt in range(1, self.connect_retries + 1):
            conn = None
            try:
                self._log(
                    f"连接 IMAP {self.server_addr}:{self.port}（第 {attempt}/{self.connect_retries} 次）"
                )
                conn = imaplib.IMAP4_SSL(
                    self.server_addr,
                    self.port,
                    ssl_context=ctx,
                    timeout=self.connect_timeout,
                )
                conn.login(self.address, self.password)
                select_status, select_data = conn.select(self.mailbox, readonly=True)
                if select_status not in ("OK", b"OK"):
                    raise imaplib.IMAP4.error(
                        f"选择邮箱失败: {self.mailbox} ({select_status})"
                    )
                selected_count = "未知"
                if select_data:
                    try:
                        selected_count = str(int(select_data[0]))
                    except (TypeError, ValueError):
                        selected_count = str(select_data[0])
                self._log(
                    f"已选择邮箱: {self.mailbox} (服务器报告邮件数={selected_count})"
                )
                return conn
            except (OSError, TimeoutError, imaplib.IMAP4.error) as exc:
                last_error = exc
                if conn is not None:
                    try:
                        conn.logout()
                    except Exception:
                        pass
                if attempt >= self.connect_retries:
                    break
                delay = self.connect_retry_delay * attempt
                self._log(f"IMAP 连接失败: {exc}；{delay:g} 秒后重试", "warning")
                time.sleep(delay)
        raise TimeoutError(
            f"无法连接 IMAP 服务器 {self.server_addr}:{self.port}（已重试 {self.connect_retries} 次）：{last_error}。"
            "请检查网络/VPN代理、企业邮箱 IMAP 开关及端口 993。"
        ) from last_error

    @staticmethod
    def _uidvalidity(conn) -> str:
        """读取当前邮箱 UIDVALIDITY；取不到时使用独立的 unknown 命名空间。"""
        try:
            _, values = conn.response("UIDVALIDITY")
            if values:
                value = values[-1]
                if isinstance(value, bytes):
                    value = value.decode("ascii", errors="ignore")
                value = str(value).strip()
                if value:
                    return value
        except Exception:
            pass
        return "unknown"

    def fetch_mails(
        self,
        date_from,
        date_to,
        self_email: str,
        progress_callback=None,
        cancel_requested=None,
    ) -> List[Dict]:
        """
        拉取指定时间范围内的收件箱邮件
        分批拉取: 每 20 封重连一次，避免连接超时
        date_from/date_to: 支持 datetime 对象或 'YYYY-MM-DD' 字符串
        """
        self.last_fetch_complete = False
        def cancelled(stage):
            if cancel_requested is not None and cancel_requested():
                self._log(f"邮件读取已停止：{stage}；不提交部分结果")
                return True
            return False

        if cancelled("连接前"):
            return []
        if isinstance(date_from, str):
            date_from = datetime.strptime(date_from, "%Y-%m-%d")
        if isinstance(date_to, str):
            date_to = datetime.strptime(date_to, "%Y-%m-%d")

        start_day = date_from.date() if isinstance(date_from, datetime) else date_from
        end_day = date_to.date() if isinstance(date_to, datetime) else date_to
        if end_day < start_day:
            raise ValueError("邮件日期范围结束日期不能早于开始日期")

        import ssl
        ctx = ssl.create_default_context()
        conn = None
        # 每项为 (uid, RFC822 原文, IMAP INTERNALDATE 兜底时间)。
        # 第三项允许 Date 头缺失的邮件仍能进入日期筛选和过滤日志。
        raw_mails = []

        self.last_search_total = 0
        self.last_range_candidate_total = 0
        self.last_mailbox_total = 0
        self.last_header_date_match_total = 0
        self.last_internal_date_fallback_total = 0
        self.last_unknown_date_total = 0
        self.last_search_date_from = date_from.strftime("%Y-%m-%d")
        self.last_search_date_to = date_to.strftime("%Y-%m-%d")
        self._log(
            f"邮件日期查询条件: 邮箱={self.mailbox}, "
            f"范围={start_day} ~ {end_day}, "
            "口径=邮件Date头优先，缺失/无效时使用IMAP INTERNALDATE"
        )

        try:
            conn = self._connect(ctx)
            self._log(f"邮箱连接成功，当前选择邮箱: {self.mailbox}")
            uidvalidity = self._uidvalidity(conn)

            # SINCE/BEFORE 使用服务器 INTERNALDATE；SENTSINCE/SENTBEFORE
            # 使用邮件 Date 头。分别查询并集，既限制在用户选择的日期范围，
            # 也覆盖“邮件较晚进入收件箱，但发件日期在范围内”的邮件。
            start_token = _format_imap_search_date(start_day)
            end_exclusive_token = _format_imap_search_date(
                end_day + timedelta(days=1)
            )

            def search_date_candidates(label, *criteria):
                status, data = conn.uid("SEARCH", None, *criteria)
                if status not in ("OK", b"OK"):
                    self._log(
                        f"IMAP UID 日期范围搜索失败 ({label}, 邮箱={self.mailbox}, status={status})",
                        "error",
                    )
                    raise RuntimeError(
                        f"IMAP UID 日期范围搜索失败 ({label}): {status}"
                    )
                found = []
                for response in data or []:
                    if isinstance(response, bytes):
                        found.extend(
                            uid.decode("ascii", errors="replace")
                            for uid in response.split()
                        )
                    elif response:
                        found.extend(str(response).split())
                return found

            internal_date_uids = search_date_candidates(
                "INTERNALDATE",
                "SINCE", start_token, "BEFORE", end_exclusive_token,
            )
            sent_date_uids = search_date_candidates(
                "Date头",
                "SENTSINCE", start_token, "SENTBEFORE", end_exclusive_token,
            )
            candidate_uid_set = {
                uid for uid in internal_date_uids + sent_date_uids
                if uid.isdigit()
            }
            uids = [
                uid.encode("ascii")
                for uid in sorted(candidate_uid_set, key=int)
            ]
            candidate_total = len(uids)
            self.last_range_candidate_total = candidate_total
            self.last_mailbox_total = candidate_total
            self._log(
                f"IMAP日期范围搜索完成: mailbox={self.mailbox}, "
                f"范围={start_day}~{end_day}, INTERNALDATE候选={len(set(internal_date_uids))}, "
                f"Date头候选={len(set(sent_date_uids))}, 范围候选UID数={candidate_total}, "
                f"UIDVALIDITY={uidvalidity}"
            )

            if self.cache_enabled:
                os.makedirs(self.cache_dir, exist_ok=True)
            else:
                self._log("阶段一原始邮件缓存已禁用，本次全部从 IMAP 重新拉取")
            uid_texts = [uid.decode("ascii", errors="replace") for uid in uids]
            uid_bytes_by_text = dict(zip(uid_texts, uids))
            cache_path_by_uid = {
                uid: os.path.join(
                    self.cache_dir, f"{self.mailbox}_uidv{uidvalidity}_{uid}.eml"
                )
                for uid in uid_texts
            }
            cached_raw_by_uid = {}
            date_records = {}
            metadata_uids = []

            # 有效 Date 头足以决定日期，不需要为缓存邮件额外查询 INTERNALDATE；
            # Date 头缺失/无效且侧车也无 INTERNALDATE 时，再纳入轻量 FETCH 清单。
            for uid in uid_texts:
                cache_path = cache_path_by_uid[uid]
                if not (
                    self.cache_enabled
                    and os.path.exists(cache_path)
                    and os.path.getsize(cache_path) > 0
                ):
                    metadata_uids.append(uid_bytes_by_text[uid])
                    continue
                try:
                    with open(cache_path, "rb") as cache_file:
                        raw_bytes = cache_file.read()
                    cached_internal_date = None
                    metadata_path = f"{cache_path}.meta.json"
                    if os.path.exists(metadata_path):
                        try:
                            with open(metadata_path, "r", encoding="utf-8") as meta_file:
                                cached_meta = json.load(meta_file)
                            cached_internal_date = _parse_imap_internal_date(
                                cached_meta.get("internal_date")
                            )
                        except (OSError, ValueError, TypeError, AttributeError):
                            cached_internal_date = None
                    cached_raw_by_uid[uid] = (raw_bytes, cached_internal_date)
                    header_date = _parse_date_header(raw_bytes)
                    if header_date is not None or cached_internal_date is not None:
                        date_records[uid] = (header_date, cached_internal_date)
                    else:
                        metadata_uids.append(uid_bytes_by_text[uid])
                except OSError as exc:
                    self._log(f"读缓存失败(uid={uid}): {exc}", "warning")
                    metadata_uids.append(uid_bytes_by_text[uid])

            def reconnect_for_retry(reason):
                nonlocal conn
                try:
                    conn.close()
                except Exception:
                    pass
                try:
                    conn.logout()
                except Exception:
                    pass
                self._log(f"IMAP连接重建: {reason}", "warning")
                conn = self._connect(ctx)
                current_uidvalidity = self._uidvalidity(conn)
                if uidvalidity != "unknown" and current_uidvalidity != uidvalidity:
                    raise RuntimeError(
                        "邮箱 UIDVALIDITY 在读取期间发生变化；为避免 UID 与缓存错配，已停止本次读取"
                    )

            def fetch_date_metadata(uid_batch):
                """批量读 Date 头，批量响应缺项时再按 UID 单条兜底。"""
                uid_set = ",".join(
                    uid.decode("ascii", errors="replace") for uid in uid_batch
                )
                query = "(UID INTERNALDATE BODY.PEEK[HEADER.FIELDS (DATE)])"
                try:
                    status, response = conn.uid("FETCH", uid_set, query)
                    if status not in ("OK", b"OK"):
                        raise RuntimeError(f"批量 Date 头 FETCH 返回 {status}")
                    batch_records = _extract_date_fetch_records(response)
                except Exception as exc:
                    self._log(
                        f"批量读取邮件日期头失败(uid范围={uid_set}): {exc}；改为逐UID读取",
                        "warning",
                    )
                    batch_records = {}

                requested = [
                    uid.decode("ascii", errors="replace") for uid in uid_batch
                ]
                missing = [uid for uid in requested if uid not in batch_records]
                for uid in missing:
                    try:
                        status, response = conn.uid("FETCH", uid, query)
                        if status not in ("OK", b"OK"):
                            raise RuntimeError(f"单UID Date 头 FETCH 返回 {status}")
                        batch_records.update(
                            _extract_date_fetch_records(response, fallback_uid=uid)
                        )
                    except Exception as exc:
                        try:
                            reconnect_for_retry(f"Date 头读取失败(uid={uid})")
                            status, response = conn.uid("FETCH", uid, query)
                            if status not in ("OK", b"OK"):
                                raise RuntimeError(f"重试 Date 头 FETCH 返回 {status}")
                            batch_records.update(
                                _extract_date_fetch_records(response, fallback_uid=uid)
                            )
                        except Exception as retry_exc:
                            self._log(
                                f"无法读取邮件日期头(uid={uid}): {retry_exc}；该邮件不计入日期范围",
                                "error",
                            )
                return batch_records

            header_batch_size = max(1, min(100, int(self.reconnect_batch_size or 20)))
            batches = [
                metadata_uids[index:index + header_batch_size]
                for index in range(0, len(metadata_uids), header_batch_size)
            ]
            reconnect_every = max(1, int(self.reconnect_batch_size or 20))
            for batch_index, uid_batch in enumerate(batches):
                if cancelled("邮件日期头扫描"):
                    return []
                if batch_index and batch_index % reconnect_every == 0:
                    reconnect_for_retry(
                        f"Date 头扫描已处理约 {batch_index * header_batch_size} 封"
                    )
                date_records.update(fetch_date_metadata(uid_batch))
                if (batch_index + 1) % 10 == 0:
                    self._log(
                        f"邮件日期头扫描进度: {min((batch_index + 1) * header_batch_size, len(metadata_uids))}"
                        f"/{len(metadata_uids)}"
                    )

            eligible_uids = []
            header_date_matches = 0
            internal_date_fallbacks = 0
            unknown_dates = 0
            out_of_range = 0
            for uid in uid_texts:
                header_date, internal_date = date_records.get(uid, (None, None))
                if header_date is not None:
                    effective_date = header_date
                    if _date_is_inclusive_range(effective_date, start_day, end_day):
                        header_date_matches += 1
                        eligible_uids.append(uid)
                    else:
                        # 合法 Date 头是主口径；即使服务器 INTERNALDATE 在范围内，
                        # 也不能覆盖一个明确的邮件发件日期。
                        out_of_range += 1
                elif internal_date is not None:
                    effective_date = internal_date
                    if _date_is_inclusive_range(effective_date, start_day, end_day):
                        internal_date_fallbacks += 1
                        eligible_uids.append(uid)
                    else:
                        out_of_range += 1
                else:
                    unknown_dates += 1

            self.last_search_total = len(eligible_uids)
            self.last_header_date_match_total = header_date_matches
            self.last_internal_date_fallback_total = internal_date_fallbacks
            self.last_unknown_date_total = unknown_dates
            self._log(
                f"本地邮件日期筛选: 范围={start_day} ~ {end_day}, "
                f"Date头命中={header_date_matches}, INTERNALDATE兜底命中={internal_date_fallbacks}, "
                f"范围外={out_of_range}, 日期未知={unknown_dates}, 符合范围={len(eligible_uids)}"
            )

            cache_hits = 0
            fetch_uids = []
            for uid in eligible_uids:
                cached = cached_raw_by_uid.get(uid)
                if cached:
                    raw_mails.append((uid, cached[0], cached[1]))
                    cache_hits += 1
                else:
                    fetch_uids.append(uid_bytes_by_text[uid])

            fetched = 0  # 真实完整邮件网络拉取数；日期头 FETCH 不计入此数。
            batch_size = max(1, int(self.reconnect_batch_size or 20))
            for idx, uid_bytes in enumerate(fetch_uids):
                if cancelled(f"完整邮件下载 {idx}/{len(fetch_uids)}"):
                    return []
                uid = uid_bytes.decode("ascii", errors="replace")

                # 每拉取 BATCH_SIZE 封完整邮件重连一次。
                if fetched > 0 and fetched % batch_size == 0:
                    reconnect_for_retry(f"完整邮件已拉取 {fetched}/{len(fetch_uids)} 封")

                try:
                    status, msg_data = conn.uid(
                        "FETCH", uid_bytes, "(RFC822 INTERNALDATE)"
                    )
                except Exception as exc:
                    self._log(f"拉取完整邮件失败(uid={uid}): {exc}", "warning")
                    try:
                        reconnect_for_retry(f"完整邮件FETCH失败(uid={uid})")
                        status, msg_data = conn.uid(
                            "FETCH", uid_bytes, "(RFC822 INTERNALDATE)"
                        )
                    except Exception as retry_exc:
                        self._log(f"完整邮件重试失败(uid={uid}): {retry_exc}", "warning")
                        continue
                if status not in ("OK", b"OK") or not msg_data or not msg_data[0]:
                    self._log(f"完整邮件 FETCH 无结果(uid={uid})", "warning")
                    continue

                raw_bytes, internal_date = _extract_fetch_payload(msg_data)
                if not raw_bytes:
                    self._log(f"FETCH 响应缺少 RFC822 正文(uid={uid})，已跳过", "warning")
                    continue
                raw_mails.append((uid, raw_bytes, internal_date))
                fetched += 1

                # 落盘缓存（失败不影响主流程）
                if self.cache_enabled:
                    cache_path = cache_path_by_uid[uid]
                    try:
                        with open(cache_path, "wb") as cache_file:
                            cache_file.write(raw_bytes)
                        if internal_date is not None:
                            with open(
                                f"{cache_path}.meta.json", "w", encoding="utf-8"
                            ) as meta_file:
                                json.dump(
                                    {"internal_date": internal_date.isoformat()},
                                    meta_file,
                                    ensure_ascii=False,
                                )
                    except OSError as exc:
                        self._log(f"写缓存失败(uid={uid}): {exc}", "warning")

                if progress_callback:
                    progress_callback(idx + 1, len(fetch_uids))

            self._log(
                f"原始邮件拉取完成: {len(raw_mails)} 封 "
                f"(完整邮件网络拉取 {fetched}, 缓存命中 {cache_hits}, "
                f"范围候选UID数 {candidate_total}, 日期命中 {len(eligible_uids)})"
            )

        except Exception as e:
            self._log(f"IMAP读取/连接失败: {e}", "error")
            raise
        finally:
            if conn:
                try:
                    conn.close()
                except Exception:
                    pass
                try:
                    conn.logout()
                except Exception:
                    pass
            self._log("已断开邮箱连接")

        # === 阶段二: 离线解析每封邮件的正文和附件 ===
        if cancelled("附件解析前"):
            return []
        complete_download = len(raw_mails) == len(eligible_uids) and not unknown_dates
        self._log("开始离线解析...")
        mails = []
        for idx, raw_item in enumerate(raw_mails):
            if cancelled(f"邮件与附件解析 {idx}/{len(raw_mails)}"):
                return []
            # 兼容旧的二元缓存/调用方；新流程第三项为 IMAP INTERNALDATE。
            uid, raw_bytes = raw_item[0], raw_item[1]
            internal_date = raw_item[2] if len(raw_item) >= 3 else None
            msg = email.message_from_bytes(raw_bytes)

            sender = _decode_str(msg.get("From", ""))
            sender_email = self._extract_email_addr(sender)
            # 保留原始收件人头。阶段一过去只导出了发件人，导致人工工作台
            # 无法展示或追溯邮件实际投递对象；这里不改过滤或提取判断，仅补齐证据字段。
            recipient = _decode_str(msg.get("To", ""))
            recipient_emails = [
                address.strip().lower()
                for _, address in getaddresses([recipient])
                if address and "@" in address
            ]
            subject = _decode_str(msg.get("Subject", ""))
            date_header = msg.get("Date", "")
            try:
                date = parsedate_to_datetime(date_header) if date_header else None
            except (TypeError, ValueError, OverflowError) as exc:
                # 日期头扫描阶段已确认这封邮件的 INTERNALDATE 落在筛选范围内；
                # 这里保留邮件并由 INTERNALDATE 补齐日期，而不是让整批阶段一崩溃。
                date = None
                self._log(
                    f"邮件日期头无效(uid={uid})，已保留邮件并标记日期待复核: {exc}",
                    "warning",
                )

            date_source = "邮件 Date 头" if date is not None else ""
            if date is None and internal_date is not None:
                date = internal_date
                date_source = "IMAP INTERNALDATE（Date 头缺失/无效）"
                self._log(
                    f"邮件 Date 头缺失/无效(uid={uid})，已使用 IMAP INTERNALDATE: {date}",
                    "warning",
                )

            # 跳过自身发送的邮件
            if self_email.lower() in sender_email.lower():
                mails.append({
                    "uid": uid,
                    "sender_email": sender_email,
                    "sender_name": sender,
                    "recipient": recipient,
                    "recipient_emails": recipient_emails,
                    "date": date,
                    "date_source": date_source,
                    "subject": subject,
                    "body_text": "",
                    "body_original": "",
                    "body_raw": "",
                    "attachments": [],
                    "skip_reason": "self_sent",
                })
                continue

            body_text, body_raw, body_original, attachments = self._parse_msg_content(msg)

            mails.append({
                "uid": uid,
                "sender_email": sender_email,
                "sender_name": sender,
                "recipient": recipient,
                "recipient_emails": recipient_emails,
                "date": date,
                "date_source": date_source,
                "subject": subject,
                "body_text": body_text,
                # body_text 用于规则/LLM 抽取；body_original 保留完整可读正文，
                # 供工作台证据区和导出文件展示，避免把清洗后的摘要误称为原文。
                "body_original": body_original,
                "body_raw": body_raw,
                "attachments": attachments,
                "skip_reason": None,
            })

            # 阶段二不再发进度回调（阶段一已按 total 汇报过，避免进度条重复回跳）
            if progress_callback:
                progress_callback(len(raw_mails), len(raw_mails))

        if cancelled("附件解析结束"):
            return []
        self._log(f"邮件解析完成, 共 {len(mails)} 封 (含跳过)")
        self.last_fetch_complete = complete_download
        return mails

    def _extract_email_addr(self, sender_str: str) -> str:
        """从发件人字段提取邮箱地址"""
        import re
        match = re.search(r"<([^>]+)>", sender_str)
        if match:
            return match.group(1).strip()
        match = re.search(r"[\w.+-]+@[\w.-]+\.\w+", sender_str)
        if match:
            return match.group(0).strip()
        return sender_str.strip().lower()

    def _parse_msg_content(self, msg) -> tuple:
        """解析邮件正文和附件"""
        body_text = ""
        body_raw = ""
        attachments = []

        if msg.is_multipart():
            for part in msg.walk():
                content_disposition = str(part.get("Content-Disposition", ""))
                content_type = part.get_content_type()
                filename = part.get_filename()
                filename = _decode_str(filename) if filename else None

                if "attachment" in content_disposition.lower() or filename:
                    if filename:
                        payload = part.get_payload(decode=True)
                        if payload:
                            ext = os.path.splitext(filename)[1].lower()
                            if ext in (".jpg", ".jpeg", ".png", ".bmp", ".tiff"):
                                # 图片附件: 持久化到 cache/attachments/, 登记 ocr_pending
                                # OCR 仅在字段提取兜底阶段按需触发（懒加载）
                                from utils.attachment_parser import (
                                    parse_attachment,
                                    sanitize_attachment_filename,
                                    save_attachment,
                                )
                                safe_filename = sanitize_attachment_filename(filename)
                                try:
                                    img_path = save_attachment(payload, safe_filename)
                                    att = parse_attachment(img_path, safe_filename)
                                    att["filename"] = safe_filename
                                    att["filepath"] = img_path
                                    attachments.append(att)
                                except Exception as exc:
                                    # 单个异常附件不能中断整批邮件；正文和其他附件仍继续解析。
                                    self._log(
                                        f"图片附件保存/解析失败，已跳过: {safe_filename} ({exc})",
                                        "warning",
                                    )
                            else:
                                from utils.attachment_parser import (
                                    parse_attachment,
                                    sanitize_attachment_filename,
                                    save_attachment,
                                )
                                safe_filename = sanitize_attachment_filename(filename)
                                # 解析仍使用临时文件，但原始附件同时持久化，
                                # 使工作台之后可以真正打开/下载 xlsx、pdf、zip 等。
                                stored_path = save_attachment(payload, safe_filename)
                                with tempfile.NamedTemporaryFile(
                                    delete=False, suffix=ext
                                ) as tmp:
                                    tmp.write(payload)
                                    tmp_path = tmp.name
                                try:
                                    att = parse_attachment(tmp_path, safe_filename)
                                    att["filename"] = safe_filename
                                    att["filepath"] = stored_path
                                    attachments.append(att)
                                    # 压缩包内的待识别图片合并进附件列表（懒加载 OCR）
                                    for p in att.pop("pending_images", []):
                                        attachments.append(p)
                                finally:
                                    try:
                                        os.unlink(tmp_path)
                                    except OSError:
                                        pass
                elif content_type == "text/plain":
                    payload = part.get_payload(decode=True)
                    if payload and not body_text:
                        charset = part.get_content_charset() or "utf-8"
                        body_raw = payload.decode(charset, errors="replace")
                        body_text = body_raw
                elif content_type == "text/html" and not body_text:
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or "utf-8"
                        html_str = payload.decode(charset, errors="replace")
                        body_raw = html_str
                        body_text = _html_to_text(html_str)
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                content_type = msg.get_content_type()
                raw = payload.decode(charset, errors="replace")
                body_raw = raw
                if content_type == "text/html":
                    body_text = _html_to_text(raw)
                else:
                    body_text = raw

        body_original = body_text
        body_text = _simplify_body(body_text)
        return body_text, body_raw, body_original, attachments
