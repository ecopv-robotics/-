"""本地工作台数据库。

数据库只使用 Python 标准库 ``sqlite3``，用于把阶段一/阶段二输入和人工操作
从一次运行的 Excel/JSON 状态中独立出来。Excel 仍作为兼容输入保留，但新数据
会按邮件身份增量写入，不会因为下一次程序运行而覆盖历史记录。
"""

from __future__ import annotations

import hashlib
import json
import sqlite3
import threading
from contextlib import contextmanager, nullcontext
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

# Publishing a refreshed day and interactive review must not interleave.
WORKBENCH_DATA_LOCK = threading.RLock()


def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _json_value(value: Any) -> Any:
    """把 Excel/日期对象转换成可稳定写入 SQLite JSON 的值。"""
    if isinstance(value, (datetime, date)):
        return value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    if isinstance(value, dict):
        return {str(k): _json_value(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_value(v) for v in value]
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _date_key(value: Any) -> str:
    text = _text(value).replace("T", " ")
    if not text:
        return ""
    for candidate in (text, text[:19], text[:10]):
        try:
            return datetime.fromisoformat(candidate).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
    return text


def _mail_identity(row: Dict[str, Any]) -> str:
    sender = _text(row.get("发件人邮箱") or row.get("sender_email") or row.get("sender")).lower()
    date_value = _date_key(row.get("发件日期") or row.get("date"))
    subject = _text(row.get("邮件主题") or row.get("主题") or row.get("subject"))
    return "|".join((sender, date_value, subject))


def _record_base(row: Dict[str, Any]) -> str:
    company = _text(row.get("客户公司名称") or row.get("客户") or row.get("company"))
    project = _text(row.get("标准化项目名称") or row.get("项目") or row.get("program"))
    request = _text(row.get("需求") or row.get("request"))
    return "|".join((_mail_identity(row), company, project, request))


def _canonical_payload(row: Dict[str, Any]) -> Dict[str, Any]:
    """兼容阶段一内部字典、Excel 导出表头和 API 邮件字典。"""
    payload = _json_value(dict(row))
    if not isinstance(payload, dict):
        return {}
    aliases = {
        "发件人邮箱": ("sender_email", "sender"),
        "收件人": ("recipient", "recipient_email"),
        "发件日期": ("date",),
        "邮件主题": ("subject", "主题"),
        "邮件正文摘要(最多300字)": ("body_text", "body"),
        "邮件正文原文": ("body_original", "body_raw", "body_text", "body"),
        "附件名称": ("attachments",),
        "主题": ("subject", "邮件主题"),
        "正文摘要(最多300字)": ("body_text", "body"),
        "过滤原因": ("filter_reason", "reason"),
        "意图LLM状态": ("intent_llm_status", "intent_status"),
        "处理时间戳": ("processed_at",),
        "代理": ("agent",),
        "客户公司名称": ("company", "客户"),
        "标准化项目名称": ("program", "项目"),
        "需求": ("request",),
    }
    for canonical, candidates in aliases.items():
        if _text(payload.get(canonical)):
            continue
        for candidate in candidates:
            value = payload.get(candidate)
            if isinstance(value, list):
                names = []
                for item in value:
                    if isinstance(item, dict):
                        name = _text(item.get("filename") or item.get("name"))
                    else:
                        name = _text(item)
                    if name:
                        names.append(name)
                value = "；".join(names)
            if _text(value):
                payload[canonical] = value
                break
    payload.setdefault("_source", _text(row.get("_source")) or "数据库导入")
    return payload


def _key(value: str) -> str:
    return hashlib.sha1(value.encode("utf-8", errors="ignore")).hexdigest()[:24]


def _mail_number(mail_key: str) -> str:
    """给邮件生成可展示、可跨重启复用的稳定编号。"""
    return f"MAIL-{_text(mail_key).upper()}" if _text(mail_key) else ""


def _detail_number(record_key: str) -> str:
    """给业务明细生成可展示、可跨重启复用的稳定编号。"""
    return f"DETAIL-{_text(record_key).upper()}" if _text(record_key) else ""


class WorkbenchDatabase:
    """线程安全的轻量数据库访问层。

    ``dataset`` 目前约定为 ``active``（询单/人工补全）、``filtered``（过滤日志）、
    ``workorder``（阶段二结果）或 ``workorder_manual``（工作台人工同步的工单结果）。
    同一数据集按邮件+业务明细身份幂等更新，
    新邮件则追加新记录。
    """

    DATASETS = {"active", "filtered", "workorder", "workorder_manual"}

    def __init__(self, path: str | Path):
        self.path = Path(path).resolve()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.path), timeout=20, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA busy_timeout = 20000")
        conn.execute("PRAGMA journal_mode = WAL")
        return conn

    def _initialize(self) -> None:
        with self._lock, self._connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS input_records (
                    record_key TEXT PRIMARY KEY,
                    dataset TEXT NOT NULL,
                    mail_key TEXT NOT NULL,
                    mail_number TEXT,
                    detail_number TEXT,
                    mail_date TEXT,
                    operation_date TEXT,
                    source_path TEXT,
                    payload_json TEXT NOT NULL,
                    first_seen_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_input_dataset_date
                    ON input_records(dataset, mail_date DESC, updated_at DESC);
                CREATE INDEX IF NOT EXISTS idx_input_mail_key
                    ON input_records(dataset, mail_key);
                CREATE TABLE IF NOT EXISTS operation_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    operation_date TEXT NOT NULL,
                    action TEXT NOT NULL,
                    record_id TEXT,
                    mail_id TEXT,
                    reason TEXT,
                    result_json TEXT,
                    created_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_operation_date
                    ON operation_log(operation_date DESC, created_at DESC);
                CREATE TABLE IF NOT EXISTS stage1_refresh_days (
                    mail_day TEXT PRIMARY KEY, refreshed_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS retired_stage1_records (
                    refresh_id TEXT NOT NULL, record_key TEXT NOT NULL,
                    dataset TEXT NOT NULL, payload_json TEXT NOT NULL,
                    PRIMARY KEY(refresh_id, record_key)
                );
                CREATE TABLE IF NOT EXISTS ingest_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    dataset TEXT NOT NULL,
                    source_path TEXT,
                    row_count INTEGER NOT NULL,
                    inserted_count INTEGER NOT NULL,
                    updated_count INTEGER NOT NULL,
                    operation_date TEXT NOT NULL
                );
                """
            )
            # 兼容第一版已经创建过的 workbench.db。
            columns = {row[1] for row in conn.execute("PRAGMA table_info(input_records)").fetchall()}
            if "mail_number" not in columns:
                conn.execute("ALTER TABLE input_records ADD COLUMN mail_number TEXT")
            if "detail_number" not in columns:
                conn.execute("ALTER TABLE input_records ADD COLUMN detail_number TEXT")
            if "operation_date" not in columns:
                conn.execute("ALTER TABLE input_records ADD COLUMN operation_date TEXT")
                conn.execute("UPDATE input_records SET operation_date=substr(updated_at,1,10) WHERE operation_date IS NULL")

    @staticmethod
    def _now() -> str:
        return datetime.now().isoformat(timespec="seconds")

    def ingest_rows(
        self,
        rows: Iterable[Dict[str, Any]],
        dataset: str = "active",
        source_path: str = "",
        _connection=None,
    ) -> Dict[str, int | str]:
        """幂等增量写入一批 Excel/阶段输出行。"""
        dataset = _text(dataset).lower() or "active"
        if dataset not in self.DATASETS:
            raise ValueError(f"未知数据库数据集：{dataset}")
        values = [row for row in rows if isinstance(row, dict)]
        if not values:
            return {"dataset": dataset, "rows": 0, "inserted": 0, "updated": 0}
        now = self._now()
        counts: Dict[str, int] = {"inserted": 0, "updated": 0}
        occurrences: Dict[str, int] = {}
        with self._lock, (nullcontext(_connection) if _connection is not None else self._connect()) as conn:
            refreshed = {r[0] for r in conn.execute("SELECT mail_day FROM stage1_refresh_days")} if _connection is None and dataset in {"active", "filtered"} else set()
            for row in values:
                if _date_key(row.get("发件日期") or row.get("date"))[:10] in refreshed:
                    continue  # Old Excel/date-cache imports must not resurrect replaced rows.
                base = _record_base(row)
                occurrence = occurrences.get(base, 0)
                occurrences[base] = occurrence + 1
                # 同一封邮件中完全相同的明细也要保留，使用本批出现序号区分。
                record_key = _key(f"{dataset}|{base}|{occurrence}")
                mail_key = _key(_mail_identity(row))
                mail_number = _mail_number(mail_key)
                # 明细编号不把 dataset 放进种子；同一邮件+公司+项目+需求在
                # active、工单结果和人工同步覆盖层中应能对应到同一个编号。
                detail_number = _detail_number(_key(f"{base}|{occurrence}"))
                payload = _canonical_payload(row)
                if not isinstance(payload, dict):
                    continue
                payload["_db_record_key"] = record_key
                payload["_db_mail_key"] = mail_key
                payload["mail_number"] = mail_number
                payload["detail_number"] = detail_number
                payload["_db_operation_date"] = now[:10]
                payload.setdefault("_legacy_id", _text(row.get("_id")))
                payload.setdefault("_source_path", _text(source_path))
                mail_date = _date_key(row.get("发件日期") or row.get("date"))
                old = conn.execute(
                    "SELECT first_seen_at FROM input_records WHERE record_key=?",
                    (record_key,),
                ).fetchone()
                conn.execute(
                    """
                    INSERT INTO input_records
                        (record_key, dataset, mail_key, mail_number, detail_number, mail_date, operation_date, source_path,
                         payload_json, first_seen_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(record_key) DO UPDATE SET
                        dataset=excluded.dataset,
                        mail_key=excluded.mail_key,
                        mail_number=excluded.mail_number,
                        detail_number=excluded.detail_number,
                        mail_date=excluded.mail_date,
                        operation_date=excluded.operation_date,
                        source_path=excluded.source_path,
                        payload_json=excluded.payload_json,
                        updated_at=excluded.updated_at
                    """,
                    (
                        record_key, dataset, mail_key, mail_number, detail_number, mail_date, now[:10], _text(source_path),
                        json.dumps(payload, ensure_ascii=False),
                        _text(old["first_seen_at"]) if old else now,
                        now,
                    ),
                )
                counts["updated" if old else "inserted"] += 1
            conn.execute(
                """
                INSERT INTO ingest_log
                    (dataset, source_path, row_count, inserted_count, updated_count, operation_date)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (dataset, _text(source_path), len(values), counts["inserted"], counts["updated"], now),
            )
        return {"dataset": dataset, "rows": len(values), **counts}

    def unrefreshed_rows(self, rows):
        """Only legacy days may be read back from Excel; refreshed days use SQLite."""
        with self._lock, self._connect() as conn:
            days = {r[0] for r in conn.execute("SELECT mail_day FROM stage1_refresh_days")}
        return [r for r in rows if _date_key(r.get("发件日期") or r.get("date"))[:10] not in days]

    @contextmanager
    def replace_stage1_range(self, active, filtered, date_from, date_to, source_path="", legacy_active=(), legacy_filtered=()):
        """Atomically retire both queues for an inclusive day range, including empty results."""
        from datetime import timedelta
        import uuid
        start = date.fromisoformat(str(date_from)[:10])
        end = date.fromisoformat(str(date_to)[:10])
        if end < start:
            raise ValueError("Invalid replacement date range")
        batches = {"active": list(active), "filtered": list(filtered)}
        for rows in batches.values():
            for row in rows:
                day = date.fromisoformat(_date_key(row.get("发件日期") or row.get("date"))[:10])
                if not start <= day <= end:
                    raise ValueError("Replacement row outside selected mail dates")
        refresh_id = uuid.uuid4().hex
        with self._lock, self._connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            refreshed = {r[0] for r in conn.execute("SELECT mail_day FROM stage1_refresh_days")}
            for dataset, legacy in (("active", legacy_active), ("filtered", legacy_filtered)):
                keep = [r for r in legacy
                        if (day := _date_key(r.get("发件日期") or r.get("date"))[:10]) not in refreshed
                        and not start.isoformat() <= day <= end.isoformat()]
                self.ingest_rows(keep, dataset, source_path, _connection=conn)
            where = "dataset IN ('active','filtered') AND substr(mail_date,1,10) BETWEEN ? AND ?"
            bounds = (start.isoformat(), end.isoformat())
            conn.execute(
                "INSERT INTO retired_stage1_records SELECT ?, record_key, dataset, payload_json FROM input_records WHERE " + where,
                (refresh_id, *bounds),
            )
            removed = conn.execute("DELETE FROM input_records WHERE " + where, bounds).rowcount
            for dataset, rows in batches.items():
                clean = [{**r, "_legacy_id": ""} for r in rows]
                self.ingest_rows(clean, dataset, source_path, _connection=conn)
            for offset in range((end - start).days + 1):
                conn.execute("INSERT OR REPLACE INTO stage1_refresh_days VALUES (?, ?)",
                             ((start + timedelta(days=offset)).isoformat(), self._now()))
            yield {"refresh_id": refresh_id, "removed": removed,
                   "active": len(batches["active"]), "filtered": len(batches["filtered"])}

    def read_rows(self, dataset: str = "active") -> List[Dict[str, Any]]:
        dataset = _text(dataset).lower() or "active"
        if dataset not in self.DATASETS:
            raise ValueError(f"未知数据库数据集：{dataset}")
        with self._lock, self._connect() as conn:
            rows = conn.execute(
                """
                SELECT record_key, payload_json
                       , mail_key, mail_number, detail_number
                FROM input_records
                WHERE dataset=?
                ORDER BY CASE WHEN mail_date='' THEN 1 ELSE 0 END,
                         mail_date DESC, updated_at DESC, record_key
                """,
                (dataset,),
            ).fetchall()
        result: List[Dict[str, Any]] = []
        missing: List[tuple[str, str, str, str]] = []
        for item in rows:
            try:
                payload = json.loads(item["payload_json"])
            except (TypeError, ValueError):
                continue
            if not isinstance(payload, dict):
                continue
            payload["_id"] = _text(payload.get("_db_record_key")) or _text(item["record_key"])
            payload["_db_record_key"] = _text(item["record_key"])
            stored_mail_number = _text(payload.get("mail_number")) or _text(item["mail_number"])
            stored_detail_number = _text(payload.get("detail_number")) or _text(item["detail_number"])
            mail_number = stored_mail_number or _mail_number(_text(item["mail_key"]))
            detail_number = stored_detail_number or _detail_number(_text(item["record_key"]))
            payload["mail_number"] = mail_number
            payload["detail_number"] = detail_number
            if (
                not stored_mail_number
                or not stored_detail_number
            ):
                missing.append((mail_number, detail_number, json.dumps(payload, ensure_ascii=False), _text(item["record_key"])))
            result.append(payload)
        if missing:
            now = self._now()
            with self._lock, self._connect() as conn:
                for mail_number, detail_number, payload_json, record_key in missing:
                    conn.execute(
                        "UPDATE input_records SET mail_number=?, detail_number=?, payload_json=?, updated_at=? WHERE record_key=?",
                        (mail_number, detail_number, payload_json, now, record_key),
                    )
        return result

    def record_operation(
        self,
        action: str,
        *,
        record_id: str = "",
        mail_id: str = "",
        reason: str = "",
        result: Optional[Dict[str, Any]] = None,
        operation_date: Optional[str] = None,
    ) -> int:
        """把人工或系统动作写入按操作日期索引的日志。"""
        now = self._now()
        operation_date = _text(operation_date) or now[:10]
        with self._lock, self._connect() as conn:
            cursor = conn.execute(
                """
                INSERT INTO operation_log
                    (operation_date, action, record_id, mail_id, reason, result_json, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    operation_date, _text(action), _text(record_id), _text(mail_id), _text(reason),
                    json.dumps(_json_value(result or {}), ensure_ascii=False), now,
                ),
            )
            return int(cursor.lastrowid)

    def read_operations(self, operation_date: str = "") -> List[Dict[str, Any]]:
        """按操作日期读取留痕；传日期时只返回该日。"""
        operation_date = _text(operation_date)
        with self._lock, self._connect() as conn:
            if operation_date:
                rows = conn.execute(
                    "SELECT * FROM operation_log WHERE operation_date=? ORDER BY created_at DESC, id DESC",
                    (operation_date,),
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM operation_log ORDER BY operation_date DESC, created_at DESC, id DESC"
                ).fetchall()
        result: List[Dict[str, Any]] = []
        for row in rows:
            item = dict(row)
            try:
                item["result"] = json.loads(item.pop("result_json") or "{}")
            except (TypeError, ValueError):
                item["result"] = {}
            result.append(item)
        return result

    def summary(self) -> Dict[str, Any]:
        with self._lock, self._connect() as conn:
            counts = {
                dataset: int(
                    conn.execute("SELECT COUNT(*) FROM input_records WHERE dataset=?", (dataset,)).fetchone()[0]
                )
                for dataset in sorted(self.DATASETS)
            }
            operations = int(conn.execute("SELECT COUNT(*) FROM operation_log").fetchone()[0])
            latest = conn.execute("SELECT MAX(operation_date) FROM operation_log").fetchone()[0] or ""
        return {"path": str(self.path), "records": counts, "operations": operations, "latest_operation_date": latest}


def default_database_path(app_root: str | Path) -> Path:
    """供 GUI/独立脚本共用的正式数据库位置。"""
    return Path(app_root).resolve() / "storage" / "workbench.db"
