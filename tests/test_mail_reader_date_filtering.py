import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import modules.mail_reader as mail_reader_module
from modules.mail_reader import MailReader


def _message(uid, date_header=None):
    headers = [
        b"From: sender@example.test\r\n",
        b"To: receiver@example.test\r\n",
        f"Subject: message {uid}\r\n".encode(),
    ]
    if date_header is not None:
        headers.append(f"Date: {date_header}\r\n".encode())
    return b"".join(headers) + b"\r\nbody"


class _FakeConnection:
    def __init__(
        self,
        messages,
        internal_dates,
        internal_search_uids=None,
        sent_search_uids=None,
    ):
        self.messages = messages
        self.internal_dates = internal_dates
        self.internal_search_uids = list(
            messages if internal_search_uids is None else internal_search_uids
        )
        self.sent_search_uids = list(
            messages if sent_search_uids is None else sent_search_uids
        )
        self.uid_calls = []

    def response(self, key):
        return key, [b"77"]

    def uid(self, command, *args):
        self.uid_calls.append((command, args))
        if command == "SEARCH":
            if len(args) >= 2 and args[1] == "SINCE":
                result = self.internal_search_uids
            elif len(args) >= 2 and args[1] == "SENTSINCE":
                result = self.sent_search_uids
            else:
                raise AssertionError(f"unexpected IMAP date search: {args}")
            return "OK", [b" ".join(uid.encode() for uid in result)]

        if command != "FETCH":
            raise AssertionError(f"unexpected UID command: {command}")

        uid_arg, query = args
        uid_values = (
            uid_arg.decode().split(",") if isinstance(uid_arg, bytes)
            else str(uid_arg).split(",")
        )
        response = []
        for uid in uid_values:
            internal = self.internal_dates.get(uid)
            internal_clause = (
                f' INTERNALDATE "{internal}"' if internal else ""
            )
            if "BODY.PEEK[HEADER.FIELDS (DATE)]" in query:
                raw = self.messages[uid]
                header_payload = raw.split(b"\r\n\r\n", 1)[0] + b"\r\n\r\n"
                metadata = (
                    f"{uid} (UID {uid}{internal_clause} "
                    f"BODY[HEADER.FIELDS (DATE)] {{{len(header_payload)}}}"
                ).encode()
                response.append((metadata, header_payload))
            elif "RFC822" in query:
                response.append(
                    (f"{uid} (UID {uid}{internal_clause} RFC822 {{{len(self.messages[uid])}}}".encode(),
                     self.messages[uid])
                )
            else:
                raise AssertionError(f"unexpected FETCH query: {query}")
        return "OK", response

    def close(self):
        pass

    def logout(self):
        pass


class _Logger:
    def __init__(self):
        self.lines = []

    def info(self, message):
        self.lines.append(("info", message))

    def warning(self, message):
        self.lines.append(("warning", message))

    def error(self, message):
        self.lines.append(("error", message))


@pytest.mark.parametrize('stage', ['before_connect', 'download', 'parse'])
def test_cancel_stops_remaining_mails_and_closes_connection(tmp_path, stage):
    from unittest.mock import Mock
    from threading import Event
    stop = Event()
    messages = {str(i): _message(str(i), 'Tue, 25 Aug 2026 10:00:00 +0800') for i in (101, 102)}
    conn = _FakeConnection(messages, {uid: '25-Aug-2026 02:00:00 +0000' for uid in messages})
    conn.close, conn.logout = Mock(), Mock()
    logger = _Logger()
    reader = MailReader({'imap_server': 'unused', 'imap_port': 993, 'address': 'reader@example.test',
                         'password': 'unused', 'mailbox': 'INBOX', 'cache_enabled': False,
                         'cache_dir': str(tmp_path)}, logger)
    reader._connect = Mock(return_value=conn)
    progress = []
    def on_progress(current, total):
        progress.append((current, total))
        if stage == 'download' or (stage == 'parse' and len(progress) == 3):
            stop.set()
    if stage == 'before_connect':
        stop.set()
    result = reader.fetch_mails('2026-08-25', '2026-08-26', 'self@example.test',
                                progress_callback=on_progress, cancel_requested=stop.is_set)
    assert result == []
    assert any('邮件读取已停止' in text for level, text in logger.lines)
    if stage == 'before_connect':
        reader._connect.assert_not_called()
    else:
        conn.close.assert_called_once()
        conn.logout.assert_called_once()
    if stage == 'download':
        assert len([1 for cmd, args in conn.uid_calls if cmd == 'FETCH' and 'RFC822' in args[-1]]) == 1
    if stage == 'parse':
        assert len(progress) == 3  # Two downloads, one parsed mail; second mail never parsed.


def test_header_date_range_is_used_instead_of_internaldate_prefilter(tmp_path):
    messages = {
        # 用户日期在范围内，但邮件后来才进入收件箱：必须命中。
        "101": _message("101", "Tue, 04 Aug 2026 10:00:00 +0800"),
        # 无 Date 头时才使用 INTERNALDATE。
        "102": _message("102"),
        # Date 头明确在范围外，即使 INTERNALDATE 在范围内也不能命中。
        "103": _message("103", "Tue, 01 Sep 2026 10:00:00 +0800"),
        "104": _message("104", "not a valid date"),
        # 截止日为闭区间。
        "105": _message("105", "Wed, 26 Aug 2026 23:59:00 +0800"),
    }
    internal_dates = {
        "101": "01-Sep-2026 10:00:00 +0000",
        "102": "10-Aug-2026 10:00:00 +0000",
        "103": "10-Aug-2026 10:00:00 +0000",
        "104": "01-Sep-2026 10:00:00 +0000",
        "105": "01-Sep-2026 10:00:00 +0000",
    }
    conn = _FakeConnection(
        messages,
        internal_dates,
        internal_search_uids=["102", "103"],
        sent_search_uids=["101", "105"],
    )
    logger = _Logger()
    reader = MailReader({
        "imap_server": "unused", "imap_port": 993,
        "address": "reader@example.test", "password": "unused",
        "mailbox": "INBOX", "cache_dir": str(tmp_path),
    }, logger)
    reader._connect = lambda _ctx: conn

    mails = reader.fetch_mails("2026-08-04", "2026-08-26", "self@example.test")

    assert [mail["uid"] for mail in mails] == ["101", "102", "105"]
    assert mails[0]["date"].date().isoformat() == "2026-08-04"
    assert mails[0]["date_source"] == "邮件 Date 头"
    assert mails[1]["date"].date().isoformat() == "2026-08-10"
    assert mails[1]["date_source"].startswith("IMAP INTERNALDATE")
    assert reader.last_mailbox_total == 4  # 兼容属性现在表示范围候选数，不是全邮箱数
    assert reader.last_range_candidate_total == 4
    assert reader.last_search_total == 3
    assert reader.last_header_date_match_total == 2
    assert reader.last_internal_date_fallback_total == 1
    assert reader.last_unknown_date_total == 0

    search_calls = [args for command, args in conn.uid_calls if command == "SEARCH"]
    assert search_calls == [
        (None, "SINCE", "04-Aug-2026", "BEFORE", "27-Aug-2026"),
        (None, "SENTSINCE", "04-Aug-2026", "SENTBEFORE", "27-Aug-2026"),
    ]
    full_fetch_uids = [
        args[0].decode() for command, args in conn.uid_calls
        if command == "FETCH" and "RFC822" in args[1]
    ]
    assert full_fetch_uids == ["101", "102", "105"]
    log_text = "\n".join(message for _, message in logger.lines)
    assert "邮箱=INBOX" in log_text
    assert "IMAP日期范围搜索完成" in log_text
    assert "范围候选UID数=4" in log_text
    assert "Date头命中=2" in log_text
    assert "INTERNALDATE兜底命中=1" in log_text


def test_unknown_date_is_not_misreported_as_in_range(tmp_path):
    conn = _FakeConnection(
        {"201": _message("201", "broken")},
        {"201": None},
        internal_search_uids=["201"],
        sent_search_uids=[],
    )
    reader = MailReader({
        "imap_server": "unused", "imap_port": 993,
        "address": "reader@example.test", "password": "unused",
        "cache_dir": str(tmp_path),
    })
    reader._connect = lambda _ctx: conn

    mails = reader.fetch_mails("2026-08-04", "2026-08-26", "self@example.test")

    assert mails == []
    assert reader.last_mailbox_total == 1
    assert reader.last_range_candidate_total == 1
    assert reader.last_search_total == 0
    assert reader.last_unknown_date_total == 1
    assert not any(
        command == "FETCH" and "RFC822" in args[1]
        for command, args in conn.uid_calls
    )


def test_cached_iso_internaldate_is_used_for_missing_date_header(tmp_path):
    uid = "301"
    raw = _message(uid)
    cache_path = tmp_path / f"INBOX_uidv77_{uid}.eml"
    cache_path.write_bytes(raw)
    (tmp_path / f"INBOX_uidv77_{uid}.eml.meta.json").write_text(
        json.dumps({"internal_date": "2026-08-10T13:00:00+08:00"}),
        encoding="utf-8",
    )
    conn = _FakeConnection(
        {uid: raw},
        {uid: "10-Aug-2026 05:00:00 +0000"},
        internal_search_uids=[uid],
        sent_search_uids=[],
    )
    reader = MailReader({
        "imap_server": "unused", "imap_port": 993,
        "address": "reader@example.test", "password": "unused",
        "cache_dir": str(tmp_path),
    })
    reader._connect = lambda _ctx: conn

    mails = reader.fetch_mails("2026-08-04", "2026-08-26", "self@example.test")

    assert len(mails) == 1
    assert mails[0]["uid"] == uid
    assert mails[0]["date"].date().isoformat() == "2026-08-10"
    assert mails[0]["date_source"].startswith("IMAP INTERNALDATE")
    assert not any(command == "FETCH" for command, _ in conn.uid_calls)


def test_connect_logs_the_configured_mailbox_selection(monkeypatch):
    logger = _Logger()

    class _SelectableConnection:
        def __init__(self):
            self.selected = None

        def login(self, address, password):
            assert address == "reader@example.test"
            assert password == "unused"

        def select(self, mailbox, readonly=False):
            self.selected = (mailbox, readonly)
            return "OK", [b"42"]

        def logout(self):
            pass

    fake = _SelectableConnection()
    monkeypatch.setattr(
        mail_reader_module.imaplib,
        "IMAP4_SSL",
        lambda *args, **kwargs: fake,
    )
    reader = MailReader({
        "imap_server": "imap.example.test", "imap_port": 993,
        "address": "reader@example.test", "password": "unused",
        "mailbox": "INBOX",
    }, logger)

    assert reader._connect(object()) is fake
    assert fake.selected == ("INBOX", True)
    assert any("已选择邮箱: INBOX" in message and "42" in message
               for _, message in logger.lines)


def test_date_range_search_failure_is_reported_not_returned_as_empty_success(tmp_path):
    class _FailedSearchConnection(_FakeConnection):
        def uid(self, command, *args):
            self.uid_calls.append((command, args))
            if command == "SEARCH":
                return "NO", [b"date search not allowed"]
            raise AssertionError(f"unexpected UID command: {command}")

    conn = _FailedSearchConnection({}, {})
    logger = _Logger()
    reader = MailReader({
        "imap_server": "unused", "imap_port": 993,
        "address": "reader@example.test", "password": "unused",
        "mailbox": "INBOX", "cache_dir": str(tmp_path),
    }, logger)
    reader._connect = lambda _ctx: conn

    with pytest.raises(RuntimeError, match="UID 日期范围搜索失败"):
        reader.fetch_mails("2026-08-04", "2026-08-26", "self@example.test")
    assert any("邮箱=INBOX" in message for _, message in logger.lines)
    assert any("IMAP读取/连接失败" in message for _, message in logger.lines)
