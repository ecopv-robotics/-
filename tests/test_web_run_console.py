"""Anonymous isolated tests; never connect to mail, LLM, or work-order services."""
import io
import json
import secrets
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from types import SimpleNamespace
from urllib.error import HTTPError
from urllib.request import Request, urlopen

import pytest
import yaml
from openpyxl import Workbook

from web_run_controller import RunController, validate_workbook, DesktopWorkerAdapter
from workbench_server import WorkbenchServer


def workbook_bytes(role="stage2"):
    wb = Workbook()
    ws = wb.active
    ws.title = "工单待查" if role == "stage2" else "Sheet1"
    headers, values = {
        "stage2": (["代理", "客户公司名称", "标准化项目名称"], ["sample", "Sample LLC", "波兰包装法"]),
        "agent": (["代理", "简称", "邮箱"], ["sample", "sample", "sample@example.invalid"]),
        "internal": (["邮箱"], ["internal@example.invalid"]),
        "project": (["项目编号", "项目名称", "国家", "业务类型"], ["1", "波兰包装法", "波兰", "注册"]),
    }[role]
    ws.append(headers)
    ws.append(values)
    buffer = io.BytesIO()
    wb.save(buffer)
    wb.close()
    return buffer.getvalue()


class FakeWorker:
    def __init__(self, kwargs, callbacks):
        self.kwargs, self.callbacks = kwargs, callbacks
        self.started = False
        self.stopped = False

    def start(self):
        self.started = True

    def stop(self):
        self.stopped = True


@pytest.fixture
def controller(tmp_path):
    config = {"email": {"address": "mail@example.invalid", "password": secrets.token_hex(16)},
              "workorder": {"username": "sample", "password": secrets.token_hex(16)},
              "llm": {"api_key": secrets.token_hex(16)}, "output": {"dir": "output"}}
    (tmp_path / "config.yaml").write_text(yaml.safe_dump(config), encoding="utf-8")
    data = tmp_path / "data"
    data.mkdir()
    for role, name in (("agent", "agent_emails"), ("internal", "internal_email_cache"), ("project", "project_names")):
        (data / (name + ".xlsx")).write_bytes(workbook_bytes(role))
    return RunController(tmp_path, worker_factory=FakeWorker)


def start_payload():
    return {"mode": "stage1", "date_from": "2026-08-04", "date_to": "2026-08-26"}


def test_start_exact_closed_date_range_and_shared_worker(controller):
    controller.start(start_payload())
    worker = controller.worker
    assert worker.started
    assert worker.kwargs["date_from"].isoformat() == "2026-08-04T00:00:00"
    assert worker.kwargs["date_to"].isoformat() == "2026-08-26T00:00:00"
    assert worker.kwargs["prefer_internal_email_path"] is True
    assert Path(worker.kwargs["config"]["output"]["dir"]).is_absolute()
    assert controller.busy


def test_double_click_and_multi_tab_are_serialized(controller):
    def attempt():
        try:
            controller.start(start_payload())
            return True
        except ValueError:
            return False
    with ThreadPoolExecutor(max_workers=2) as pool:
        assert sum(pool.map(lambda _: attempt(), range(2))) == 1


def test_stop_waits_for_actual_completion(controller):
    controller.start(start_payload())
    controller.stop()
    assert controller.worker.stopped and controller.busy
    with pytest.raises(ValueError):
        controller.start(start_payload())
    controller.worker.callbacks["cancelled"]("stopped")
    assert controller.busy
    controller.worker.callbacks["done"]()
    assert not controller.busy and controller.status == "stopped"


def test_result_signal_does_not_unlock_before_thread_exits(controller):
    controller.start(start_payload())
    out = controller.root / "output"
    out.mkdir()
    a, b = out / "to_workorder_list.xlsx", out / "to_review_list.xlsx"
    a.write_bytes(workbook_bytes())
    b.write_bytes(workbook_bytes())
    controller.worker.callbacks["result"](str(a), str(b))
    assert controller.busy and controller.status == "finishing"
    controller.worker.callbacks["done"]()
    assert controller.status == "completed" and not controller.busy
    assert controller.download(controller.results[0]["token"]) == a


def test_errors_are_terminal_only_after_thread_exit(controller):
    controller.start(start_payload())
    controller.worker.callbacks["error"]("sample failure")
    assert controller.busy
    controller.worker.callbacks["done"]()
    assert not controller.busy and controller.status == "error"


@pytest.mark.parametrize("payload", [dict(mode="unknown"), dict(mode="stage1", date_from="bad"),
                                     dict(mode="stage1", date_from="2026-09-01", date_to="2026-08-01")])
def test_invalid_start_cannot_create_worker(controller, payload):
    with pytest.raises(ValueError):
        controller.start(payload)
    assert not controller.busy and controller.worker is None


def test_snapshot_and_logs_do_not_contain_credentials(controller):
    sensitive = [controller.config["email"]["password"], controller.config["workorder"]["password"], controller.config["llm"]["api_key"]]
    controller.log(" ".join(sensitive))
    snapshot = json.dumps(controller.snapshot())
    assert all(value not in snapshot for value in sensitive)
    assert "[已隐藏]" in controller.snapshot()["logs"][0]["text"]


def test_log_cursor_does_not_skip_bursts(controller):
    for index in range(450):
        controller.log(str(index))
    first = controller.snapshot()
    second = controller.snapshot(first["cursor"])
    assert len(first["logs"]) == 300 and len(second["logs"]) == 150
    assert second["cursor"] == 450
    assert not controller.snapshot(450)["logs"]


def test_blank_password_keeps_existing_and_session_has_no_secrets(controller):
    original = controller.config["email"]["password"]
    controller.save_settings({"email_address": "new@example.invalid", "email_password": ""})
    assert controller.config["email"]["password"] == original
    controller.start(start_payload())
    saved = controller.session_path.read_text(encoding="utf-8")
    assert original not in saved


@pytest.mark.parametrize("role", ["agent", "internal", "project", "stage2"])
def test_upload_each_role_preserves_original_and_survives_restart(controller, role):
    originals = {p: p.read_bytes() for p in (controller.root / "data").glob("*")}
    controller.upload(role, "source.xlsx", workbook_bytes(role))
    assert controller.imports[role].parent.name == "web_imports"
    assert all(p.read_bytes() == before for p, before in originals.items())
    restored = RunController(controller.root, FakeWorker)
    assert restored.imports[role] == controller.imports[role]


def test_upload_rejects_invalid_file_and_paths_do_not_escape(controller):
    with pytest.raises(ValueError):
        controller.upload("stage2", "invalid.xlsx", b"not a workbook")
    controller.upload("stage2", "../../input.xlsx", workbook_bytes())
    assert controller.imports["stage2"].is_relative_to(controller.root / "storage/web_imports")
    with pytest.raises(ValueError):
        controller.download("../../config.yaml")


def test_stage2_uses_selected_file_and_force_flag(controller):
    uploaded = controller.upload("stage2", "reviewed.xlsx", workbook_bytes())
    controller.start({"mode": "stage2", "input_token": uploaded["file"]["token"], "force_live_query": True})
    args = controller.worker.kwargs
    assert args["date_from"] is None and args["date_to"] is None
    assert args["config"]["workorder"]["force_live_query"]
    assert args["stage2_input_path"] == str(controller.imports["stage2"])


def test_stage2_requires_correct_sheet(controller):
    with pytest.raises(ValueError, match="工单待查"):
        validate_workbook(workbook_bytes("project"), "stage2")


def test_cache_selection_only_rebinds_sources(controller):
    folder = controller.root / "output/stage1_email/cache/2026-08-04_to_2026-08-26"
    folder.mkdir(parents=True)
    for name in ("to_workorder_list.xlsx", "to_review_list.xlsx"):
        (folder / name).write_bytes(workbook_bytes())
    controller.store = SimpleNamespace(_lock=threading.RLock(), state_path="unchanged-review-state")
    controller.refresh_catalog()
    controller.use_cache(controller.catalog["caches"][0]["token"])
    assert controller.store.primary_path.parent == folder
    assert controller.store.state_path == "unchanged-review-state"
    assert controller.session["date_to"] == "2026-08-26"
    assert controller.worker is None


@pytest.fixture
def http_server(controller):
    server = WorkbenchServer(test_mode=True, run_controller=controller, port=0)
    url = server.start(open_browser=False).rstrip("/")
    yield controller, server, url
    server.stop()


def test_http_origin_token_html_and_download(http_server):
    controller, _, url = http_server
    with urlopen(url + "/run") as response:
        assert b"2026.09.30-r7" in response.read()
    with urlopen(url + "/api/run/state") as response:
        state = json.load(response)
    assert state["csrf"] == controller.csrf
    for headers in ({"Origin": "https://untrusted.invalid"}, {"Origin": "null"}, {"Host": "untrusted.invalid"}):
        with pytest.raises(HTTPError) as exc:
            urlopen(Request(url + "/api/run/state", headers=headers))
        assert exc.value.code == 403
    with pytest.raises(HTTPError) as exc:
        urlopen(Request(url + "/api/run/start", data=b"{}", headers={"Content-Type": "application/json"}))
    assert exc.value.code == 403
    headers = {"Content-Type": "application/json", "Origin": url, "X-Run-Token": state["csrf"]}
    with urlopen(Request(url + "/api/run/start", data=json.dumps(start_payload()).encode(), headers=headers)) as response:
        assert json.load(response)["ok"]
    assert controller.busy
    with pytest.raises(HTTPError):
        urlopen(Request(url + "/api/run/shutdown", data=b"{}", headers=headers))
    controller.stop()
    controller.worker.callbacks["cancelled"]("stopped")
    controller.worker.callbacks["done"]()


def test_web_console_cannot_bind_public_interface(controller):
    with pytest.raises(ValueError, match="127.0.0.1"):
        WorkbenchServer(host="0.0.0.0", run_controller=controller)


def test_shutdown_prevents_late_start(controller):
    controller.closing = True
    with pytest.raises(ValueError, match="退出"):
        controller.start(start_payload())


def test_zero_port_is_ephemeral_and_live_port_cannot_be_shared(controller):
    first = WorkbenchServer(test_mode=True, port=0, run_controller=controller)
    second = None
    try:
        first.start(open_browser=False)
        port = first.httpd.server_address[1]
        assert first.port == 0
        second = WorkbenchServer(test_mode=True, port=port)
        second.start(open_browser=False)
        assert second.httpd.server_address[1] != port
    finally:
        if second:
            second.stop()
        first.stop()


def test_actual_qt_adapter_forwards_signals_without_running_network(monkeypatch):
    from PyQt5.QtCore import QCoreApplication, QThread, pyqtSignal
    import gui
    app = QCoreApplication.instance() or QCoreApplication([])
    class Worker(QThread):
        log_signal = pyqtSignal(str)
        progress_signal = pyqtSignal(int, int)
        finished_signal = pyqtSignal(str, str)
        error_signal = pyqtSignal(str)
        cancelled_signal = pyqtSignal(str)
        def __init__(self, **kwargs):
            super().__init__()
        def run(self):
            self.log_signal.emit("sample")
            self.progress_signal.emit(1, 1)
            self.finished_signal.emit("first", "second")
        def stop(self):
            pass
    monkeypatch.setattr(gui, "WorkerThread", Worker)
    events, done = [], threading.Event()
    callbacks = {name: lambda *args, n=name: events.append((n, args)) for name in ("log", "progress", "result", "error", "cancelled")}
    callbacks["done"] = done.set
    adapter = DesktopWorkerAdapter({}, callbacks)
    adapter.start()
    assert done.wait(3)
    assert not adapter.worker.isRunning()
    assert [x[0] for x in events] == ["log", "progress", "result"]
