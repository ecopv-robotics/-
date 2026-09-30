"""Exercise shutdown without opening the real GUI or loading credentials."""
import ast
import os
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from workbench_server import WorkbenchServer, WorkbenchStore


GUI_TREE = ast.parse((Path(__file__).parents[1] / "gui.py").read_text(encoding="utf-8-sig"))


def gui_function(name, globals_, method=False):
    nodes = next(node for node in GUI_TREE.body if isinstance(node, ast.ClassDef)
                 and node.name == "MainWindow").body if method else GUI_TREE.body
    found = [node for node in nodes if isinstance(node, ast.FunctionDef) and node.name == name]
    assert len(found) == 1, "A later definition must not override the safe shutdown"
    module = ast.Module(body=found, type_ignores=[])
    exec(compile(module, "gui.py", "exec"), globals_)
    return globals_[name]


def closing(running=False, answer=0, save=True):
    order = []
    messages = SimpleNamespace(Yes=1, No=0, question=Mock(return_value=answer), warning=Mock())
    timers = SimpleNamespace(singleShot=Mock())
    close = gui_function("closeEvent", {"QMessageBox": messages, "QTimer": timers}, method=True)
    window = SimpleNamespace(
        worker=SimpleNamespace(isRunning=Mock(return_value=running), stop=Mock()),
        save_session=Mock(side_effect=lambda: order.append("save") or save),
        workbench_server=SimpleNamespace(stop=Mock(side_effect=lambda: order.append("stop"))),
        log=Mock(), close=Mock(),
    )
    event = SimpleNamespace(accept=Mock(), ignore=Mock())
    return close, window, event, messages, timers, order


def test_idle_close_saves_before_stopping_http():
    close, window, event, _, _, order = closing()
    close(window, event)
    assert order == ["save", "stop"]
    event.accept.assert_called_once()


def test_cancel_running_close_keeps_task_and_http():
    close, window, event, messages, timers, order = closing(running=True)
    close(window, event)
    assert messages.question.call_args.args[-1] == messages.No
    event.ignore.assert_called_once()
    window.worker.stop.assert_not_called()
    timers.singleShot.assert_not_called()
    assert order == []


def test_running_close_waits_for_actual_thread_exit():
    close, window, event, messages, timers, order = closing(running=True, answer=1)
    close(window, event)
    close(window, event)
    window.worker.stop.assert_called_once()
    messages.question.assert_called_once()
    assert timers.singleShot.call_count == 2
    event.accept.assert_not_called()
    assert order == []
    window.worker.isRunning.return_value = False
    close(window, event)
    assert order == ["save", "stop"]
    event.accept.assert_called_once()


def test_failed_save_cannot_close_or_stop_http():
    close, window, event, messages, _, order = closing(save=False)
    close(window, event)
    assert order == ["save"]
    event.accept.assert_not_called()
    messages.warning.assert_called_once()


def test_pending_history_keeps_window_open_for_retry():
    close, window, event, messages, _, _ = closing()
    window.workbench_server.stop.side_effect = RuntimeError("history pending")
    close(window, event)
    event.accept.assert_not_called()
    messages.warning.assert_called_once()


@pytest.mark.parametrize("stdout,expected", [("", []), ("101\n202\n", [101, 202])])
def test_legacy_process_detection_is_read_only(stdout, expected):
    runner = Mock(return_value=SimpleNamespace(returncode=0, stdout=stdout))
    fake_os = SimpleNamespace(name="nt", path=os.path, getpid=lambda: 1234)
    detect = gui_function("_find_running_instances", {
        "os": fake_os, "APP_DIR": "C:/sample-app", "subprocess": SimpleNamespace(run=runner),
    })
    assert detect() == expected
    command = runner.call_args.args[0][-1]
    assert "Stop-Process" not in command and "taskkill" not in command
    assert "python.exe" in command and "pythonw.exe" in command
    assert "Get-CimInstance" in command


def test_legacy_detection_failure_is_not_silently_ignored():
    detect = gui_function("_find_running_instances", {
        "os": SimpleNamespace(name="nt", path=os.path, getpid=lambda: 1234),
        "APP_DIR": "C:/sample-app",
        "subprocess": SimpleNamespace(run=Mock(return_value=SimpleNamespace(returncode=1, stdout=""))),
    })
    with pytest.raises(RuntimeError, match="无法检查"):
        detect()


def test_history_wait_is_bounded_and_refuses_incomplete_shutdown():
    store = object.__new__(WorkbenchStore)
    thread = Mock()
    thread.is_alive.return_value = True
    store._history_sync_thread = thread
    with pytest.raises(RuntimeError, match="历史记录仍在保存"):
        store.wait_background_tasks(timeout=0.01)
    thread.join.assert_called_once_with(0.01)
    thread.is_alive.return_value = False
    store.wait_background_tasks(timeout=0.01)


def test_server_stop_waits_for_history_before_releasing_instance():
    server = object.__new__(WorkbenchServer)
    order = []
    server.httpd = SimpleNamespace(shutdown=lambda: order.append("http"),
                                   server_close=lambda: order.append("close"))
    server.store = SimpleNamespace(wait_background_tasks=lambda: order.append("history"))
    server.thread = object()
    server.stop()
    assert order == ["http", "close", "history"]
    assert server.httpd is None and server.thread is None
