"""Anonymous cancellation regression tests; no real GUI, credentials or network."""
import ast
import os
import threading
from pathlib import Path
from types import SimpleNamespace, MethodType
from unittest.mock import Mock

import pytest

TREE = ast.parse((Path(__file__).parents[1] / 'gui.py').read_text(encoding='utf-8-sig'))
WORKER = next(n for n in TREE.body if isinstance(n, ast.ClassDef) and n.name == 'WorkerThread')


def worker_method(name, extra=None):
    node = next(n for n in WORKER.body if isinstance(n, ast.FunctionDef) and n.name == name)
    scope = dict(extra or {})
    exec(compile(ast.Module(body=[node], type_ignores=[]), 'gui.py', 'exec'), scope)
    return scope[name]


def worker(mode, stopped):
    obj = SimpleNamespace(mode=mode, _stop=stopped, _stage1_phase='邮件下载与附件解析',
                          _create_run_logger=Mock(return_value=Mock()),
                          _run_pipeline_stage1=Mock(return_value=None),
                          _detach_gui_log_handler=Mock(),
                          error_signal=Mock(), cancelled_signal=Mock(), finished_signal=Mock())
    for name in ('run', '_run_stage1', '_run_all'):
        setattr(obj, name, MethodType(worker_method(name), obj))
    return obj


@pytest.mark.parametrize('mode', ['stage1', 'all'])
def test_stop_before_unpack_emits_cancelled_not_error(mode):
    obj = worker(mode, True)
    obj.run()
    obj.cancelled_signal.emit.assert_called_once()
    obj.error_signal.emit.assert_not_called()
    obj.finished_signal.emit.assert_not_called()
    obj._detach_gui_log_handler.assert_called_once()
    assert '已停止' in obj._create_run_logger.return_value.info.call_args.args[0]


@pytest.mark.parametrize('mode', ['stage1', 'all'])
def test_unexpected_none_is_error_not_success_or_cancellation(mode):
    obj = worker(mode, False)
    obj.run()
    message = obj.error_signal.emit.call_args.args[0]
    assert '阶段一未返回有效结果' in message
    assert 'cannot unpack' not in message
    obj.cancelled_signal.emit.assert_not_called()
    obj.finished_signal.emit.assert_not_called()


@pytest.mark.parametrize('mode', ['stage1', 'all'])
def test_stop_after_valid_empty_result_does_not_write_outputs(mode):
    obj = worker(mode, True)
    obj._run_pipeline_stage1.return_value = ([], [])
    obj.run()
    obj.cancelled_signal.emit.assert_called_once()
    obj.error_signal.emit.assert_not_called()
    obj.finished_signal.emit.assert_not_called()


def test_invalid_project_table_is_explicit_input_error():
    run = worker_method('_run_pipeline_stage1', {'os': os, 'INTERNAL_EMAIL_CACHE_FILE': __file__})
    obj = SimpleNamespace(config={'reference_tables': {'agent_emails': 'agents.xlsx', 'project_names': 'projects.xlsx'}},
                          prefer_internal_email_path=False, internal_email_path=None,
                          agent_email_path=None, project_table_path=None,
                          _load_internal_emails=Mock(return_value=set()),
                          _load_agent_emails=Mock(return_value={}), _load_project_names=Mock(return_value=[]))
    with pytest.raises(ValueError, match='项目名称表'):
        run(obj, Mock())


def test_stop_records_source_and_phase():
    obj = SimpleNamespace(_stop=False, _stop_event=threading.Event(), _log=Mock(),
                          mode='stage1', _stage1_phase='邮件下载与附件解析')
    worker_method('stop')(obj, '匿名停止入口')
    assert obj._stop and obj._stop_event.is_set()
    assert '匿名停止入口' in obj._log.call_args.args[0]
    assert '邮件下载与附件解析' in obj._log.call_args.args[0]


def test_excel_error_names_the_archive_member(tmp_path, monkeypatch):
    import utils.attachment_parser as parser
    logger = Mock()
    monkeypatch.setattr(parser, '_get_logger', lambda: logger)
    path = tmp_path / 'invalid.xlsx'
    path.write_bytes(b'not a workbook')
    result = parser.parse_attachment(str(path), 'batch/nested/registration.xlsx')
    assert result['sheets'] == []
    assert any('batch/nested/registration.xlsx' in str(call) for call in logger.error.call_args_list)
