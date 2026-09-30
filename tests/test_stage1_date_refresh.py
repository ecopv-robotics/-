"""Anonymous date replacement regressions; no mailbox, LLM, or live business data."""
import json
from pathlib import Path

import pytest

from modules.excel_writer import ExcelWriter
from utils.stage1_refresh import publish_stage1, reset_review_range
from workbench_database import WorkbenchDatabase, default_database_path
from workbench_server import WorkbenchStore, _read_sheet


def row(company='Sample Alpha Ltd', day='2026-08-26'):
    return {'sender_email': 'sender@example.invalid', 'date': day + ' 10:00:00',
            'subject': 'Sample registration', '客户': company, '项目': '法国包装法',
            '需求': '注册', '代理': '示例代理', '代理命中数': 1, '置信度': 'high'}


def test_replace_both_queues_and_keep_other_dates(tmp_path):
    db = WorkbenchDatabase(tmp_path / 'check.db')
    db.ingest_rows([row('Fixture 7f43c514 Ltd'), row('Fixture 506454f2 Ltd', '2026-08-25')])
    db.ingest_rows([row('Fixture 9068251e Ltd')], 'filtered')
    with db.replace_stage1_range([row('Fixture 342f33b5 Ltd')], [], '2026-08-26', '2026-08-26') as report:
        assert report['removed'] == 2
    assert {r['客户公司名称'] for r in db.read_rows()} == {'Fixture 342f33b5 Ltd', 'Fixture 506454f2 Ltd'}
    assert db.read_rows('filtered') == []
    db.ingest_rows([row('Fixture 7f43c514 Ltd')])
    db.ingest_rows([row('Fixture 9068251e Ltd')], 'filtered')
    assert len(db.read_rows()) == 2 and not db.read_rows('filtered')
    with db._connect() as conn:
        assert conn.execute('SELECT count(*) FROM retired_stage1_records').fetchone()[0] == 2


def test_empty_success_clears_only_selected_day_and_survives_reopen(tmp_path):
    path = tmp_path / 'check.db'
    db = WorkbenchDatabase(path)
    db.ingest_rows([row(), row('Fixture 506454f2 Ltd', '2026-08-27')])
    with db.replace_stage1_range([], [], '2026-08-26', '2026-08-26'):
        pass
    reopened = WorkbenchDatabase(path)
    assert len(reopened.read_rows()) == 1
    assert reopened.unrefreshed_rows([row()]) == []


def test_preserve_unimported_other_day_excel_rows(tmp_path):
    output = tmp_path / 'output'
    ExcelWriter(str(output), categorized=True).write_stage1_outputs([row('Fixture 506454f2 Ltd', '2026-08-25')], [])
    publish_stage1(tmp_path, output, True, [row('Fixture ff637f3e Ltd')], [], '2026-08-26', '2026-08-26')
    db = WorkbenchDatabase(default_database_path(tmp_path))
    assert {r['客户公司名称'] for r in db.read_rows()} == {'Fixture 506454f2 Ltd', 'Fixture ff637f3e Ltd'}


def test_empty_publish_does_not_fall_back_to_old_excel(tmp_path):
    output, primary, review, db, state = setup_old(tmp_path)
    cache = tmp_path / 'cached.xlsx'
    cache.write_bytes(primary.read_bytes())
    publish_stage1(tmp_path, output, True, [], [], '2026-08-26', '2026-08-26')
    store = WorkbenchStore(str(cache), str(review), state_path=str(state))
    store.database = db
    assert {r['客户公司名称'] for r in store._raw_records()} == {'Fixture 506454f2 Ltd'}


def test_transaction_failure_restores_records_and_date_marker(tmp_path):
    db = WorkbenchDatabase(tmp_path / 'check.db')
    db.ingest_rows([row('Fixture 7167c001 Ltd')])
    before = db.read_rows()
    with pytest.raises(RuntimeError):
        with db.replace_stage1_range([row('Fixture 40db430e Ltd')], [], '2026-08-26', '2026-08-26'):
            raise RuntimeError('test rollback')
    assert db.read_rows() == before
    assert len(db.unrefreshed_rows([row()])) == 1


def test_reject_outside_row_before_deletion(tmp_path):
    db = WorkbenchDatabase(tmp_path / 'check.db')
    db.ingest_rows([row()])
    with pytest.raises(ValueError):
        with db.replace_stage1_range([row(day='2026-08-25')], [], '2026-08-26', '2026-08-26'):
            pass
    assert len(db.read_rows()) == 1


def test_overlapping_ranges_replace_instead_of_append(tmp_path):
    db = WorkbenchDatabase(tmp_path / 'check.db')
    with db.replace_stage1_range([row('Fixture ca8999fd Ltd', '2026-08-25'), row('Fixture 910a255b Ltd')], [], '2026-08-25', '2026-08-26'):
        pass
    with db.replace_stage1_range([row('Fixture 3e5443ab Ltd')], [], '2026-08-26', '2026-08-26'):
        pass
    assert {r['客户公司名称'] for r in db.read_rows()} == {'Fixture ca8999fd Ltd', 'Fixture 3e5443ab Ltd'}


def test_reset_manual_added_confirmed_deleted_and_routes_only_in_range():
    state = {'records': {'old': {'status': 'confirmed'}, 'manual': {'status': 'confirmed'},
                         'keep': {'status': 'confirmed'}},
             'deleted': {'old': {'reason': 'old'}},
             'added': {'target': {'mail': {'date': '2026-08-26'}, 'items': [{'id': 'manual'}]},
                       'other': {'mail': {'date': '2026-08-25'}, 'items': [{'id': 'keep'}]}},
             'mail_routes': {'target': {'route': 'filtered'}, 'other': {'route': 'review'}},
             'request_results': {'old-reply': {'record_id': 'old'}, 'keep-reply': {'record_id': 'keep'}},
             'change_log': [{'kind': 'audit'}]}
    updated = reset_review_range(state, [{**row(), '_id': 'old'}], '2026-08-26', '2026-08-26')
    assert updated['records'] == {'keep': {'status': 'confirmed'}}
    assert updated['deleted'] == {}
    assert set(updated['added']) == {'other'}
    assert set(updated['mail_routes']) == {'other'}
    assert set(updated['request_results']) == {'keep-reply'}
    assert updated['change_log'] == state['change_log']
    assert 'old' in state['records']  # input was not mutated


def setup_old(tmp_path):
    output = tmp_path / 'output'
    primary, review = ExcelWriter(str(output), categorized=True).write_stage1_outputs([row('Fixture 7f43c514 Ltd')], [])
    old = _read_sheet(Path(primary), '工单待查', '待查名单') + _read_sheet(Path(review), '漏单复查', '人工补全')
    db = WorkbenchDatabase(default_database_path(tmp_path))
    db.ingest_rows(old)
    db.ingest_rows([row('Fixture 506454f2 Ltd', '2026-08-25')])
    rid = next(r['_id'] for r in db.read_rows() if r['客户公司名称'] == 'Fixture 7f43c514 Ltd')
    state = tmp_path / 'storage' / 'workbench_review.json'
    state.write_text(json.dumps({'records': {rid: {'status': 'confirmed'}}}), encoding='utf-8')
    return output, Path(primary), Path(review), db, state


def test_publish_resets_confirmation_backs_up_and_old_cache_cannot_return(tmp_path):
    output, primary, review, db, state = setup_old(tmp_path)
    old_bytes = primary.read_bytes()
    cached = tmp_path / 'old-cache.xlsx'
    cached.write_bytes(old_bytes)
    publish_stage1(tmp_path, output, True, [row('Fixture 7f43c514 Ltd'), row('Fixture ff637f3e Ltd')], [], '2026-08-26', '2026-08-26')
    assert json.loads(state.read_text())['records'] == {}
    assert len(db.read_rows()) == 3  # two fresh + one other-day
    assert all(not r.get('_legacy_id') for r in db.read_rows() if r['发件日期'].startswith('2026-08-26'))
    backups = list((tmp_path / 'storage' / 'stage1_refresh_backups').glob('*/manifest.json'))
    assert len(backups) == 1
    manifest = json.loads(backups[0].read_text(encoding='utf-8'))
    assert Path(manifest['files'][str(primary)]).read_bytes() == old_bytes
    store = WorkbenchStore(str(cached), str(tmp_path / 'missing.xlsx'), state_path=str(state))
    store.database = db
    assert {r['客户公司名称'] for r in store._raw_records()} == {'Fixture 7f43c514 Ltd', 'Fixture ff637f3e Ltd', 'Fixture 506454f2 Ltd'}
    assert all(store._record(r, store._state())['status'] != 'confirmed' for r in store._raw_records())


def test_cancelled_run_never_replaces(tmp_path):
    output, primary, review, db, state = setup_old(tmp_path)
    before = (primary.read_bytes(), state.read_bytes(), db.read_rows())
    assert publish_stage1(tmp_path, output, True, [row('Fixture ff637f3e Ltd')], [], '2026-08-26', '2026-08-26', cancelled=lambda: True) is None
    assert (primary.read_bytes(), state.read_bytes(), db.read_rows()) == before


def test_file_publication_failure_rolls_back_db_files_state(tmp_path, monkeypatch):
    import utils.stage1_refresh as refresh
    output, primary, review, db, state = setup_old(tmp_path)
    before = (primary.read_bytes(), review.read_bytes(), state.read_bytes(), db.read_rows())
    real = refresh.os.replace
    def fail_on_state(source, target):
        if Path(target) == state:
            raise OSError('simulated locked state')
        return real(source, target)
    monkeypatch.setattr(refresh.os, 'replace', fail_on_state)
    with pytest.raises(OSError, match='simulated'):
        publish_stage1(tmp_path, output, True, [row('Fixture ff637f3e Ltd')], [], '2026-08-26', '2026-08-26')
    assert (primary.read_bytes(), review.read_bytes(), state.read_bytes(), db.read_rows()) == before
    assert db.unrefreshed_rows([row()])
