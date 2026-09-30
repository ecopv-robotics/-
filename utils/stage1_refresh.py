"""Publish a complete stage-one date range; preserve an explicit local rollback copy."""
import copy
import hashlib
import json
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

from workbench_database import (
    WORKBENCH_DATA_LOCK, WorkbenchDatabase, _canonical_payload, _date_key,
    _key, _record_base, default_database_path,
)


def reset_review_range(state, rows, start, end):
    """Remove current manual overlays only; historical audit events remain readable."""
    result = copy.deepcopy(state)
    ids, mail_ids = set(), set()

    def in_range(value):
        day = _date_key(value)[:10]
        return bool(day and start <= day <= end)

    occurrences = {}
    for raw in rows:
        row = _canonical_payload(raw)
        if not in_range(row.get('发件日期')):
            continue
        for name in ('_id', '_db_record_key', '_legacy_id', 'detail_number'):
            if row.get(name):
                ids.add(str(row[name]))
        base = _record_base(row)
        occurrence = occurrences.get(base, 0)
        occurrences[base] = occurrence + 1
        for dataset in ('active', 'filtered'):
            ids.add(_key(f'{dataset}|{base}|{occurrence}'))
        identity = '|'.join(str(row.get(k) or '') for k in ('发件人邮箱', '发件日期', '邮件主题'))
        mail_ids.add(hashlib.sha1(identity.encode('utf-8', errors='ignore')).hexdigest()[:16])

    added = result.get('added', {})
    for mail_id, bucket in list(added.items()):
        if isinstance(bucket, dict) and in_range((bucket.get('mail') or {}).get('date')):
            mail_ids.add(mail_id)
            ids.update(str(item.get('id')) for item in bucket.get('items', []) if isinstance(item, dict))
            del added[mail_id]
    for mail_id, item in result.get('deleted_mails', {}).items():
        if isinstance(item, dict) and in_range(item.get('date')):
            mail_ids.add(mail_id)
    mail_ids.update('filtered:' + rid for rid in ids)
    for field in ('records', 'deleted', 'mail_routes', 'deleted_mails'):
        mapping = result.get(field)
        if not isinstance(mapping, dict):
            continue
        for rid, entry in list(mapping.items()):
            if rid in ids or rid in mail_ids or (isinstance(entry, dict) and entry.get('mail_id') in mail_ids):
                del mapping[rid]

    # Stale idempotent replies must not claim the new extraction is already confirmed.
    targets = ids | mail_ids
    def references(value):
        if isinstance(value, str):
            return value in targets
        if isinstance(value, dict):
            return any(references(v) for v in value.values())
        if isinstance(value, list):
            return any(references(v) for v in value)
        return False
    replies = result.get('request_results', {})
    if isinstance(replies, dict):
        result['request_results'] = {k: v for k, v in replies.items() if not references(v)}
    return result


def publish_stage1(app_root, output_dir, categorized, active, filtered, date_from, date_to,
                   logger=None, cancelled=lambda: False):
    """Stage files before acquiring the review lock; failed publication rolls back."""
    from modules.excel_writer import ExcelWriter
    from workbench_server import _read_sheet, _save_json
    root, output = Path(app_root), Path(output_dir)
    start, end = _date_key(date_from)[:10], _date_key(date_to)[:10]
    # Validate before any official output is touched.
    if not start or not end or start > end:
        raise ValueError('阶段一覆盖需要有效的邮件日期范围')
    state_path = root / 'storage' / 'workbench_review.json'
    backup_root = root / 'storage' / 'stage1_refresh_backups'
    backup_root.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='stage1-publish-', dir=backup_root) as staging:
        writer = ExcelWriter(staging, logger, categorized=categorized)
        primary, review = writer.write_stage1_outputs(active, filtered)
        filter_path = Path(primary).parent / 'filtered_mail_record.xlsx'
        prepared = _read_sheet(Path(primary), '工单待查', '待查名单')
        prepared.extend(_read_sheet(Path(review), '漏单复查', '人工补全'))
        prepared_filtered = _read_sheet(filter_path, '过滤日志', '过滤日志')
        if len(prepared) != len(active) or len(prepared_filtered) != len(filtered):
            raise RuntimeError('新结果文件校验未通过，未覆盖原结果')
        files = [(p, output / p.relative_to(staging)) for p in Path(staging).rglob('*.xlsx')]
        primary_target = output / Path(primary).relative_to(staging)
        review_target = output / Path(review).relative_to(staging)
        if cancelled():
            return None
        with WORKBENCH_DATA_LOCK:
            db = WorkbenchDatabase(default_database_path(root))
            old_rows = db.read_rows('active') + db.read_rows('filtered')
            legacy_active, legacy_filtered = [], []
            # Capture legacy row-number review IDs before replacing the old files.
            for path, sheet, source in (
                (primary_target, '工单待查', '待查名单'),
                (review_target, '漏单复查', '人工补全'),
                (primary_target.parent / 'filtered_mail_record.xlsx', '过滤日志', '过滤日志'),
            ):
                legacy = _read_sheet(path, sheet, source)
                old_rows.extend(legacy)
                (legacy_filtered if source == '过滤日志' else legacy_active).extend(legacy)
            state = json.loads(state_path.read_text(encoding='utf-8')) if state_path.exists() else {}
            if not isinstance(state, dict):
                raise ValueError('人工复核记录异常，已取消覆盖')
            updated_state = reset_review_range(state, old_rows + prepared + prepared_filtered, start, end)
            if cancelled():
                return None
            backup = Path(tempfile.mkdtemp(prefix=f'{start}_to_{end}-', dir=backup_root))
            originals = {}
            for index, target in enumerate([state_path] + [target for _, target in files]):
                saved = backup / f'{index:03d}-{target.name}'
                if target.exists():
                    shutil.copy2(target, saved)
                    originals[target] = saved
                else:
                    originals[target] = None
            manifest = {'date_from': start, 'date_to': end,
                        'created_at': datetime.now().isoformat(),
                        'files': {str(k): str(v) if v else None for k, v in originals.items()}}
            _save_json(backup / 'manifest.json', manifest)
            touched = []
            try:
                with db.replace_stage1_range(prepared, prepared_filtered, start, end, str(primary_target),
                                             legacy_active, legacy_filtered) as report:
                    for source, target in files:
                        target.parent.mkdir(parents=True, exist_ok=True)
                        os.replace(source, target)
                        touched.append(target)
                    _save_json(state_path, updated_state)
                    touched.append(state_path)
            except Exception:
                for target in reversed(touched):
                    saved = originals[target]
                    if saved is not None:
                        shutil.copy2(saved, target)
                    elif target.exists():
                        target.unlink()  # Only a newly-created exact publication target.
                raise
            manifest.update(report)
            _save_json(backup / 'manifest.json', manifest)
            if logger:
                logger.info(f'阶段一按日期覆盖完成: {start} ~ {end}；旧明细 {report["removed"]} 条已归档，'
                            f'新询单 {report["active"]} 条、过滤 {report["filtered"]} 条；人工记录已重置；备份: {backup}')
            return str(primary_target), str(review_target)
