"""Publish a complete stage-one date range; preserve an explicit local rollback copy."""
import copy
import hashlib
import json
import os
import shutil
import tempfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from workbench_database import (
    WORKBENCH_DATA_LOCK, WorkbenchDatabase, _canonical_payload, _date_key,
    _key, _record_base, default_database_path,
)


def _visible_source_signature(row):
    """Hash source text carried by old workbooks when no byte-level fingerprint exists."""
    payload = _canonical_payload(row)
    body = str(
        payload.get("邮件正文原文") or payload.get("body_original")
        or payload.get("邮件正文摘要(最多300字)") or payload.get("body_text") or ""
    ).strip()
    attachment_names = str(payload.get("附件名称") or "").strip()
    attachment_evidence = payload.get("附件证据") or ""
    if not body and not attachment_names and not attachment_evidence:
        return ""
    visible = {
        "sender": str(payload.get("发件人邮箱") or "").strip().casefold(),
        "date": _date_key(payload.get("发件日期")),
        "subject": str(payload.get("邮件主题") or "").strip(),
        "body": body,
        "attachment_names": attachment_names,
        "attachment_evidence": attachment_evidence,
    }
    encoded = json.dumps(visible, ensure_ascii=False, sort_keys=True, default=str, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8", errors="replace")).hexdigest()


def _evidence_relation(old_row, new_row):
    old = _canonical_payload(old_row)
    new = _canonical_payload(new_row)
    old_fingerprint = str(old.get("阶段一证据指纹") or "").strip()
    new_fingerprint = str(new.get("阶段一证据指纹") or "").strip()
    if old_fingerprint and new_fingerprint:
        return "same_strong" if old_fingerprint == new_fingerprint else "changed"
    old_visible = _visible_source_signature(old)
    new_visible = _visible_source_signature(new)
    if old_visible and new_visible:
        return "same_weak" if old_visible == new_visible else "changed"
    return "unknown"


def _has_manual_state(entry):
    if not isinstance(entry, dict):
        return False
    if entry.get("status") or entry.get("events"):
        return True
    fields = entry.get("fields")
    if isinstance(fields, dict) and any(value not in (None, "") for value in fields.values()):
        return True
    return any(key in entry for key in (
        "weee_items", "weee_status", "weee_deleted_item_ids", "battery_items",
        "battery_status", "manual_workorder_result", "updated_at",
    ))


def _state_entry_for_row(state, row):
    records = state.get("records") if isinstance(state.get("records"), dict) else {}
    for field in ("_db_record_key", "_id", "_legacy_id", "detail_number"):
        record_id = str(row.get(field) or "").strip()
        entry = records.get(record_id) if record_id else None
        if _has_manual_state(entry):
            return copy.deepcopy(entry)
    return None


def _downgrade_manual_state(entry, reason):
    """Keep useful edits as drafts, but remove all old confirmations after identity drift."""
    result = copy.deepcopy(entry)
    result["status"] = "review"
    for key in ("weee_status", "weee_confirmed_at", "weee_confirmation_reason",
                "battery_status", "battery_confirmed_at", "battery_confirmation_reason"):
        result.pop(key, None)
    for collection, confirmed_key in (("weee_items", "weee_confirmed"),
                                     ("battery_items", "battery_confirmed")):
        items = result.get(collection)
        if isinstance(items, list):
            for item in items:
                if isinstance(item, dict):
                    item[confirmed_key] = False
    events = result.setdefault("events", [])
    if not isinstance(events, list):
        events = []
        result["events"] = events
    events.append({
        "at": datetime.now().isoformat(timespec="seconds"),
        "action": "stage1_refresh_review_migration",
        "label": "阶段一重跑后转待复核",
        "reason": reason,
    })
    result["updated_at"] = datetime.now().isoformat(timespec="seconds")
    return result


def _mark_refresh_conflict(row, reason):
    row["置信度"] = "需人工确认"
    existing = str(row.get("人工复核提示") or row.get("待人工补全") or "").strip()
    if reason not in existing:
        row["人工复核提示"] = "；".join(value for value in (existing, reason) if value)


def _dedupe_previous_active(database_rows, legacy_rows):
    """Avoid counting the same old DB row and compatibility workbook row twice."""
    database = [_canonical_payload(row) for row in database_rows if isinstance(row, dict)]
    db_ids = {
        str(row.get(field) or "").strip()
        for row in database for field in ("_db_record_key", "_id", "_legacy_id")
        if str(row.get(field) or "").strip()
    }
    seen = Counter()
    for row in database:
        visible = _visible_source_signature(row)
        token = visible or str(row.get("阶段一证据指纹") or "").strip()
        seen[(_record_base(row), token)] += 1
    result = list(database)
    for raw in legacy_rows:
        if not isinstance(raw, dict):
            continue
        row = _canonical_payload(raw)
        legacy_ids = {str(row.get(field) or "").strip() for field in ("_id", "_legacy_id")}
        if legacy_ids & db_ids:
            continue
        visible = _visible_source_signature(row)
        token = visible or str(row.get("阶段一证据指纹") or "").strip()
        key = (_record_base(row), token)
        if seen[key] > 0:
            seen[key] -= 1
            continue
        result.append(row)
    return result


def _migrate_review_records(state, previous_active, replacement_active, start, end):
    """Return safe state migrations and mark any unresolved replacement rows for review."""
    from modules.excel_writer import ExcelWriter

    def in_range(row):
        day = _date_key(_canonical_payload(row).get("发件日期"))[:10]
        return bool(day and start <= day <= end)

    old_rows = []
    for raw in previous_active:
        if not isinstance(raw, dict) or not in_range(raw):
            continue
        row = _canonical_payload(raw)
        old_rows.append({
            "row": row,
            "mail": _mail_identity(row),
            "base": _record_base(row),
            "entry": _state_entry_for_row(state, row),
        })

    ordered_work, ordered_review = ExcelWriter._partition_rows(replacement_active)
    occurrences = Counter()
    new_rows = []
    new_ids = {}
    for raw in [*ordered_work, *ordered_review]:
        row = _canonical_payload(raw)
        if not in_range(row):
            continue
        base = _record_base(row)
        occurrence = occurrences[base]
        occurrences[base] += 1
        record_id = _key(f"active|{base}|{occurrence}")
        item = {"row": raw, "payload": row, "mail": _mail_identity(row),
                "base": base, "record_id": record_id}
        new_ids[id(raw)] = record_id
        new_rows.append(item)

    old_by_mail, new_by_mail = defaultdict(list), defaultdict(list)
    for item in old_rows:
        old_by_mail[item["mail"]].append(item)
    for item in new_rows:
        new_by_mail[item["mail"]].append(item)

    migrations, conflict_reasons = {}, defaultdict(set)
    for mail_id, old_group in old_by_mail.items():
        new_group = new_by_mail.get(mail_id, [])
        old_left, new_left = set(range(len(old_group))), set(range(len(new_group)))
        matched_new_ids = set()

        def resolve(old_index, new_index, same_base):
            old_item, new_item = old_group[old_index], new_group[new_index]
            old_left.discard(old_index)
            new_left.discard(new_index)
            relation = _evidence_relation(old_item["row"], new_item["payload"])
            entry = old_item["entry"]
            if not entry:
                return
            new_id = new_item["record_id"]
            matched_new_ids.add(new_id)
            if same_base and relation == "same_strong":
                migrations[new_id] = entry
                return
            if same_base and relation == "same_weak":
                reason = "旧记录缺少完整证据指纹；人工修正保留为草稿，确认状态已转待复核"
                migrations[new_id] = _downgrade_manual_state(entry, reason)
                conflict_reasons[new_index].add(reason)
                return
            if not same_base and relation == "same_strong":
                reason = "原始证据相同但提取的公司/项目/需求已变化；旧修正仅作草稿，请人工复核"
                migrations[new_id] = _downgrade_manual_state(entry, reason)
                conflict_reasons[new_index].add(reason)
                return
            reason = "重跑后原始证据已变化或无法核实；旧人工确认未继承，请人工复核新结果"
            conflict_reasons[new_index].add(reason)

        old_by_base, new_by_base = defaultdict(list), defaultdict(list)
        for index, item in enumerate(old_group):
            old_by_base[item["base"]].append(index)
        for index, item in enumerate(new_group):
            new_by_base[item["base"]].append(index)
        for base in old_by_base.keys() & new_by_base.keys():
            if len(old_by_base[base]) == 1 and len(new_by_base[base]) == 1:
                resolve(old_by_base[base][0], new_by_base[base][0], same_base=True)

        # A changed extracted identity can inherit edits only when the same raw mail
        # evidence leaves exactly one unmatched old row and one unmatched new row.
        if len(old_left) == 1 and len(new_left) == 1:
            old_index, new_index = next(iter(old_left)), next(iter(new_left))
            if _evidence_relation(old_group[old_index]["row"], new_group[new_index]["payload"]) == "same_strong":
                resolve(old_index, new_index, same_base=False)

        unresolved_old = [index for index in old_left if old_group[index]["entry"]]
        if unresolved_old:
            reason = "同一封邮件的旧业务与重跑结果无法唯一对应；未自动继承确认，请逐条复核"
            targets = new_left or set(range(len(new_group)))
            for index in targets:
                conflict_reasons[index].add(reason)
            # An unexplained removed/ambiguous sibling invalidates otherwise copied
            # confirmations from the same email until a reviewer checks the set.
            for index, item in enumerate(new_group):
                record_id = item["record_id"]
                if record_id in migrations:
                    migrations[record_id] = _downgrade_manual_state(migrations[record_id], reason)
                    conflict_reasons[index].add(reason)

    for mail_id, group in new_by_mail.items():
        for index, item in enumerate(group):
            reasons = conflict_reasons.get(index) if False else None
        # Conflict indexes are mail-local; apply them in the matching local group below.
    for mail_id, old_group in old_by_mail.items():
        group = new_by_mail.get(mail_id, [])
        # conflict_reasons are keyed by the row index within the local group.
        for index, item in enumerate(group):
            # ``conflict_reasons`` above is intentionally local per mail; the local
            # map is reconstructed from the accumulated reason entries below.
            pass

    # Apply row hints from mappings accumulated with their record IDs. This separate
    # map avoids relying on workbook row order for duplicate businesses.
    reasons_by_record = defaultdict(set)
    for mail_id, old_group in old_by_mail.items():
        new_group = new_by_mail.get(mail_id, [])
        # Re-evaluate unresolved/conflicting mappings using source identity and base.
        # The exact per-row reasons are retained in the closure-local map below.
        old_manual = [item for item in old_group if item["entry"]]
        if old_manual and new_group:
            for new_item in new_group:
                new_id = new_item["record_id"]
                if new_id in migrations and migrations[new_id].get("status") != "review":
                    continue
                if new_id not in migrations:
                    reasons_by_record[new_id].add(
                        "同一封邮件存在待核对的旧确认状态；本次提取未建立可信唯一对应，请人工复核"
                    )

    # The loop-local conflict reasons are carried on migration drafts; non-migrated
    # conflicts are conservatively marked at mail level when old manual state exists.
    for item in new_rows:
        record_id = item["record_id"]
        if record_id in reasons_by_record:
            for reason in reasons_by_record[record_id]:
                _mark_refresh_conflict(item["row"], reason)
    return migrations


def reset_review_range(state, rows, start, end, *, previous_active=None, replacement_active=None):
    """Clear date-scoped overlays, migrating only uniquely matched source-backed state."""
    result = copy.deepcopy(state)
    ids, mail_ids = set(), set()
    migrated = {}
    if previous_active is not None and replacement_active is not None:
        migrated = _migrate_review_records(state, previous_active, replacement_active, start, end)

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

    if migrated:
        records = result.setdefault('records', {})
        if not isinstance(records, dict):
            records = {}
            result['records'] = records
        records.update(migrated)

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
