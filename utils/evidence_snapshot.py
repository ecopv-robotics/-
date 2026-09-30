"""Keep Excel evidence snapshots valid without losing the workbook directory."""
from __future__ import annotations

import copy
import json


def encode_evidence_snapshot(payload: list[dict], max_chars: int = 30000) -> str:
    """Bound previews before serialization; never cut a serialized JSON string.

    Original files remain available through the attachment index. Large previews
    retain every sheet locator and favor rows containing extracted field evidence.
    Extraction itself always uses the full parsed attachments, not this preview.
    """
    def encode(value):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))

    encoded = encode(payload)
    if len(encoded) <= max_chars:
        return encoded
    compact = copy.deepcopy(payload)
    row_sets = []
    for attachment in compact:
        attachment["preview_truncated"] = True
        attachment["text"] = str(attachment.get("text") or "")[:240]
        for sheet in attachment.get("sheets") or []:
            rows = sheet.get("rows") or []
            sheet["preview_total_rows"] = len(rows)
            locators = {
                str(record.get("row_number"))
                for record in attachment.get("records") or []
                if record.get("sheet_name") == sheet.get("sheet_name")
                and (not sheet.get("member_path") or
                     record.get("attachment_name") == sheet.get("member_path"))
            }
            prioritized = sorted(enumerate(rows), key=lambda pair: (
                str(pair[1].get("row_number")) not in locators, pair[0]))
            row_sets.append((sheet, prioritized))

    def with_row_limit(limit):
        for sheet, prioritized in row_sets:
            sheet["rows"] = [row for _, row in sorted(prioritized[:limit])]
        return encode(compact)

    # Find the largest shared per-sheet preview budget. Every sheet stays listed.
    lower, upper = 0, max((len(rows) for _, rows in row_sets), default=0)
    best = with_row_limit(0)
    if len(best) <= max_chars:
        while lower <= upper:
            middle = (lower + upper) // 2
            candidate = with_row_limit(middle)
            if len(candidate) <= max_chars:
                best, lower = candidate, middle + 1
            else:
                upper = middle - 1
        return best

    # Very large structured lists: retain field values and locators, remove the
    # duplicate cell/raw-text copies already available in the original workbook.
    for attachment in compact:
        for record in attachment.get("records") or []:
            record.pop("cells", None)
            record.pop("raw_text", None)
    encoded = with_row_limit(0)
    if len(encoded) <= max_chars:
        return encoded
    for attachment in compact:
        attachment["record_count"] = len(attachment.get("records") or [])
        attachment["records"] = []
        attachment["text"] = "预览已精简，请打开原附件查看完整表格。"
    encoded = encode(compact)
    if len(encoded) <= max_chars:
        return encoded
    # An extreme sheet directory can exceed a cell by itself. Report this
    # explicitly and retain attachment-level entries, rather than corrupt JSON.
    for attachment in compact:
        attachment["sheet_count"] = len(attachment.get("sheets") or [])
        attachment["sheets"] = []
    encoded = encode(compact)
    if len(encoded) <= max_chars:
        return encoded
    return encode([{"preview_truncated": True, "attachment_count": len(payload),
                    "text": "附件目录超出预览容量，请从附件清单打开原文件。"}])
