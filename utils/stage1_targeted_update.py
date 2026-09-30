"""Replace selected stage-one results without moving unrelated review identities.

Caller owns backups, concurrency exclusion, and the final atomic file replacement.
This module never touches a database, review state, or a network service.
"""
from copy import copy

from workbench_database import _key, _mail_identity, _record_base


def mail_key(row):
    return _key(_mail_identity(row))


def replace_selected_rows(workbook, replacements):
    """Patch the first sheet in memory; preserve other rows/sheets and row numbers.

    replacements maps stable mail keys to complete rows keyed by existing headers.
    Empty replacement lists remove only those selected rows' contents. Surplus
    rows are appended, never inserted. Exact business identities reuse their slots.
    """
    sheet = workbook.worksheets[0]
    headers = [cell.value for cell in sheet[1]]
    if not headers or len(set(headers)) != len(headers) or None in headers:
        raise ValueError("Results must have unique, nonempty headers")
    required = {"发件人邮箱", "发件日期", "邮件主题", "客户公司名称", "标准化项目名称", "需求"}
    if not required.issubset(headers):
        raise ValueError("Not a stage-one results sheet")

    # Validate every replacement before changing any cells.
    prepared = {}
    for key, incoming in replacements.items():
        rows = [dict(row) for row in incoming]
        identities = [_record_base(row) for row in rows]
        if len(identities) != len(set(identities)):
            raise ValueError("Duplicate replacement business identity")
        for row in rows:
            if mail_key(row) != key:
                raise ValueError("Replacement belongs to a different mail")
            if set(row) - set(headers):
                raise ValueError("Replacement has columns absent from the source")
            if any(isinstance(value, str) and len(value) > 32767 for value in row.values()):
                raise ValueError("Replacement exceeds the Excel cell limit")
        prepared[key] = rows

    slots = {key: [] for key in prepared}
    for cells in sheet.iter_rows(min_row=2):
        old = dict(zip(headers, (cell.value for cell in cells)))
        key = mail_key(old)
        if key in slots:
            slots[key].append((cells[0].row, old))

    result = {}
    for key, rows in prepared.items():
        old_slots = slots[key]
        available = {number: old for number, old in old_slots}
        assignments = []
        remaining = []
        for row in rows:
            number = next((n for n, old in available.items()
                           if _record_base(old) == _record_base(row)), None)
            if number is None:
                remaining.append(row)
            else:
                assignments.append((number, row))
                del available[number]
        for row in remaining:
            if available:
                number = next(iter(available))
                del available[number]
            else:
                number = max([sheet.max_row] + [n for n, _ in assignments]) + 1
                template = old_slots[0][0] if old_slots else (2 if sheet.max_row >= 2 else None)
                if template:
                    for col in range(1, len(headers) + 1):
                        sheet.cell(number, col)._style = copy(sheet.cell(template, col)._style)
                    sheet.row_dimensions[number].height = sheet.row_dimensions[template].height
            assignments.append((number, row))
        for number in available:
            for cell in sheet[number]:
                cell.value = None
        for number, row in assignments:
            for col, header in enumerate(headers, 1):
                cell = sheet.cell(number, col)
                value = row.get(header)
                cell.value = value
                # Mail content is data, never a newly injected formula.
                if isinstance(value, str):
                    cell.data_type = "s"
        result[key] = {"before": len(old_slots), "after": len(rows),
                       "rows": [number for number, _ in assignments]}
    return result
