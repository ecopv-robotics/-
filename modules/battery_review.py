"""German battery evidence extraction and independent item-by-item review."""
import hashlib
import json
import re


BATTERY_CLASSES = ("设备电池", "轻型交通工具电池", "工业电池", "启动电池", "汽车电池")
_BATTERY_CONTEXT = re.compile(r"电池|batter(?:y|ies|ie|ien)|BattG|BattVO|BattDG", re.I)
_WEEE_CONTEXT = re.compile(r"WEEE|ElektroG|电子电气", re.I)
_BRAND_LABEL = re.compile(r"(?:新增品牌|品牌名称|品牌名|品牌|brand(?:\s+name)?|marke)\s*[:：]?\s*([^,，;；|\n\r+＋]{1,90})", re.I)
_TYPE_LABEL = re.compile(r"(?:电池(?:类型|类别|品类|分类)|原始品类|注册类别|battery\s*(?:type|category)|batteriekategorie)\s*[:：]?\s*([^,，;；|\n\r+＋]{1,90})", re.I)
_PLACEHOLDER = re.compile(r"^(?:品牌|品牌名称|电池类型|电池类别|示例|例如|例|待填写|请选择|必填|选填|brand|category|type)(?:\s*[:：].*)?$", re.I)
_CLASS_PATTERNS = (
    ("汽车电池", re.compile(r"电动汽车电池|电动车(?:辆)?电池|电动汽车动力电池|Elektrofahrzeugbatterie|(?:Anlage[\s_-]*)?FZG|electric vehicle batter", re.I)),
    ("轻型交通工具电池", re.compile(r"轻型交通工具电池|轻型交通电池|轻型车辆电池|LV[-\s]?Batter|Leichte\s+Verkehrsmittel|light means of transport batter|(?:Anlage[\s_-]*)?LV(?:[-_.\s]|$)", re.I)),
    ("启动电池", re.compile(r"汽车启动电池|启动电池|起动电池|Starterbatterie|starting.*battery|(?:Anlage[\s_-]*)?STB(?:[-_.\s]|$)", re.I)),
    ("工业电池", re.compile(r"工业电池|Industriebatterie|industrial batter|(?:NB[-_\s]?)?BESS|stationäre[sn]?\s+Batterie|储能电池", re.I)),
    ("设备电池", re.compile(r"设备电池|便携式电池|Gerätebatterie|portable batter", re.I)),
)


def classify_battery_category(value):
    """Return a class only when the same evidence names exactly one category."""
    text = str(value or "").strip()
    hits = [name for name, pattern in _CLASS_PATTERNS if pattern.search(text)]
    return hits[0] if len(hits) == 1 else ""


def _clean(value):
    text = re.sub(r"&nbsp;|\s+", " ", str(value or ""), flags=re.I).strip(" \t:：,，;；|+＋")
    return "" if not text or _PLACEHOLDER.fullmatch(text) else text[:1000]


def _norm(value):
    return re.sub(r"[\s\W_]+", "", str(value or "").casefold())


def _company_matches(record, company, filename):
    if not company:
        return True
    owner = _clean(record.get("customer") or record.get("company"))
    if owner:
        return _norm(owner) == _norm(company) or _norm(owner) in _norm(company) or _norm(company) in _norm(owner)
    # A shared table without an owner must not be copied to every company.
    return bool(_norm(company) and _norm(company) in _norm(filename))


def _sheet_battery_records(attachment, company):
    """Read selected values below battery section labels in EPR worksheets."""
    filename = _clean(attachment.get("filename"))
    for sheet in attachment.get("sheets") or []:
        if not isinstance(sheet, dict):
            continue
        sheet_name = _clean(sheet.get("sheet_name"))
        source_name = _clean(sheet.get("member_path")) or filename
        rows = sheet.get("preview_rows") or sheet.get("rows") or []
        normalized = []
        for row in rows:
            if isinstance(row, dict):
                normalized.append((row.get("row_number") or row.get("index") or "?", [str(cell or "").strip() for cell in row.get("cells") or row.get("values") or []]))
            elif isinstance(row, (list, tuple)):
                normalized.append(("?", [str(cell or "").strip() for cell in row]))
        if not normalized:
            continue
        whole_sheet = " ".join(cell for _, cells in normalized for cell in cells if cell)
        if company and _norm(company) not in _norm(source_name + whole_sheet):
            continue
        active = bool(_BATTERY_CONTEXT.search(sheet_name) and not _WEEE_CONTEXT.search(sheet_name))
        headers = None
        for index, (row_number, cells) in enumerate(normalized):
            row_text = " ".join(cells)
            if re.search(r"WEEE产品信息|包装(?:法)?产品信息|WEEE\s*(?:product|information)", row_text, re.I):
                active, headers = False, None
                continue
            if re.search(r"电池(?:法)?产品信息|电池类型信息|battery\s*(?:product\s*)?information|Batterie.*Information", row_text, re.I):
                active, headers = True, None
                continue
            if not active:
                continue
            brand_cols = [column for column, value in enumerate(cells) if re.fullmatch(r"\*?\s*(?:品牌|品牌名称|brand|brand name|marke)\s*", value, re.I)]
            category_cols = [column for column, value in enumerate(cells) if re.fullmatch(r"\*?\s*(?:类别|品类|电池类型|电池类别|battery type|battery category|batteriekategorie)\s*", value, re.I)]
            if brand_cols or category_cols:
                headers = (brand_cols, category_cols)
                continue
            if not headers:
                continue
            brands = [_clean(cells[column]) for column in headers[0] if column < len(cells)]
            categories = [_clean(cells[column]) for column in headers[1] if column < len(cells)]
            brand = next((value for value in brands if value), "")
            category = next((value for value in categories if value), "")
            if not (brand or category) or re.search(r"^(?:示例|例如|例\s*[:：]|example|e\.g\.)", row_text, re.I):
                continue
            yield {"brand": brand, "category": category, "sheet_name": sheet_name,
                   "row_number": row_number, "raw_text": row_text,
                   "attachment_name": source_name, "record_type": "battery_section"}


def extract_battery_items(subject="", body="", attachments=None, project="", company=""):
    """Extract battery brand/type from the selected business and attachment evidence.

    Contract codes identify candidates, not a human confirmation. The generic OfH
    contract names all five categories, so it never establishes a specific one.
    """
    if not is_german_battery(project):
        return {"enabled": False, "items": [], "status": "not_applicable"}
    attachments = attachments or []
    items = []
    contract_classes = {}
    for attachment in attachments:
        if not isinstance(attachment, dict):
            continue
        filename = _clean(attachment.get("filename"))
        member_names = [filename]
        member_names.extend(str(sheet.get("member_path") or "") for sheet in attachment.get("sheets") or [] if isinstance(sheet, dict))
        member_names.extend(str(record.get("attachment_name") or "") for record in attachment.get("structured_records") or [] if isinstance(record, dict))
        member_names.extend(re.findall(r"(?m)\[([^\]\r\n]{1,500}\.(?:pdf|docx|xlsx|xls|csv|txt))\]", str(attachment.get("text_content") or ""), re.I))
        for member_name in dict.fromkeys(member_names):
            lower_name = member_name.casefold()
            if re.search(r"(?:vertrag|合同|anlage|grs)", lower_name, re.I) and not re.search(r"(?:^|[/\\_\s-])ofh(?:[/\\_\s.-]|$)", lower_name, re.I):
                category_class = classify_battery_category(member_name)
                if category_class:
                    contract_classes.setdefault(category_class, member_name)
        records = list(attachment.get("structured_records") or [])
        records.extend(_sheet_battery_records(attachment, company))
        for record in records:
            if not isinstance(record, dict):
                continue
            record_filename = _clean(record.get("attachment_name")) or filename
            if record.get("record_type") != "battery_section" and not _company_matches(record, company, record_filename):
                continue
            context = " ".join(str(record.get(key) or "") for key in ("project", "business", "sheet_name", "record_type")) + " " + record_filename
            if _WEEE_CONTEXT.search(context) and not _BATTERY_CONTEXT.search(context):
                continue
            if not _BATTERY_CONTEXT.search(context):
                continue
            brand = _clean(record.get("brand"))
            category = _clean(record.get("category"))
            if not (brand or category):
                continue
            source = f"附件表格：{record_filename} / {_clean(record.get('sheet_name')) or '工作表'} 第{record.get('row_number') or '?'}行"
            items.append({"brand": brand, "category": category, "category_class": classify_battery_category(category),
                          "source": source, "evidence": _clean(record.get("raw_text")) or " | ".join(map(str, record.get("cells") or [])),
                          "confidence": "high" if category else "medium"})

    # Only battery-specific lines can contribute mail fields in a mixed WEEE mail.
    lines = [line.strip() for line in re.split(r"[\r\n]+", str(subject or "") + "\n" + str(body or "")) if line.strip()]
    for line in lines:
        if company and _norm(company) not in _norm(line):
            continue
        if not _BATTERY_CONTEXT.search(line) or (_WEEE_CONTEXT.search(line) and not classify_battery_category(line)):
            continue
        brands = [_clean(match.group(1)) for match in _BRAND_LABEL.finditer(line)]
        categories = [_clean(match.group(1)) for match in _TYPE_LABEL.finditer(line)]
        brands = [value for value in brands if value]
        categories = [value for value in categories if value]
        if not categories:
            category_class = classify_battery_category(line)
            if category_class:
                categories = [category_class]
        for index in range(max(len(brands), len(categories))):
            brand = brands[index] if index < len(brands) else ""
            category = categories[index] if index < len(categories) else ""
            if brand or category:
                items.append({"brand": brand, "category": category, "category_class": classify_battery_category(category),
                              "source": "邮件正文/标题", "evidence": line[:1000], "confidence": "medium"})


    merged = {}
    for item in items:
        key = (_norm(item["brand"]), _norm(item["category"]))
        if not any(key):
            continue
        old = merged.setdefault(key, item)
        if old is not item and item["source"] not in old["source"]:
            old["source"] += "；" + item["source"]
    result = list(merged.values())
    for category_class, filename in contract_classes.items():
        if not any(item.get("category_class") == category_class for item in result):
            result.append({"brand": "", "category": category_class, "category_class": category_class,
                           "source": "附件合同名称", "evidence": filename, "confidence": "medium"})
    if attachments and not contract_classes:
        for item in result:
            if not item.get("category_class"):
                item["category_class"] = "设备电池"
                item["confidence"] = "low"
                item["source"] += "；未见专用类别合同"
    if not result and attachments:
        # Absence of a dedicated contract is only the user's weak equipment-battery heuristic.
        result.append({"brand": "", "category": "", "category_class": "设备电池",
                       "source": "未见专用类别合同", "evidence": "仅作设备电池候选，待人工核对", "confidence": "low"})
    if not result:
        result.append({"brand": "", "category": "", "category_class": "",
                       "source": "邮件/附件未找到明确品类", "evidence": "请人工核对电池类别", "confidence": "low"})
    for index, item in enumerate(result):
        item["item_id"] = "BAT-" + hashlib.sha256(repr((company, index, item["brand"], item["category"], item["source"])).encode("utf-8")).hexdigest()[:20]
        item["battery_confirmed"] = False
    return {"enabled": True, "items": result, "status": "pending"}


def is_german_battery(program):
    return re.sub(r"\s+", "", str(program or "")) in {"德国电池", "德国电池法"}


def battery_review(program, saved, source_items, status):
    items = saved.get("battery_items", source_items)
    items = items if isinstance(items, list) else []
    required = "battery_items" in saved or bool(items)
    confirmed = bool(items) and all(item.get("battery_confirmed") for item in items)
    if not required:
        confirmed = status == "confirmed"  # Preserve previously confirmed business.
    if status in {"review", "returned", "needs_info"}:
        confirmed = False
    return {"enabled": is_german_battery(program), "items": items,
            "requires_review": required, "confirmed": confirmed,
            "status": "confirmed" if confirmed else "pending"}


def normalize_battery_items(raw_items, previous_items=(), confirm_id="", delete_id=""):
    if not isinstance(raw_items, list) or len(raw_items) > 200:
        raise ValueError("电池品牌/品类格式不正确，一次最多200项")
    previous = {item.get("item_id"): item for item in previous_items if isinstance(item, dict)}
    result, seen = [], set()
    keys = ("brand", "category", "category_class")
    for index, raw in enumerate(raw_items):
        if not isinstance(raw, dict):
            raise ValueError("电池品牌/品类项目格式不正确")
        item = {key: str(raw.get(key) or "").strip() for key in keys}
        if any(len(value) > 1000 for value in item.values()):
            raise ValueError("电池品牌/品类字段过长")
        item_id = str(raw.get("item_id") or "BAT-" + hashlib.sha256(
            repr((index, item)).encode("utf-8")).hexdigest()[:20])
        if item_id in seen:
            raise ValueError("电池品牌/品类项目编号重复")
        seen.add(item_id)
        old = previous.get(item_id, {})
        complete = all(item.values())
        # Only an explicit confirmation or unchanged saved confirmation is trusted.
        item["battery_confirmed"] = bool(complete and (item_id == confirm_id or (
            old.get("battery_confirmed") and all(old.get(key) == item[key] for key in keys))))
        item["item_id"] = item_id
        for key in ("source", "evidence", "confidence"):
            value = str(raw.get(key) or old.get(key) or "").strip()
            if value:
                item[key] = value[:1000]
        if item_id != delete_id:
            result.append(item)
    if confirm_id and confirm_id not in seen or delete_id and delete_id not in seen:
        raise ValueError("电池品牌/品类项目不存在")
    return result


def compare_battery_categories(items, workorders):
    """Compare confirmed battery classes with the matching order's category tab."""
    if isinstance(items, str):
        try:
            items = json.loads(items)
        except (TypeError, ValueError):
            items = []
    mail_items = [item for item in (items or []) if isinstance(item, dict)]
    order_categories, order_brands = [], []
    for order in workorders or []:
        if not isinstance(order, dict):
            continue
        category_rows = order.get("品类明细") or []
        if not isinstance(category_rows, list):
            category_rows = []
        for row in category_rows:
            if not isinstance(row, dict):
                continue
            for key, value in row.items():
                label = str(key).casefold()
                if any(word in label for word in ("品类", "类别", "分类", "电池类型", "category", "battery type", "batteriekategorie")):
                    category = _clean(value)
                    if category:
                        order_categories.append(category)
                if any(word in label for word in ("品牌", "brand", "marke")):
                    brand = _clean(value)
                    if brand:
                        order_brands.append(brand)
    checked = []
    for item in mail_items:
        brand = _clean(item.get("brand"))
        category = _clean(item.get("category"))
        category_class = _clean(item.get("category_class")) or classify_battery_category(category)
        if not category or category_class not in BATTERY_CLASSES:
            status, reason = "pending", "邮件电池品类尚未确认"
        elif not order_categories:
            status, reason = "pending", "工单未返回品类明细"
        elif not any(_norm(category) == _norm(value) or classify_battery_category(value) == category_class
                     for value in order_categories):
            status, reason = "missing", "工单品类明细未找到对应电池类别"
        elif brand and not order_brands:
            status, reason = "pending", "工单未返回品牌字段"
        elif brand and not any(_norm(brand) == _norm(value) for value in order_brands):
            status, reason = "pending", "电池类别已找到，品牌仍需人工核对"
        else:
            status, reason = "matched", "工单电池类别与品牌已找到"
        checked.append({"brand": brand, "category": category, "category_class": category_class,
                        "status": status, "reason": reason})
    overall = ("missing" if any(item["status"] == "missing" for item in checked)
               else "pending" if not checked or any(item["status"] == "pending" for item in checked)
               else "matched")
    return {"enabled": bool(mail_items), "status": overall, "items": checked,
            "workorder_categories": list(dict.fromkeys(order_categories)),
            "workorder_brands": list(dict.fromkeys(order_brands))}
