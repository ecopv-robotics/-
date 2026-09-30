"""Independent manual battery review; never reuse WEEE equipment categories."""
import hashlib
import re


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
        if item_id != delete_id:
            result.append(item)
    if confirm_id and confirm_id not in seen or delete_id and delete_id not in seen:
        raise ValueError("电池品牌/品类项目不存在")
    return result
