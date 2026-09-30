import json

import pytest
from openpyxl import Workbook, load_workbook

from modules.battery_review import battery_review, normalize_battery_items
from workbench_server import WorkbenchStore, _detail_confirmation_ready


def item(item_id="B1", brand="Example"):
    return {"item_id": item_id, "brand": brand, "category": "Original battery type",
            "category_class": "Manual battery class"}


def make_store(tmp_path):
    path = tmp_path / "primary.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["邮件编号", "明细编号", "发件人邮箱", "发件日期", "邮件主题",
                  "代理", "客户公司名称", "标准化项目名称", "需求", "置信度",
                  "德国WEEE专项", "德国WEEE品类明细"])
    for index, project in enumerate(["德国WEEE", "德国电池法"]):
        sheet.append(["MIXED", f"D{index}", "fixture@example.test", "2026-08-26 09:00:00",
                      "Anonymous mixed projects", "Example Agent", "Example LLC", project,
                      "注册", "high", "是" if index == 0 else "",
                      json.dumps([item("W1")]) if index == 0 else ""])
    workbook.save(path)
    workbook.close()
    return WorkbenchStore(primary_path=str(path), review_path=str(tmp_path / "review.xlsx"),
                          filtered_path=str(tmp_path / "filtered.xlsx"),
                          state_path=str(tmp_path / "state.json"), test_mode=True)


def battery_detail(store):
    return next(d for m in store.snapshot()["mails"] for d in m["details"]
                if d["fields"]["program"] == "德国电池法")


def test_battery_draft_cannot_forge_confirmation():
    assert not normalize_battery_items([{**item(), "battery_confirmed": True}])[0]["battery_confirmed"]


def test_battery_edit_invalidates_confirmation_and_empty_list_is_pending():
    saved = normalize_battery_items([item()], confirm_id="B1")
    assert normalize_battery_items([item()], saved)[0]["battery_confirmed"]
    assert not normalize_battery_items([item(brand="Changed")], saved)[0]["battery_confirmed"]
    assert not battery_review("德国电池法", {"battery_items": []}, [], "confirmed")["confirmed"]


def test_battery_preserves_legacy_confirmation_but_return_blocks_it():
    assert battery_review("德国电池法", {}, [], "confirmed")["confirmed"]
    saved = {"battery_items": normalize_battery_items([item()], confirm_id="B1")}
    assert not battery_review("德国电池法", saved, [], "returned")["confirmed"]


@pytest.mark.parametrize("items,kwargs", [([item(), item()], {}), ([item()], {"confirm_id": "missing"}),
                                          (None, {}), (["invalid"], {})])
def test_invalid_battery_payload_rejected(items, kwargs):
    with pytest.raises(ValueError):
        normalize_battery_items(items, **kwargs)


def test_mixed_mail_battery_actions_do_not_change_weee(tmp_path):
    store = make_store(tmp_path)
    before = store.snapshot()["mails"][0]
    battery = battery_detail(store)
    weee_before = next(d["weee"] for d in before["details"] if d["weee"]["enabled"])
    payload = {"record_id": battery["id"], "mail_id": before["id"],
               "battery_items": [item(), item("B2", "Second")], "battery_item_id": "B1"}
    store.action({**payload, "action": "save_battery_draft"})
    assert not _detail_confirmation_ready(battery_detail(store))
    store.action({**payload, "action": "confirm_battery_item"})
    first = battery_detail(store)
    assert first["battery"]["items"][0]["battery_confirmed"]
    assert not _detail_confirmation_ready(first)
    store.action({**payload, "action": "confirm_battery_item", "battery_item_id": "B2"})
    assert _detail_confirmation_ready(battery_detail(store))
    weee_after = next(d["weee"] for d in store.snapshot()["mails"][0]["details"] if d["weee"]["enabled"])
    assert weee_after == weee_before
    reopened = WorkbenchStore(primary_path=store.primary_path, review_path=store.review_path,
                               filtered_path=store.filtered_path, state_path=store.state_path, test_mode=True)
    assert _detail_confirmation_ready(battery_detail(reopened))
    store.action({**payload, "action": "delete_battery_item", "battery_item_id": "B2"})
    current = battery_detail(store)
    assert len(current["battery"]["items"]) == 1
    store.action({**payload, "action": "save_battery_draft", "battery_items": [item(brand="Edited")]})
    assert not _detail_confirmation_ready(battery_detail(store))


def test_battery_action_rejects_wrong_project(tmp_path):
    store = make_store(tmp_path)
    weee = next(d for d in store.snapshot()["mails"][0]["details"] if d["weee"]["enabled"])
    with pytest.raises(ValueError, match="德国电池法"):
        store.action({"action": "confirm_battery_item", "record_id": weee["id"],
                      "battery_items": [item()], "battery_item_id": "B1"})


def test_battery_last_item_delete_is_not_completion(tmp_path):
    store = make_store(tmp_path)
    detail = battery_detail(store)
    payload = {"record_id": detail["id"], "battery_items": [item()], "battery_item_id": "B1"}
    store.action({**payload, "action": "confirm_battery_item"})
    assert _detail_confirmation_ready(battery_detail(store))
    store.action({**payload, "action": "delete_battery_item"})
    assert battery_detail(store)["battery"]["items"] == []
    assert not _detail_confirmation_ready(battery_detail(store))


def test_battery_draft_does_not_create_confirmation_event(tmp_path):
    store = make_store(tmp_path)
    detail = battery_detail(store)
    store.action({"action": "save_battery_draft", "record_id": detail["id"], "battery_items": [item()]})
    assert not battery_detail(store)["events"]


def test_battery_export_retains_independent_manual_fields(tmp_path, monkeypatch):
    import workbench_server
    monkeypatch.setattr(workbench_server, "MANUAL_OUTPUT", tmp_path / "export")
    store = make_store(tmp_path)
    detail = battery_detail(store)
    store.action({"action": "confirm_battery_item", "record_id": detail["id"],
                  "battery_items": [item()], "battery_item_id": "B1"})
    workbook = load_workbook(store.export(), read_only=True)
    rows = list(workbook.active.values)
    assert len(rows) == 2  # Unconfirmed WEEE business must not be exported.
    row = dict(zip(rows[0], rows[1]))
    assert row["标准化项目名称"] == "德国电池法"
    assert json.loads(row["德国电池法品类明细"])[0]["category_class"] == "Manual battery class"
    workbook.close()
