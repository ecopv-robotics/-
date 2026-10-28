"""德国 WEEE 一行只留一个品类：多品类必须拆成多条明细。

客户口径：一个主体确实有多个品类时要拆行，公司名称重复没关系；但不能让人
在同一张卡片里看到 N 段重复的「品牌 / 德国分类」。
"""

import json

from modules.field_extractor import FieldExtractor
from modules.weee_category_audit import extract_weee_items


def _row(items, **extra):
    row = {"客户公司名称": "四川伐火网络科技有限公司", "标准化项目名称": "德国WEEE",
           "需求": "注册", "德国WEEE品类明细": json.dumps(items, ensure_ascii=False)}
    row.update(extra)
    return row


def test_multi_category_row_is_split_into_one_row_per_category():
    items = [
        {"brand": "SEZEAIE", "category": "小型设备，单边尺寸小于50cm", "category_class": "5"},
        {"brand": "SEZEAIE", "category": "大型设备，单边尺寸大于50cm", "category_class": "4"},
    ]
    rows = FieldExtractor._split_weee_category_rows([_row(items)])
    assert len(rows) == 2
    assert [row["品类拆分行"] for row in rows] == ["1/2", "2/2"]
    for row in rows:
        single = json.loads(row["德国WEEE品类明细"])
        assert len(single) == 1
    assert [json.loads(row["德国WEEE品类明细"])[0]["category_class"] for row in rows] == ["5", "4"]
    # 公司名称按客户要求重复，其他字段原样保留
    assert {row["客户公司名称"] for row in rows} == {"四川伐火网络科技有限公司"}
    assert {row["标准化项目名称"] for row in rows} == {"德国WEEE"}


def test_single_category_row_is_untouched():
    rows = FieldExtractor._split_weee_category_rows(
        [_row([{"brand": "ophimu", "category": "小型设备", "category_class": "5"}])]
    )
    assert len(rows) == 1
    assert "品类拆分行" not in rows[0]


def test_rows_without_weee_items_are_untouched():
    for raw in ("", None, "[]"):
        row = {"客户公司名称": "A", "德国WEEE品类明细": raw}
        assert FieldExtractor._split_weee_category_rows([row]) == [row]


def test_battery_class_is_not_treated_as_weee_category():
    """「设备电池」是电池法类别，不能混进德国 WEEE 品类（实测 四川伐火）。"""
    result = extract_weee_items(
        subject="德国WEEE 注册 四川伐火网络科技有限公司",
        body="四川伐火网络科技有限公司\n品牌 SEZEAIE 品类 设备电池",
        attachments=[], project="德国WEEE",
    )
    categories = [item.get("category") for item in result.get("items") or []]
    assert "设备电池" not in categories


def test_battery_row_is_split_per_category_too():
    """电池法同样一卡一品类（客户口径不区分 WEEE / 电池法）。"""
    row = {"客户公司名称": "香港郡為電子商務有限公司", "标准化项目名称": "德国电池法",
           "需求": "注册", "德国电池法品类明细": json.dumps([
               {"brand": "Amicool", "category": "便携式电池 内置锂电池", "category_class": "设备电池"},
               {"brand": "43G", "category": "便携式电池", "category_class": "设备电池"},
           ], ensure_ascii=False)}
    rows = FieldExtractor._split_weee_category_rows([row])
    assert len(rows) == 2
    assert [json.loads(r["德国电池法品类明细"])[0]["brand"] for r in rows] == ["Amicool", "43G"]
    assert [r["品类拆分行"] for r in rows] == ["1/2", "2/2"]
    assert all(r["客户公司名称"] == "香港郡為電子商務有限公司" for r in rows)


def test_battery_class_filler_row_is_dropped_when_real_categories_exist():
    """有真实电池品类时，「空品牌 + 设备电池」填充行不该占一条。"""
    from modules.battery_review import extract_battery_items
    attachment = {
        "filename": "香港酷康電子商務有限公司＋epr注册申请表.pdf",
        "sheets": [],
        "structured_records": [{"record_type": "weee_catalog", "customer": "",
                                "brand": "Amicool", "category": "便携式电池 内置锂电池",
                                "attachment_name": "香港酷康電子商務有限公司＋epr注册申请表.pdf",
                                "sheet_name": "Sheet1", "row_number": 9}],
    }
    result = extract_battery_items(
        subject="德国电池法 香港酷康電子商務有限公司",
        body="香港酷康電子商務有限公司\n电池品类：设备电池",
        attachments=[attachment], project="德国电池法", company="香港酷康電子商務有限公司",
    )
    items = result["items"]
    assert len(items) == 1, items
    assert items[0]["brand"] == "Amicool"
