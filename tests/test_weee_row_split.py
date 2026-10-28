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
