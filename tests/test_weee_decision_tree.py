"""WEEE 类别判定：按注册产品信息表的判断图（1→2→3→4→6→5）走。"""

from modules.weee_category_audit import _record_size_cm, classify_weee_product, parse_size_cm


def _cls(text, size=None):
    return classify_weee_product(text, size_cm=size)["category_class"]


def test_exchange_equipment_wins_over_size():
    assert _cls("空调", 120) == "1"
    assert _cls("冰箱", 200) == "1"


def test_screen_and_lamp_before_size_check():
    assert _cls("显示器", 70) == "2"
    assert _cls("LED灯", 80) == "3"


def test_size_over_50cm_is_large_equipment():
    assert _cls("backpack", 60) == "4"
    assert _cls("backpack", 50.1) == "4"


def test_size_at_or_below_50cm_is_small_equipment():
    assert _cls("backpack", 50) == "5"
    assert _cls("backpack", 12) == "5"


def test_telecom_without_size_is_class_six():
    assert _cls("手机") == "6"
    assert _cls("路由器", 20) == "6"


def test_missing_size_and_no_product_evidence_stays_pending():
    result = classify_weee_product("打印机")
    assert result["status"] == "ambiguous"
    assert result["category_class"] == ""


def test_explicit_class_number_wins_over_tree():
    assert _cls("第五类——小型非光伏设备") == "5"
    assert _cls("电气与电子设备废料-大型设备(光伏面板除外)") == "4"


def test_size_units_are_parsed_to_cm():
    assert parse_size_cm("60cm") == 60
    assert parse_size_cm("1.2m") == 120
    assert parse_size_cm("500mm") == 50
    assert parse_size_cm("50 厘米") == 50
    assert parse_size_cm("") == 0


def test_record_size_column_is_read():
    record = {"电器单边最大尺寸/cm": "55", "品牌名称": "X", "产品名称": "backpack"}
    assert _record_size_cm(record) == 55
    record_text = {"raw_text": "品牌 X | 产品 backpack | 最大尺寸 1.2m"}
    assert _record_size_cm(record_text) == 120
