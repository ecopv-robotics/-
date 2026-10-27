"""品牌抽取：WEEE 与电池法共用同一套规则（标签并集 + 值边界 + 多值拆分）。"""

from modules.brand_extract import extract_brand_values


def test_colon_form():
    assert extract_brand_values("品牌：YWNYT\n新增类目：第五类") == ["YWNYT"]


def test_insert_brand_without_colon():
    text = "深圳市琴和贸易有限公司德国WEEE，增加品牌Preciva，\n第五类，小型电子电气设备"
    assert extract_brand_values(text) == ["Preciva"]


def test_original_registered_brand_keeps_suffix():
    text = "原主体注册品牌：Shenzhenshenglongtaishiye Co., Ltd."
    assert extract_brand_values(text) == ["Shenzhenshenglongtaishiye Co., Ltd."]


def test_cancel_brand_with_count():
    text = "德国电池法注册号：81090045，注销1品牌：Loekneue，请尽快提交"
    assert extract_brand_values(text) == ["Loekneue"]


def test_two_subjects_glued_in_one_line():
    text = ("新增品牌 YBOEG 第五类HST011 郑州兰序网络科技有限公司 "
            "新增品牌 eroidn 第五类")
    assert extract_brand_values(text) == ["YBOEG", "eroidn"]


def test_brand_stops_before_battery_type():
    text = "品牌：ABC Co., Ltd. 电池类型：设备电池"
    assert extract_brand_values(text) == ["ABC Co., Ltd."]


def test_multiple_brands_split_by_count():
    assert extract_brand_values("取消2个品牌：Foo、Bar") == ["Foo", "Bar"]


def test_placeholder_is_not_brand():
    assert extract_brand_values("原订单OXH017-东莞市强翰科技有限公司-品牌无+新注册") == []
