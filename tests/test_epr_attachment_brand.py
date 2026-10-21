"""EPR 附件表格的品牌/品类提取回归（真实申请表结构，已脱敏）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from utils import attachment_parser as parser  # noqa: E402


def test_battery_type_header_is_recognized_as_category():
    """电池法申请表用「电池类型/电池类别」作为品类列，不能被当成非业务表头。"""
    assert parser._xlsx_header_role("电池类型 type（必填）") == "category"
    assert parser._xlsx_header_role("电池类别") == "category"
    assert parser._xlsx_header_role("对应品牌\nMarke") == "brand"


def test_estimated_volume_column_does_not_disqualify_header_row():
    """表头行里的「预估销售量」是正常业务列名，不能整行否决。

    真实申请表把 WEEE/电池/包装三个区块纵向堆叠在同一张 sheet，
    表头行常同时带「预估销售量」；旧逻辑把该词当成“表单信息区”特征，
    整行跳过，导致品牌列从未被识别。
    """
    rows = [
        ["公司中文名称", "产品类别", "产品名称", "品牌", "预估销售量"],
        ["示例甲贸易有限公司", "小型设备", "示例产品A", "Qynera", "5T"],
    ]
    records = parser._xlsx_structured_records(rows, "Sheet1")
    assert len(records) == 1
    assert records[0]["brand"] == "Qynera"
    assert records[0]["category"] == "小型设备"


def test_multi_section_application_extracts_brand_per_section():
    """一张 sheet 内 WEEE 与电池两个区块都要提取品牌，不能只读第一个区块。"""
    rows = [
        ["EPR合规注册申请表"],
        ["公司英文名称", "公司英文名值", "产品类别", "产品名称", "对应的品牌logo（必填）",
         "单边最大尺寸cm", "预估销售量"],
        ["公司中文名称", "示例甲贸易有限公司", "小型设备", "示例产品A", "Qynera", "34.6", "5T"],
        ["公司中文地址", "示例地址"],
        [],
        ["电池法信息填写"],
        ["", "", "电池类型 type（必填）", "设备电池成分", "对应品牌\nMarke",
         "单个电池重量", "预计全年销售产品含电池总量"],
        ["", "", "设备电池", "锂", "Cnkeeo", "32.9g", "1000"],
    ]
    records = parser._xlsx_structured_records(rows, "Sheet1")
    brands = [record["brand"] for record in records if record["brand"]]
    assert "Qynera" in brands
    assert "Cnkeeo" in brands
    # 输出里不允许出现表头文字或相邻列的值。
    for brand in brands:
        for noise in ("品牌", "必填", "logo", "单个电池重量", "预估销售量"):
            assert noise not in brand


def test_section_header_row_is_not_stored_as_business_data():
    """第二个区块的表头行不能作为上一个区块的数据记录。"""
    rows = [
        ["公司英文名称", "产品类别", "品牌", "预估销售量"],
        ["示例甲贸易有限公司", "小型设备", "Qynera", "5T"],
        ["电池法信息填写"],
        ["", "电池类型 type（必填）", "对应品牌\nMarke", "单个电池重量"],
        ["", "设备电池", "Cnkeeo", "32.9g"],
    ]
    records = parser._xlsx_structured_records(rows, "Sheet1")
    for record in records:
        assert record["brand"] in {"Qynera", "Cnkeeo"}


def test_catalog_record_with_brand_and_category_reaches_battery_items():
    """类别写法不在五类词表内（如“锂电池”）时，品牌也不能被下游丢弃。"""
    from modules.battery_review import extract_battery_items

    attachments = [{
        "filename": "application.zip",
        "structured_records": [{
            "record_type": "weee_catalog",
            "sheet_name": "Sheet1",
            "row_number": 3,
            "brand": "HONGSAMDE",
            "category": "锂电池",
            "raw_text": "示例产品 | 锂电池 | HONGSAMDE",
        }],
    }]
    result = extract_battery_items(project="德国电池法", attachments=attachments)
    brands = [item["brand"] for item in result["items"] if item["brand"]]
    assert "HONGSAMDE" in brands


def test_header_text_is_never_returned_as_brand_or_category():
    """表头文字/计量说明不能被当成品牌或类别值。

    真实泛欧模板里同一列会叠两行表头（中文一行、英文一行），
    或把“预估销售量/单个电池重量”写在相邻列；这些文字一旦进入
    brand 字段，业务会误以为提取成功，比留空更糟。
    """
    rows = [
        ["公司中文名称", "产品类别", "产品名称", "对应的品牌\nBrand", "预估销售量"],
        ["示例甲贸易有限公司", "小型设备", "示例产品A", "对应的品牌\nBrand", "5T"],
        ["", "", "", "uwant", ""],
    ]
    records = parser._xlsx_structured_records(rows, "Sheet1")
    for record in records:
        brand = str(record.get("brand") or "")
        assert "品牌" not in brand
        assert "Brand" not in brand
        assert "销售" not in brand
