import json

from modules.weee_category_audit import (
    classify_weee_product,
    compare_weee_categories,
    decompose_weee_category,
    extract_weee_items,
)


def test_classification_table_rules_are_conservative():
    assert classify_weee_product("手机")["category_class"] == "6"
    assert classify_weee_product("空调")["category_class"] == "1"
    assert classify_weee_product("LED灯")["category_class"] == "3"
    # 传统白炽灯在文件中明确不属于第 3 类，不能被灯具关键词误收。
    assert classify_weee_product("传统白炽灯")["status"] == "unmatched"
    # 未提供尺寸的“打印机”无法在大/小/IT 设备之间确定，不自动猜类别。
    assert classify_weee_product("打印机")["status"] == "ambiguous"


def test_mail_brand_and_category_are_paired():
    result = extract_weee_items(
        "德国 WEEE 品类清单",
        "品牌：ACME 类别：手机",
        [],
        "德国WEEE",
    )
    assert result["enabled"] is True
    assert result["status"] == "ready"
    assert result["items"] == [
        {
            "brand": "ACME",
            "category": "手机",
            "sources": ["邮件正文/标题"],
            "evidences": ["品牌：ACME；类别：手机"],
            "confidence": "medium",
            "evidence": "品牌：ACME；类别：手机",
            "category_original": "手机",
            "category_class": "6",
            "category_class_name": "小型信息和电信设备",
            "category_class_status": "matched",
            "category_candidates": [
                {"category_class": "6", "category_class_name": "小型信息和电信设备", "score": 12}
            ],
            "category_class_evidence": ["手机"],
            "category_rule_source": "产品分类表 中文版.docx（德国 ElektroG/WEEE 六类）",
        }
    ]


def test_composite_category_label_is_extracted_and_mapped_to_large_equipment():
    result = extract_weee_items(
        "上海古道-示例62f251e9有限公司-新增类目：电气与电子设备废料-大型设备(光伏面板除外)",
        "新增类目：电气与电子设备废料-大型设备(光伏面板除外)\n品牌：YWNYT",
        [],
        "德国WEEE",
    )
    assert result["status"] == "ready"
    assert len(result["items"]) == 1
    item = result["items"][0]
    assert item["brand"] == "YWNYT"
    assert item["category"] == "电气与电子设备废料-大型设备(光伏面板除外)"
    assert item["category_class"] == "4"
    assert item["category_class_name"] == "大型设备"
    assert item["category_components"] == {
        "category_family": "电气与电子设备废料",
        "equipment_size": "大型设备",
        "exclusions": ["光伏面板"],
    }
    assert decompose_weee_category(item["category"])["equipment_size"] == "大型设备"


def test_weee_items_are_scoped_to_the_current_company():
    attachments = [{
        "filename": "德国weee新增品牌.xlsx",
        "structured_records": [
            {
                "customer": "示例c53c9bb4有限公司",
                "brand": "BrandA",
                "category": "Kleingeräte 第五类",
                "raw_text": "示例c53c9bb4有限公司 BrandA Kleingeräte 第五类",
            },
            {
                "customer": "示例22203a3f有限公司",
                "brand": "BrandB",
                "category": "Kleingeräte 第五类",
                "raw_text": "示例22203a3f有限公司 BrandB Kleingeräte 第五类",
            },
        ],
    }]
    body = "示例c53c9bb4有限公司 新增品牌 BrandA 第五类\n示例22203a3f有限公司 新增品牌 BrandB 第五类"

    company_a = extract_weee_items(
        "德国WEEE新增品牌", body, attachments, "德国WEEE", company="示例c53c9bb4有限公司"
    )
    company_b = extract_weee_items(
        "德国WEEE新增品牌", body, attachments, "德国WEEE", company="示例22203a3f有限公司"
    )

    assert [(item["brand"], item["category_class"]) for item in company_a["items"]] == [("BrandA", "5")]
    assert [(item["brand"], item["category_class"]) for item in company_b["items"]] == [("BrandB", "5")]


def test_workorder_category_detail_is_compared_by_class_and_brand():
    mail = extract_weee_items(
        "德国 WEEE",
        "品牌：ACME；品类：手机",
        [],
        "德国WEEE",
    )
    matched = compare_weee_categories(
        mail["items"],
        [{"品类明细": [{"品类": "小型信息和电信设备"}], "品牌": "ACME"}],
    )
    assert matched["status"] == "matched"
    assert matched["items"][0]["status"] == "matched"

    missing = compare_weee_categories(
        mail["items"],
        [{"品类明细": [{"品类": "屏幕和显示设备"}], "品牌": "ACME"}],
    )
    assert missing["status"] == "missing"

    pending = compare_weee_categories(mail["items"], [{"工单编号": "WO-1"}])
    assert pending["status"] == "pending"


class _FakeWEEEReviewLLM:
    enabled = True

    def __init__(self, review, reextract=None):
        self.review = review
        self.reextract = reextract
        self.review_calls = 0
        self.reextract_calls = 0

    def review_weee_extraction(self, **_kwargs):
        self.review_calls += 1
        return self.review

    def extract_weee_fields(self, **_kwargs):
        self.reextract_calls += 1
        return self.reextract

    def classify_weee_categories(self, _items):
        return []


def test_llm_reviews_rule_extraction_without_replacing_valid_items():
    llm = _FakeWEEEReviewLLM({"status": "valid", "confidence": "high", "issues": []})
    result = extract_weee_items(
        "FormiPow 德国WEEE注册",
        "品牌：FormiPow；品类：热交换设备",
        [],
        "德国WEEE",
        llm_client=llm,
    )

    assert llm.review_calls == 1
    assert llm.reextract_calls == 0
    assert result["extraction_method"] == "规则提取·LLM复检通过"
    assert result["items"][0]["brand"] == "FormiPow"
    assert result["items"][0]["category"] == "热交换设备"


def test_llm_reextracts_when_review_rejects_rule_items_and_requires_evidence():
    llm = _FakeWEEEReviewLLM(
        {"status": "invalid", "confidence": "low", "issues": [{"code": "MISSING"}]},
        {
            "items": [
                {
                    "item_id": "0",
                    "brand": "PURELLEL",
                    "category": "小型非光伏",
                    "source": "邮件正文",
                    "evidence": "新增品牌：PURELLEL；新增类别：小型非光伏",
                    "confidence": "high",
                },
                # 不在原文出现的模型臆测必须被本地证据校验过滤。
                {
                    "item_id": "1",
                    "brand": "幻觉品牌",
                    "category": "手机",
                    "source": "邮件正文",
                    "evidence": "模型推测",
                    "confidence": "high",
                },
            ],
            "reason": "规则字段错配，已按原文重新提取",
        },
    )
    result = extract_weee_items(
        "德国WEEE 新增品牌：PURELLEL",
        "新增品牌：PURELLEL；新增类别：小型非光伏",
        [],
        "德国WEEE",
        llm_client=llm,
    )

    assert llm.review_calls == 1
    assert llm.reextract_calls == 1
    assert result["extraction_method"] == "LLM重新提取"
    assert [(item["brand"], item["category"]) for item in result["items"]] == [("PURELLEL", "小型非光伏")]


def test_scoped_llm_reextract_cannot_restore_other_company_rows():
    llm = _FakeWEEEReviewLLM(
        {"status": "invalid", "confidence": "low", "issues": [{"code": "MISMATCH"}]},
        {
            "items": [
                {"item_id": "0", "brand": "BrandA", "category": "Kleingeräte 第五类", "confidence": "high"},
                {"item_id": "1", "brand": "BrandB", "category": "Kleingeräte 第五类", "confidence": "high"},
            ],
            "reason": "复检重提取",
        },
    )
    attachment = {
        "filename": "德国weee新增品牌.xlsx",
        "text_content": "示例c53c9bb4有限公司 BrandA Kleingeräte 第五类\n示例22203a3f有限公司 BrandB Kleingeräte 第五类",
        "structured_records": [
            {
                "customer": "示例c53c9bb4有限公司",
                "brand": "BrandA",
                "category": "Kleingeräte 第五类",
                "raw_text": "示例c53c9bb4有限公司 BrandA Kleingeräte 第五类",
            },
            {
                "customer": "示例22203a3f有限公司",
                "brand": "BrandB",
                "category": "Kleingeräte 第五类",
                "raw_text": "示例22203a3f有限公司 BrandB Kleingeräte 第五类",
            },
        ],
    }

    result = extract_weee_items(
        "德国WEEE新增品牌",
        "",
        [attachment],
        "德国WEEE",
        llm_client=llm,
        company="示例c53c9bb4有限公司",
    )

    assert [(item["brand"], item["category_class"]) for item in result["items"]] == [
        ("BrandA", "5")
    ]
