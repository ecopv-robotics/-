import json

from modules.field_extractor import FieldExtractor


def _extractor():
    return FieldExtractor({}, [])


def test_english_legal_suffix_keeps_full_company_name():
    extractor = _extractor()
    assert extractor._company_substring("公司名称：North Star Corp s. r. o") == "North Star Corp s. r. o"


def test_cancel_sentence_does_not_reduce_company_to_suffix_only():
    extractor = _extractor()
    text = "现正式申请撤销明日科技 (香港) 有限公司项下德国 WEEE 业务订单"
    assert extractor._company_substring(text) == "明日科技 (香港) 有限公司"


def test_structured_body_keeps_two_cancel_companies():
    extractor = _extractor()
    groups = extractor._structured_groups(
        "2026.08.24 批量撤单",
        "TBA+Fixture 327da5a9 Inc+波兰包装法\n\nTBA+Fixture 72d0887b LLC+波兰包装法",
    )
    assert [(g["customer"], g["projects"]) for g in groups] == [
        ("Fixture 327da5a9 Inc", ["波兰包装法"]),
        ("Fixture 72d0887b LLC", ["波兰包装法"]),
    ]


def test_subject_agent_alias_with_code_prefix_is_recognized():
    extractor = FieldExtractor(
        {"agent@example.test": {"代理": "首阳跨境咨询", "代理简称": "SY"}},
        [],
    )
    result = extractor._extract_agent("unknown@example.test", "SY003-首阳跨境咨询-德国电池法", "", [])
    assert result["agent"] == "首阳跨境咨询"


def test_pan_europe_form_is_gated_by_declared_subject_projects():
    names = [
        "意大利WEEE",
        "意大利电池法",
        "意大利包装法",
        "西班牙WEEE",
        "西班牙电池法",
        "西班牙包装法",
    ]
    extractor = FieldExtractor({}, [{"项目名称": name} for name in names])
    result = extractor._extract_projects(
        "客户-意大利WEEE+意大利电池法+意大利包装法",
        "",
        [
            {
                "filename": "泛欧EPR申请表-8国.xlsx",
                "epr_forms": [
                    {
                        "filename": "泛欧EPR申请表-8国.xlsx",
                        "projects": names,
                    }
                ],
            }
        ],
        [],
    )
    assert [item["standard_name"] for item in result["projects"]] == [
        "意大利WEEE",
        "意大利电池法",
        "意大利包装法",
    ]


def test_epr_application_company_precedes_subject_and_body_candidates():
    extractor = _extractor()
    rows = extractor.extract_fields({
        "subject": "示例62462b35有限公司 德国WEEE注册",
        "body_text": "示例39ecdfaf有限公司请办理德国WEEE注册",
        "attachments": [{
            "filename": "EPR申请表.xlsx",
            "text_content": "公司中文名称：示例84ed7eaa有限公司",
        }],
    })

    assert len(rows) == 1
    assert rows[0]["客户"] == "示例84ed7eaa有限公司"
    assert rows[0]["客户提取来源"] == "附件EPR申请表"


def test_non_epr_attachment_does_not_block_text_fallback():
    extractor = _extractor()
    result = extractor._extract_customer(
        "示例62462b35有限公司 德国WEEE注册",
        "请查收资料",
        [{"filename": "logo.png", "text_content": ""}],
    )

    assert result["customer"] == "示例62462b35有限公司"
    assert result["source"] == "主题"


def test_attachment_file_index_keeps_same_named_attachments_separate():
    items = json.loads(FieldExtractor._attachment_file_index_json([
        {"filename": "公司资料.jpg", "filepath": r"cache\attachments\first.jpg"},
        {"filename": "公司资料.jpg", "filepath": r"cache\attachments\second.jpg"},
    ]))

    assert [item["filename"] for item in items] == ["公司资料.jpg", "公司资料.jpg"]
    assert [item["token"] for item in items] == ["first.jpg", "second.jpg"]
    assert [item["attachment_index"] for item in items] == [0, 1]
def test_numbered_bilingual_companies_do_not_expand_agent_signature():
    extractor = FieldExtractor({}, [], ocr_fallback=False, agent_identity_aliases={
        "示例代理": ["深圳示例代理检测有限公司", "Shenzhen Example Agent Co., Ltd"],
    })
    subject = "示例代理-甲甲科技有限公司、乙乙贸易有限公司-2家公司的荷兰包装法取消注册，撤单"
    body = ("以下2家公司的荷兰包装法客户决定取消注册，撤单处理。\n"
            "1、甲甲科技有限公司（ALPHA TECHNOLOGY LIMITED）\n"
            "2、乙乙贸易有限公司（BETA TRADING LIMITED）\n\n"
            "深圳示例代理检测有限公司\nShenzhen Example Agent Co., Ltd")
    groups = extractor._structured_groups(subject, body)
    assert [group["customer"] for group in groups] == ["甲甲科技有限公司", "乙乙贸易有限公司"]
    assert groups[0]["company_aliases"] == ["ALPHA TECHNOLOGY LIMITED"]
    assert not any(group["needs_review"] for group in groups)
    assert extractor._is_agent_name("shenzhen unrelated test co., ltd") is False
    assert extractor._is_agent_name("SHENZHEN EXAMPLE AGENT CO., LTD") is True
    rows = extractor.extract_fields({"subject": subject, "body_text": body, "attachments": []})
    assert [(row["客户"], row["项目"]) for row in rows] == [
        ("甲甲科技有限公司", "荷兰包装法"), ("乙乙贸易有限公司", "荷兰包装法")
    ]


def test_weee_category_phrase_implies_project_only_without_other_country():
    extractor = FieldExtractor({}, [], ocr_fallback=False)
    phrase = "新增类目：电气与电子设备废料-大型设备(光伏面板除外)"
    assert [item["standard_name"] for item in extractor._extract_projects_by_rules(phrase)] == ["德国WEEE"]
    assert extractor._extract_projects_by_rules("法国" + phrase) == []


def test_full_extraction_preserves_english_company_leading_letters():
    extractor = FieldExtractor({}, [], ocr_fallback=False)
    rows = extractor.extract_fields({
        "subject": "回复: 示例代理-ExampleZone Malta Limited-2026年意大利包装法注册",
        "body_text": "你好，ExampleZone Malta Limited意大利包装都撤单，取消注册\n\n公司：ExampleZone Malta Limited",
        "attachments": [],
    })
    assert [row["客户"] for row in rows] == ["ExampleZone Malta Limited"]
    assert extractor._repair_truncated_company(
        "ampleZone Malta Limited", "ExampleZone Malta Limited 意大利包装法", ""
    ) == "ExampleZone Malta Limited"


def test_company_suffix_from_unmatched_parenthetical_note_is_not_a_customer():
    extractor = FieldExtractor({}, [], ocr_fallback=False)
    rows = extractor.extract_fields({
        "subject": "Example Agent-示例甲有限公司-荷兰电池法注销（欧洲店不运营）",
        "body_text": "",
        "attachments": [],
    })
    assert rows
    assert all(row["客户"] == "示例甲有限公司" for row in rows)


def test_short_latin_proxy_alias_does_not_truncate_real_company_prefix():
    extractor = FieldExtractor({}, [], ocr_fallback=False, agent_identity_aliases={
        "示例代理": ["Cl"],
    })
    assert extractor._strip_known_agent_prefix("ClickZone Malta Limited") == "ClickZone Malta Limited"
    assert extractor._strip_known_agent_prefix("Cl-ClickZone Malta Limited") == "ClickZone Malta Limited"


def test_registered_proxy_signature_matches_without_sender_email_candidate():
    extractor = FieldExtractor({}, [], ocr_fallback=False, agent_identity_aliases={
        "示例代理": ["华之星检测有限公司", "华之星"],
    })
    result = extractor._extract_agent(
        "unknown@example.invalid", "普通业务主题", "正文内容\n\n华之星 WEEE-陈助理", [],
    )
    assert result["agent"] == "示例代理"
    assert result["match_method"] == "签名代理别名精确匹配"


def test_epr_application_projects_fall_back_to_explicit_single_applicant_scope():
    extractor = FieldExtractor({}, [], ocr_fallback=False)
    attachments = [{
        "filename": "application.xlsx",
        "epr_forms": [{"filename": "application.xlsx", "projects": []}],
        "structured_records": [
            {"record_type": "epr_application", "attachment_name": "application.xlsx",
             "customer": "示例甲有限公司", "company_field": "company_name_zh",
             "raw_text": "公司中文名称：示例甲有限公司"},
            {"record_type": "epr_application", "attachment_name": "application.xlsx",
             "customer": "SAMPLE ALPHA LIMITED", "company_field": "company_name_en",
             "raw_text": "公司英文名称：SAMPLE ALPHA LIMITED"},
            {"record_type": "weee_catalog", "attachment_name": "application.xlsx",
             "customer": "SAMPLE ALPHA LIMITED", "brand": "SAMPLE BRAND",
             "category": "Large equipment", "row_number": 8},
        ],
    }]
    rows = extractor.extract_fields({
        "subject": "示例甲有限公司-德国WEEE+德国电池法注册",
        "body_text": "",
        "attachments": attachments,
    })
    assert {(row["客户"], row["项目"]) for row in rows} == {
        ("示例甲有限公司", "德国WEEE"), ("示例甲有限公司", "德国电池法"),
    }


def test_weee_category_record_is_never_expanded_as_customer_business():
    extractor = FieldExtractor({}, [], ocr_fallback=False)
    groups = extractor._attachment_structured_groups([{
        "filename": "application.xlsx",
        "structured_records": [{
            "record_type": "weee_catalog", "customer": "SAMPLE ALPHA LIMITED",
            "brand": "SAMPLE BRAND", "category": "Large equipment",
            "row_number": 8,
        }],
    }], {"projects": [{"standard_name": "德国WEEE"}]})
    assert groups == []


def test_decode_replacement_character_forces_manual_review():
    rows = FieldExtractor({}, [], ocr_fallback=False).extract_fields({
        "subject": "SAMPLE ALPHA LIMITED 德国WEEE注册�",
        "body_text": "",
        "attachments": [],
    })
    assert rows
    assert all(row["置信度"] == "需人工确认" for row in rows)
    assert all(row["邮件解码状态"] == "检测到解码替代字符" for row in rows)


def test_company_candidate_after_agent_prefix_drops_client_code():
    """R5/R6：代理前缀在客户编号之前时，编号也必须被剥离。

    截图实例：代理「出海者」+ 编号 CHZ489 + 公司全名，最终只输出公司全名。
    """
    extractor = FieldExtractor({"agent@example.invalid": {"代理": "出海者"}}, [])
    assert extractor._company_substring("出海者 CHZ489 深圳市聚恒希科技有限公司") == "深圳市聚恒希科技有限公司"
    result = extractor._sanitize_customer_result(
        {"customer": "出海者 CHZ489 深圳市聚恒希科技有限公司", "source": "规则"}
    )
    assert result["customer"] == "深圳市聚恒希科技有限公司"


def test_company_candidate_without_agent_prefix_still_drops_client_code():
    """无代理前缀的直接形态仍保持原有剥离行为。"""
    extractor = FieldExtractor({}, [])
    assert extractor._company_substring("CHZ489 深圳市聚恒希科技有限公司") == "深圳市聚恒希科技有限公司"


# 发件人 hflee@bnlsupplies.example 本身未登记，只有同域名的其他邮箱登记为 BNL，
# 与生产现象一致：发件邮箱走了主题别名匹配，域名证据被忽略。
_AGENT_MAP = {
    "sales@bnlsupplies.example": {"代理": "BNL"},
    "ops@yisi.example": {"代理": "易思"},
}


def test_agent_domain_and_subject_conflict_goes_to_manual_review():
    """R22：发件域名→BNL、主题别名→易思，两条证据冲突时不得自动改判。"""
    extractor = FieldExtractor(_AGENT_MAP, [])
    result = extractor._extract_agent(
        "hflee@bnlsupplies.example", "易思 挪威WEEE 注册", "正文", []
    )
    assert result["agent"] == ""
    assert "待确认" in result["match_method"]
    assert "冲突" in result["match_method"]
    assert result["conflict"]["domain_agent"] == "BNL"
    assert result["conflict"]["subject_agent"] == "易思"


def test_agent_domain_and_subject_agreeing_is_accepted():
    extractor = FieldExtractor(_AGENT_MAP, [])
    result = extractor._extract_agent(
        "hflee@bnlsupplies.example", "BNL 挪威WEEE 注册", "正文", []
    )
    assert result["agent"] == "BNL"


def test_agent_subject_match_still_used_without_domain_evidence():
    """发件域名不在代理表内时，主题别名匹配保持原有行为。"""
    extractor = FieldExtractor(_AGENT_MAP, [])
    result = extractor._extract_agent(
        "someone@unknown.example", "易思 挪威WEEE 注册", "正文", []
    )
    assert result["agent"] == "易思"
