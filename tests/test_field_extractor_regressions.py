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
