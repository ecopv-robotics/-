import json

from workbench_server import _hide_superseded_review_rows

BASE = {
    "发件人邮箱": "lingting@funtax.cn",
    "发件日期": "2026-08-26 17:32:34",
    "邮件主题": "QSW582 深圳晟隆钛实业有限公司 德国WEEE 新增类目",
    "_source": "人工补全",
}


def _row(**overrides):
    row = dict(BASE)
    row.update(overrides)
    return row


def _weee(items):
    return json.dumps(items, ensure_ascii=False)


def _valid_row():
    return _row(
        客户公司名称="深圳晟隆钛实业有限公司",
        语义校验状态="valid",
        德国WEEE品类明细=_weee([
            {"brand": "", "category": "第五类——小型非光伏设备", "category_class": "5"},
        ]),
        人工复核提示="完成后再导入工单核对",
    )


def _conflict_row():
    return _row(
        客户公司名称="Shenzhenshenglongtaishiye Co., Ltd",
        语义校验状态="invalid",
        语义问题编号="COMPANY_SOURCE_CONFLICT",
        语义建议值="company: 深圳晟隆钛实业有限公司",
        德国WEEE品类明细=_weee([{"brand": "Shenzhenshenglongtaishiye Co., Ltd.", "category": ""}]),
        人工复核提示="LLM语义校验: 程序误将原主体注册品牌当作客户公司",
    )


def test_company_conflict_row_is_hidden_and_brand_moves_to_valid_row():
    kept = _hide_superseded_review_rows([_conflict_row(), _valid_row()])

    assert len(kept) == 1
    assert kept[0]["客户公司名称"] == "深圳晟隆钛实业有限公司"
    items = json.loads(kept[0]["德国WEEE品类明细"])
    # 品牌补进已有分类项，不能拆成两张卡。
    assert len(items) == 1
    assert items[0]["brand"] == "Shenzhenshenglongtaishiye Co., Ltd."
    assert items[0]["category"] == "第五类——小型非光伏设备"
    assert "已隐藏 1 条公司冲突行" in kept[0]["人工复核提示"]


def test_conflict_row_kept_when_no_valid_sibling_exists():
    kept = _hide_superseded_review_rows([_conflict_row()])
    assert len(kept) == 1
    assert kept[0]["客户公司名称"] == "Shenzhenshenglongtaishiye Co., Ltd"


def test_conflict_row_kept_when_sibling_is_also_invalid():
    sibling = _valid_row()
    sibling["语义校验状态"] = "invalid"
    kept = _hide_superseded_review_rows([_conflict_row(), sibling])
    assert len(kept) == 2


def test_invalid_row_without_company_suggestion_is_untouched():
    """没有公司建议值的 invalid 行不能靠猜测隐藏（可能是真业务）。"""
    other = _conflict_row()
    other["语义问题编号"] = "FIELD_FORMAT"
    other["语义建议值"] = ""
    kept = _hide_superseded_review_rows([other, _valid_row()])
    assert len(kept) == 2


def test_invalid_row_with_pointer_to_sibling_is_hidden_regardless_of_code():
    """编号会变（CONFLICT / CODE_MIXED…），判据是"建议值指向同邮件有效公司"。"""
    mixed = _conflict_row()
    mixed["语义问题编号"] = "COMPANY_CODE_MIXED"
    kept = _hide_superseded_review_rows([mixed, _valid_row()])
    assert len(kept) == 1
    assert kept[0]["客户公司名称"] == "深圳晟隆钛实业有限公司"


def test_noise_company_row_is_hidden_when_valid_sibling_exists():
    noisy = _row(
        客户公司名称="新增品牌 YBOEG 第五类HST011 郑州兰序网络科技有限公司",
        语义校验状态="invalid",
        语义问题编号="COMPANY_CODE_MIXED",
    )
    kept = _hide_superseded_review_rows([noisy, _valid_row()])
    assert len(kept) == 1
    assert kept[0]["客户公司名称"] == "深圳晟隆钛实业有限公司"


def test_noise_company_row_is_kept_when_it_is_the_only_row():
    noisy = _row(
        客户公司名称="新增品牌 YBOEG 第五类HST011 郑州兰序网络科技有限公司",
        语义校验状态="invalid",
        语义问题编号="COMPANY_CODE_MIXED",
    )
    kept = _hide_superseded_review_rows([noisy])
    assert len(kept) == 1
