"""需求（业务动作）识别：注册 / 取消注册 / 变更 / 续费 / 新增 / 撤单。

真实邮件里“取消注册”包含“注册”子串，旧实现按固定优先级判断，
会把取消单识别成新注册，业务会朝相反方向去核对工单。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from modules.field_extractor import FieldExtractor  # noqa: E402


def _extractor():
    return FieldExtractor({}, [], ocr_fallback=False)


def _need(subject="", body=""):
    return _extractor()._extract_need(subject + " " + body, subject=subject)


def test_cancel_registration_is_not_read_as_registration():
    subject = "Re:创易购易+CTYG062+示例泽电电力工程有限公司 德国包装法+德国电池法 取消注册 谢谢!!!!"
    body = ("示例泽电电力工程有限公司\n\n德国包装法+德国电池法 取消注册 麻烦了\n\n"
            "主题：示例代理+示例公司 德国包装法+德国电池法 注册\n示例公司注册以下项目：\n德国包装法+德国电池法")
    assert _need(subject, body) == "取消注册"


def test_cancel_in_body_wins_when_subject_only_says_register():
    subject = "示例代理+示例公司 德国电池法 注册"
    body = "您好，这个德国电池法不做了，麻烦取消注册，谢谢。"
    assert _need(subject, body) == "取消注册"


def test_plain_registration_still_reads_as_registration():
    assert _need("示例代理+示例公司 德国电池法 注册", "请查收附件资料") == "注册"


def test_negated_cancel_is_not_treated_as_cancel():
    """“不要取消注册”是确认继续，不是取消。"""
    assert _need("示例公司 德国电池法 注册", "请不要取消注册，继续办理") in {"注册", ""}


def test_cancellation_keywords_variants():
    for text in ("撤销注册", "注销注册", "取消申请", "这个不做了"):
        assert _need("示例公司 德国电池法", text) == "取消注册", text


def test_other_actions_are_kept():
    assert _need("示例公司 德国WEEE 撤单") == "撤单"
    assert _need("示例公司 德国WEEE 新增品牌") == "新增"
    assert _need("示例公司 德国包装法 变更信息") == "变更"


def test_cancellation_row_gets_a_dedicated_review_hint():
    """取消单不走工单存在性判定，提示要写清楚，不能被当成普通漏单。"""
    from modules.excel_writer import ExcelWriter

    hint = ExcelWriter._review_hint({"需求": "取消注册"})
    assert "取消单" in hint
    assert "撤销" in hint

    plain = ExcelWriter._review_hint({"需求": "注册"})
    assert "取消单" not in plain


def test_cancellation_row_gets_a_dedicated_review_hint():
    """取消单不走工单存在性判定，提示要写清楚，不能被当成普通漏单。"""
    from modules.excel_writer import ExcelWriter

    hint = ExcelWriter._review_hint({"需求": "取消注册"})
    assert "取消单" in hint
    assert "撤销" in hint

    plain = ExcelWriter._review_hint({"需求": "注册"})
    assert "取消单" not in plain
