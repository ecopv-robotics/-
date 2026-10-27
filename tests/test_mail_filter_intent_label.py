"""「筛选依据」标签必须能区分「规则命中」与「LLM 救回」。

背景：全量模式下规则已判有效的邮件也会走一遍意图 Agent。旧逻辑无条件把
filter_reason 改写成「LLM恢复(...)」，导致 08-24~08-26 那批 100 条里 85 条
本来是「方法一」命中的，全被覆盖 —— 事后完全看不出规则层有没有起作用。
"""

from modules.mail_filter import MailFilter


class _TargetLLM:
    enabled = True

    def classify_email(self, subject, body, attachments):
        return {"is_target": True, "intent": "新增",
                "reason": "正文明确新增品牌", "confidence": "high"}


def _mail(status, reason):
    return {"subject": "示例邮件", "body_text": "", "attachments": [],
            "filter_status": status, "filter_reason": reason}


def test_rule_valid_mail_keeps_method_one_label():
    recovered, _ = MailFilter(set()).llm_second_pass(
        [_mail("valid", "方法一: 注册关键字+业务词命中")], _TargetLLM()
    )
    assert recovered[0]["filter_reason"] == "方法一: 注册关键字+业务词命中 | LLM确认(意图=新增)"


def test_rule_filtered_mail_is_labelled_as_llm_recovery():
    recovered, _ = MailFilter(set()).llm_second_pass(
        [_mail("filtered", "外部无关邮件")], _TargetLLM()
    )
    assert recovered[0]["filter_reason"].startswith("LLM恢复(意图=新增")


def test_rule_uncertain_mail_is_labelled_as_llm_recovery():
    recovered, _ = MailFilter(set()).llm_second_pass(
        [_mail("uncertain", "业务关键字命中但未见注册动作，需人工复核")], _TargetLLM()
    )
    assert recovered[0]["filter_reason"].startswith("LLM恢复(意图=新增")
