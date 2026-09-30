"""M2 邮件过滤模块 — 规则引擎 + LLM 二次意图识别

流程:
  F1 规则预分类 → 产出 valid_mails + filtered_mails，并记录命中依据
  F2 LLM 全量意图识别 → 对本次拉取的每封邮件逐封判断是否为新询单
"""
import os
import re
from typing import List, Dict, Set, Optional


# SOP 规定的 ECOPV 内部收件人域。这里保留为代码常量，避免外部表格缺一行
# 就失去内部邮件防线；内部邮箱表仍用于精确地址记录与人工审计。
INTERNAL_RECIPIENT_SUFFIXES = tuple(
    '@' + value.strip().lower().lstrip('@')
    for value in os.getenv('MAIL_INTERNAL_DOMAINS',
        'ecopv-net.example.invalid,ecopv-cn.example.invalid,ecopv-com.example.invalid,intake.example.invalid').split(',')
    if value.strip()
)

# 内部发件人域名标记。按用户约定，只检查发件人邮箱的域名，
# 不扫描正文、主题或收件人中的“ecopv”，避免误过滤客户材料中的文字。
INTERNAL_SENDER_DOMAIN_MARKERS = tuple(
    value.strip().lower() for value in os.getenv('MAIL_INTERNAL_DOMAIN_MARKERS', 'ecopv').split(',') if value.strip()
)


def is_internal_sender_address(value) -> bool:
    """Shared gate for newly read mail and historical workbench rows; sender only."""
    raw = str(value or "").strip().lower()
    addresses = re.findall(r"[\w.+-]+@[\w.-]+\.[a-z]{2,}", raw, re.I)
    address = addresses[0] if addresses else raw
    domain = address.rsplit("@", 1)[-1] if "@" in address else ""
    internal_domains = [suffix.lstrip("@") for suffix in INTERNAL_RECIPIENT_SUFFIXES]
    return bool(domain) and (
        any(domain == item or domain.endswith("." + item) for item in internal_domains)
        or any(marker in domain for marker in INTERNAL_SENDER_DOMAIN_MARKERS)
    )

# 企业邮箱常用“登录账号 + 公共收件地址”方式接收询单。登录账号可能是
# fixture-74bc52e7@intake.example.invalid，但客户邮件的 To 头写的是 report@intake.example.invalid。
# 这类公共地址属于本次审计入口，不能被当作“发给内部同事”而过滤。
DEFAULT_AUDIT_MAILBOX_ALIASES = {
    value.strip().lower() for value in os.getenv('MAIL_AUDIT_ALIASES', 'report@intake.example.invalid').split(',') if value.strip()
}

# 方法一：主题含注册类关键字（必须同时含业务词才算有效）
KEYWORDS_ACTION = {"注册", "新增", "撤单", "注销", "追加"}

# 不同代理对“要办理新项目”的写法不一致。它们与业务词同时出现时，
# 应进入解析链路；仅凭“资料/更新”仍保留人工或 LLM 复核，不直接判定为注册。
KEYWORDS_REQUEST = {
    "申请", "办理", "提交", "加入", "开通", "新注册", "注册申请", "申请表",
    "application", "register", "registration", "onboarding",
}

# 明确表达“现在要办”的请求词。把 register/registration 与普通请求词分开，
# 是为了不把“registration completed / registration certificate”这类交付通知
# 误判成新申请；它们仍可用于普通业务命中，但不能阻止完成态证书过滤。
KEYWORDS_EXPLICIT_REQUEST = {
    "申请", "办理", "提交", "加入", "开通", "新注册", "注册申请", "申请表",
    "application", "apply", "onboarding",
}

KEYWORDS_ATTACHMENT_REQUEST = {
    "申请表", "application form", "registration form", "onboarding form",
}

# 业务关键字（项目/法案名称）
KEYWORDS_BUSINESS = {
    "WEEE", "weee", "电池法", "电池", "包装法", "EPR", "epr",
    "一次性塑料", "BAT", "bat",
    "德国", "法国", "意大利", "西班牙", "荷兰", "波兰", "瑞典",
    "比利时", "爱尔兰", "葡萄牙", "奥地利",
    "丹麦", "捷克", "芬兰", "挪威", "匈牙利", "罗马尼亚", "保加利亚",
    "希腊", "克罗地亚", "卢森堡", "爱沙尼亚", "瑞士", "英国", "加拿大",
    "拉脱维亚", "包装", "packaging", "verpackg", "ppwr", "battery", "batteries",
}

# 明确无关的关键词（出现这些且不含业务词 → 过滤）
KEYWORDS_IRRELEVANT = {
    "账单", "invoice", "合同", "协议",
    "咨询", "反馈", "建议",
    "下证", "证书号", "证书下发",
    "保证金", "担保", "退款",
    "回收费", "购买",
}

# 非注册业务邮箱：收款/水单、报价/续费、信息变更、发票等，属于成交后的
# 运营财务事务，不在 SOP 的 注册/新增/撤单 范围内。
#
# STRICT：命中即过滤，即使主题里带"申请/办理"等动作词。
#   理由——"申请开票""申请修改公司名称"仍然是开票/改信息，不是新注册。
# SOFT：只有主题里没有注册动作词时才过滤。
#   理由——"请求签署授权书"本身不是注册需求，但"新增德国WEEE注册(附授权书)"
#   里的授权书只是配套材料，不能因此丢掉这条注册需求。
#
# 两个词表都只扫主题，不扫正文和附件名：正文/签名档/报价附件里出现
# "价格""发票""payment"太常见，扫全文会误杀真实注册邮件。
KEYWORDS_NON_REGISTRATION_STRICT = {
    # 收款 / 水单 / 到账
    "收款", "已收款", "付款", "回款", "汇款", "转账", "水单", "到账",
    "对账", "bank slip", "remittance", "payment received", "payment",
    # 报价 / 续费 / 年费 / 价格
    "报价", "报价单", "询价", "价格", "年费", "续费", "续期", "服务费",
    "收费标准", "quotation", "quote", "renewal",
    # 修改公司名称 / 地址 / 信息变更
    "修改公司名称", "公司名称修改", "公司名称变更", "变更公司名称",
    "修改公司名", "改公司名", "公司更名", "更名",
    "修改地址", "地址修改", "地址变更", "变更地址",
    "修改公司信息", "公司信息变更", "信息变更", "资料变更",
    # 发票 / 账单
    "发票", "开票", "开发票", "账单", "对账单", "invoice",
}

# 公司改名/主体变更是“存量信息维护”，不是新注册询单。很多代理只在正文里写
# “原公司名…现更名为…”，主题仍然保留“WEEE/注册”，因此不能只检查主题。
# 这些规则命中会作为 LLM 的上下文证据，最终以全量意图 Agent 判断为准。
KEYWORDS_COMPANY_CHANGE = {
    "修改公司名称", "公司名称修改", "公司名称变更", "变更公司名称",
    "修改证书公司名称", "修改证书公司名", "修改公司名称和地址",
    "修改公司名和地址", "公司名称及地址变更", "公司名称与地址变更",
    "修改公司名", "改公司名", "公司更名", "公司改名", "企业更名",
    "企业名称变更", "企业名称修改", "企业名称更新", "公司名称更新",
    "公司名更新", "公司名称更换", "公司名称改为", "主体变更", "主体更换",
    "公司抬头变更", "公司抬头修改", "legal entity change", "company name change",
    "change of company name", "rename company", "renamed company", "update company name",
    "change company details", "change company address", "company address change",
}

KEYWORDS_NON_REGISTRATION_SOFT = {
    "授权书", "授权函", "待签字", "签署", "签字", "盖章", "公证",
    "power of attorney",
}

# 已完成注册后的交付/通知邮件。它们常含 WEEE、电池法等业务词，若只走
# "业务词命中" 会被误送去查新工单。
KEYWORDS_CERTIFICATE_NOTICE = {
    "证书", "下证", "下号", "注册号下发", "certificate", "certificates",
    "certification", "registration number",
}

# “注册”“新增”本身不能证明是新申请：例如“注册已完成”“新增已下证”。
# 这些完成态词须与证书/下号通知结合判断，避免误滤“申请新增证书”这类真正请求。
KEYWORDS_CERTIFICATE_COMPLETION = {
    "已完成", "注册完成", "完成注册", "注册已完成",
    "已下证", "证书已下发", "证书下发", "已下发",
    "请查收附件", "请查收证书", "证书请查收", "请妥善保存",
    "completed", "certificate attached", "attached certificate",
}


class MailFilter:
    def __init__(
        self,
        internal_emails: Set[str],
        logger=None,
        audit_mailbox: str = "",
        audit_mailboxes: Optional[Set[str]] = None,
    ):
        self.internal_emails = {e.lower().strip() for e in internal_emails}
        self.logger = logger
        self.audit_mailbox = str(audit_mailbox or "").strip().lower()
        self.audit_mailboxes = set(DEFAULT_AUDIT_MAILBOX_ALIASES)
        if self.audit_mailbox:
            self.audit_mailboxes.add(self.audit_mailbox)
        for address in audit_mailboxes or set():
            value = str(address or "").strip().lower()
            if value and "@" in value:
                self.audit_mailboxes.add(value)

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    @staticmethod
    def _contains_any(text: str, keywords) -> bool:
        text_lower = text.lower()
        for kw in keywords:
            if kw.lower() in text_lower:
                return True
        return False

    @staticmethod
    def _first_hit(text: str, keywords) -> str:
        """返回第一个命中的关键词（用于日志/过滤原因），未命中返回空串。"""
        text_lower = text.lower()
        for kw in keywords:
            if kw.lower() in text_lower:
                return kw
        return ""

    @staticmethod
    def _recipient_addresses(mail: Dict) -> List[str]:
        values = mail.get("recipient_emails") or []
        if isinstance(values, str):
            values = re.findall(r"[\w.+-]+@[\w.-]+\.\w+", values)
        if not isinstance(values, (list, tuple, set)):
            values = []
        if not values:
            values = re.findall(
                r"[\w.+-]+@[\w.-]+\.\w+", str(mail.get("recipient", "") or "")
            )
        return list(dict.fromkeys(str(value).strip().lower() for value in values if value))

    def _is_internal_recipient(self, mail: Dict) -> bool:
        """判断是否投递到内部收件人，且不误过滤正在审计的收件箱本身。

        审计程序从一个内部 IMAP 收件箱读取邮件，因此所有正常来信都可能带有
        该地址。只有邮件额外投递给的对象全部是内部域时，才按 SOP 直接过滤；
        这样不会因为“收件箱就是内部邮箱”而把整批客户询单清空。
        """
        addresses = self._recipient_addresses(mail)
        # 不能只排除登录账号：公共审计邮箱通常出现在 To 头，而 IMAP 登录
        # 账号可能是个人账号。配置中的 audit_addresses 和内置 report 别名
        # 都视为本次审计入口。
        non_audit = [item for item in addresses if item not in self.audit_mailboxes]
        if not non_audit:
            return False
        return all(item.endswith(INTERNAL_RECIPIENT_SUFFIXES) for item in non_audit)

    @staticmethod
    def _is_internal_sender(mail: Dict) -> bool:
        return is_internal_sender_address(mail.get("sender_email") or mail.get("sender"))

    def filter_mails(self, mails: List[Dict]) -> tuple:
        valid_mails = []
        filtered_mails = []

        for mail in mails:
            if mail.get("skip_reason") == "self_sent":
                filtered_mails.append({**mail, "filter_reason": "自身发送"})
                continue

            if self._is_internal_sender(mail):
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = "ECOPV内部发件人邮箱"
                # 这是用户明确指定的硬过滤条件，不能在全量意图复检时被模型恢复。
                mail["hard_filter"] = True
                mail["llm_eligible"] = False
                filtered_mails.append(mail)
                self._log(f"过滤(ECOPV内部发件人): {mail.get('subject', '')[:50]}")
                continue

            if self._is_internal_recipient(mail):
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = "ECOPV系统内部收件人邮箱"
                mail["llm_eligible"] = False
                filtered_mails.append(mail)
                self._log(f"过滤(内部收件人): {mail.get('subject', '')[:50]}")
                continue

            subject = mail.get("subject", "") or ""
            body = mail.get("body_text", "") or ""
            attachments = mail.get("attachments", [])
            # 附件文件名本身经常就是唯一的业务线索(如“德国WEEE注册申请表.xlsx”)，
            # 不能只扫已解析出的正文；正文解析失败时仍应进入待查/人工复核。
            att_text = " ".join(
                f"{a.get('filename', '')} {a.get('text_content', '')}".strip()
                for a in attachments
            )
            subject_body = subject + " " + body
            all_text = subject_body + " " + att_text

            sender = mail.get("sender_email", "").lower()

            # --- 判断逻辑 ---

            # 1. 非注册业务（收款/水单、报价/续费、信息变更、发票、签署类）
            #    这些是成交后的运营事务，不属于 SOP 的 注册/新增/撤单。
            #    这里仅记录规则证据；全量意图 Agent 仍会对该邮件做最终判断。
            subject_has_action = self._contains_any(subject, KEYWORDS_ACTION) or (
                self._contains_any(subject, KEYWORDS_REQUEST)
            )
            hit_strict = self._first_hit(subject, KEYWORDS_NON_REGISTRATION_STRICT)
            subject_company_change = self._first_hit(subject, KEYWORDS_COMPANY_CHANGE)
            # 正文只取前 5000 个字符，覆盖正文主内容，同时降低引用历史邮件/签名档
            # 中偶然出现旧变更词造成误判的概率。
            body_company_change = self._first_hit(
                body[:5000], KEYWORDS_COMPANY_CHANGE
            )
            hit_soft = self._first_hit(subject, KEYWORDS_NON_REGISTRATION_SOFT)
            if hit_strict or subject_company_change or (hit_soft and not subject_has_action):
                keyword = hit_strict or subject_company_change or hit_soft
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = f"非注册业务({keyword})"
                mail["llm_eligible"] = False
                filtered_mails.append(mail)
                self._log(f"过滤(非注册业务:{keyword}): {subject[:50]}")
                continue

            if body_company_change:
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = f"非注册业务(正文公司信息变更:{body_company_change})"
                # 正文变更邮件可能夹带“新增/注册”字样，交给意图 Agent 作最后判断；
                # 若 Agent 未启用，则保持过滤，不进入待查清单。
                mail["llm_eligible"] = True
                filtered_mails.append(mail)
                self._log(
                    f"过滤候选(正文公司信息变更，待意图复核:{body_company_change}): "
                    f"{subject[:50]}"
                )
                continue

            # 2. 无关邮件排除（含无关词且不含业务词）
            has_business = self._contains_any(all_text, KEYWORDS_BUSINESS)
            has_irrelevant = self._contains_any(all_text, KEYWORDS_IRRELEVANT)
            has_action = self._contains_any(subject_body, KEYWORDS_ACTION)
            has_request = self._contains_any(subject_body, KEYWORDS_REQUEST)
            has_attachment_request = self._contains_any(att_text, KEYWORDS_ATTACHMENT_REQUEST)
            has_explicit_request = (
                self._contains_any(subject_body, KEYWORDS_EXPLICIT_REQUEST)
                or has_attachment_request
            )
            has_target_action = has_action or has_request or has_attachment_request
            has_certificate_notice = self._contains_any(subject_body, KEYWORDS_CERTIFICATE_NOTICE)
            has_certificate_completion = self._contains_any(
                subject_body, KEYWORDS_CERTIFICATE_COMPLETION
            )

            # 完成态证书通知通常不是新增工单来源，但仍交给全量意图 Agent；
            # 完成态优先于动作词：
            # “注册已完成”“新增已下证”中的动作词描述的是过去状态，不是申请。
            if has_certificate_notice and (
                (has_certificate_completion and not has_explicit_request) or not has_target_action
            ):
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = "证书/下号通知，非新增申请"
                mail["llm_eligible"] = False
                filtered_mails.append(mail)
                self._log(f"过滤(证书通知): {subject[:50]}")
                continue

            if has_irrelevant and not has_business and not has_target_action:
                if sender in self.internal_emails:
                    reason = "内部无关邮件"
                else:
                    reason = "外部无关邮件"
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = reason
                filtered_mails.append(mail)
                self._log(f"过滤({reason}): {subject[:50]}")
                continue

            # 3. 方法一：主题/正文含注册类关键字 + 含业务词
            if has_target_action and has_business:
                mail["filter_status"] = "valid"
                mail["filter_reason"] = "方法一: 注册关键字+业务词命中"
                valid_mails.append(mail)
                self._log(f"有效(方法一): {subject[:50]}")
                continue

            # 4. 方法一扩展：主题含注册关键字（即使不含业务词，主题够明确）
            if has_target_action and (
                self._contains_any(subject, KEYWORDS_ACTION)
                or self._contains_any(subject, KEYWORDS_REQUEST)
            ):
                mail["filter_status"] = "valid"
                mail["filter_reason"] = "方法一扩展: 主题含注册关键字"
                valid_mails.append(mail)
                self._log(f"有效(方法一扩展): {subject[:50]}")
                continue

            # 5. 仅业务关键字：没有注册动作时不能直接进入 RPA。保留解析结果，
            # 但强制转人工复核，避免把咨询、授权、资料补交等过程邮件误判为漏单。
            if has_business:
                # 但如果同时含无关词且无关词更突出 → 过滤
                if has_irrelevant and not has_target_action:
                    reason = "内部无关邮件" if sender in self.internal_emails else "外部无关邮件"
                    mail["filter_status"] = "filtered"
                    mail["filter_reason"] = f"{reason}(业务词+无关词混存)"
                    filtered_mails.append(mail)
                    self._log(f"过滤({reason}): {subject[:50]}")
                    continue
                mail["filter_status"] = "uncertain"
                mail["filter_reason"] = "业务关键字命中但未见注册动作，需人工复核"
                valid_mails.append(mail)
                self._log(f"不确定(仅业务词): {subject[:50]}", "warning")
                continue

            # 6. 内部无关邮件
            if sender in self.internal_emails and not has_business and not has_target_action:
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = "内部无关邮件"
                filtered_mails.append(mail)
                self._log(f"过滤(内部无关): {subject[:50]}")
                continue

            # 7. 外部无关邮件
            if not has_business and not has_target_action:
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = "外部无关邮件"
                filtered_mails.append(mail)
                self._log(f"过滤(外部无关): {subject[:50]}")
                continue

            # 8. 不确定 — 保留
            mail["filter_status"] = "uncertain"
            mail["filter_reason"] = "不确定, 需人工复核"
            valid_mails.append(mail)
            self._log(f"不确定(保留): {subject[:50]}", "warning")

        self._log(f"规则过滤完成: 有效 {len(valid_mails)} 封, 过滤 {len(filtered_mails)} 封")
        return valid_mails, filtered_mails

    def llm_second_pass(
        self,
        candidate_mails: List[Dict],
        llm_client,
        progress_callback=None,
    ) -> tuple:
        """对传入的全量候选邮件做意图识别。

        旧版本这里会跳过 ``llm_eligible=False`` 的硬过滤项，只对规则筛出的
        不确定邮件调用模型。现在调用方传入本次 IMAP 拉取的全部邮件，规则结果
        只作为上下文和审计原因，不再决定是否调用模型。
        """
        if not llm_client or not llm_client.enabled:
            self._log("LLM 未启用, 跳过二次识别", "warning")
            return [], candidate_mails

        if not candidate_mails:
            return [], []

        # 全量模式：包括规则硬过滤、自发邮件、规则已判定有效邮件，全部调用
        # 同一意图 Agent，避免词表遗漏导致错误结论。
        to_classify = list(candidate_mails)
        self._log(f"LLM 全量意图识别: 共 {len(to_classify)} 封待分类邮件")

        recovered = []
        still_filtered = []

        total = len(to_classify)
        for idx, mail in enumerate(to_classify, 1):
            subject = mail.get("subject", "") or ""
            body = mail.get("body_text", "") or ""
            attachments = mail.get("attachments", [])

            if mail.get("hard_filter"):
                mail["intent_llm_status"] = "skipped_hard_filter"
                still_filtered.append(mail)
                self._log(f"LLM跳过硬过滤邮件: {subject[:40]}")
                if progress_callback:
                    progress_callback(idx, total)
                continue

            self._log(f"LLM 分类 {idx}/{total}: {subject[:50]}")
            result = llm_client.classify_email(subject, body, attachments)

            if result is None:
                # API/JSON/Schema 任一失败都不能等价于“不是目标邮件”。保留原邮件
                # 并强制进入人工复核，避免偶发模型故障造成漏单。
                mail["filter_status"] = "uncertain"
                mail["filter_reason"] = (
                    f"{mail.get('filter_reason','')} | LLM意图识别失败，已转人工复核"
                )
                mail["intent_llm_status"] = "manual_review"
                recovered.append(mail)
                self._log(f"LLM 分类失败转人工: {subject[:40]}", "warning")
            elif result.get("is_target"):
                intent = result.get("intent", "")
                reason = result.get("reason", "")
                mail["filter_status"] = "valid"
                mail["filter_reason"] = f"LLM恢复(意图={intent}, 原因={reason})"
                mail["intent_llm_status"] = "accepted_target"
                recovered.append(mail)
                self._log(f"LLM 恢复: {subject[:40]} → {intent}")
            elif result.get("confidence") == "low":
                # 全量识别仍无法确定时，不能把低置信度的 false 当作确定过滤，
                # 否则模型的“我不确定”会直接造成漏单。转人工队列保留原邮件。
                reason = result.get("reason", "")
                mail["filter_status"] = "uncertain"
                mail["filter_reason"] = (
                    f"{mail.get('filter_reason','')} | LLM低置信度，转人工复核({reason})"
                )
                mail["intent_llm_status"] = "manual_review"
                recovered.append(mail)
                self._log(f"LLM 低置信度转人工: {subject[:40]}", "warning")
            else:
                email_type = result.get("email_type", "")
                reason = result.get("reason", "")
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = f"{mail.get('filter_reason','')} | LLM确认={email_type}({reason})"
                mail["intent_llm_status"] = "accepted_non_target"
                still_filtered.append(mail)
                self._log(f"LLM 确认过滤: {subject[:40]} → {email_type}")

            if progress_callback:
                progress_callback(idx, total)

        self._log(f"LLM 全量意图识别完成: 保留 {len(recovered)} 封, 确认过滤 {len(still_filtered)} 封")
        return recovered, still_filtered
