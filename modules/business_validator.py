"""模型结构校验之后的确定性业务风险检查。"""
import re
from typing import Dict, List


_PLACEHOLDER_COMPANIES = {
    "一家公司", "一个公司", "一家企业", "某公司", "公司", "客户", "客户公司",
    "公司a", "公司b", "companya", "companyb", "unknown", "未知公司",
    "待确认", "暂无", "不详", "无", "非中国公司", "非中国企业",
    "中国公司", "外国公司", "国公司", "不能与其它公司", "不能与其他公司",
    "公司中文名称", "公司英文名称", "中文公司名称", "英文公司名称",
    "公司中文名", "公司英文名", "companyname", "companynameinchinese", "companynameinenglish",
}
_SENTENCE_PREFIXES = (
    "适用于", "適用於", "请提供", "請提供", "请查收", "請查收",
    "以下公司", "上述公司", "公司名称", "公司名稱", "客户名称", "客戶名稱",
    "在其他欧盟国家或第三国设立", "在其他欧盟国家设立", "在第三国设立",
)
_FILE_SUFFIX_RE = re.compile(r"\.(?:zip|rar|7z|pdf|docx?|xlsx?|jpe?g|png)$", re.I)
_PROJECT_TYPE_RE = re.compile(r"WEEE|电池法|電池法|包装法|包裝法|一次性塑料|EPR", re.I)
_COMPANY_COUNT_PLACEHOLDER_RE = re.compile(r"^[0-9一二三四五六七八九十多几]+\s*家\s*(?:公司|主体|企业)$", re.I)
_COMPANY_COUNT_PREFIX_RE = re.compile(
    r"^\s*[0-9一二三四五六七八九十多几]+\s*家\s*(?:公司|主体|企业)"
    r"(?:\s*[-—:：,，/／+＋].*)?$",
    re.I,
)
_COMPANY_PREFIX_POLLUTION_RE = re.compile(
    r"^\s*(?:(?:[\u4e00-\u9fff]+)?(?:WEEE|EEE|电池法|電池法|包装法|包裝法|EPR)\s*)?"
    r"(?:[A-Za-z]{1,3}\s*[-_]\s*)?[A-Za-z]{1,8}\s*[-_]?\s*\d{2,8}\b",
    re.I,
)
_ENGLISH_PERSON_NAME_RE = re.compile(
    r"^[A-Z][a-z]{1,24}(?:\s+[A-Z][a-z]{1,24}){1,3}$"
)
_CONTACT_OR_ID_RE = re.compile(r"^[+()\-\s\d]{6,}$")
_COMPANY_CODE_ONLY_RE = re.compile(r"^[A-Za-z]{1,8}[-_]?\d{4,}$")


def inspect_row(row: Dict) -> List[dict]:
    """只拦截高确定性的明显异常，不猜测或改写业务字段。"""
    issues: List[dict] = []
    def add_issue(field: str, code: str, reason: str) -> None:
        issues.append({"field": field, "code": code, "reason": reason})

    company = str(row.get("客户", "") or "").strip()
    normalized_company = re.sub(r"[\s\-_]+", "", company).lower()

    if not company:
        add_issue("company", "FIELD_MISSING", "客户公司名称为空")
    elif (
        normalized_company in _PLACEHOLDER_COMPANIES
        or _COMPANY_COUNT_PLACEHOLDER_RE.fullmatch(company)
        or _COMPANY_COUNT_PREFIX_RE.fullmatch(company)
    ):
        add_issue("company", "COMPANY_PLACEHOLDER", "客户字段是占位词，不是真实公司名称")
    elif "@" in company:
        add_issue("company", "COMPANY_EMAIL", "客户字段看起来是邮箱地址")
    elif _CONTACT_OR_ID_RE.fullmatch(company):
        add_issue("company", "COMPANY_CONTACT_OR_ID", "客户字段是电话或证件号，不是公司名称")
    elif _COMPANY_CODE_ONLY_RE.fullmatch(company):
        add_issue("company", "COMPANY_CODE_ONLY", "客户字段是客户编号/注册号，不是公司名称")
    elif _ENGLISH_PERSON_NAME_RE.fullmatch(company) and not re.search(
        r"\b(?:inc|incorporated|ltd|limited|llc|llp|plc|corp|corporation|gmbh|ug|srl|bv|co)\.?$", company, re.I
    ):
        add_issue("company", "COMPANY_PERSON_NAME", "客户字段疑似法人/联系人姓名，不是公司名称")
    elif _FILE_SUFFIX_RE.search(company):
        add_issue("company", "COMPANY_FILENAME", "客户字段看起来是附件文件名")
    elif any(ch in company for ch in "\r\n\t"):
        add_issue("company", "COMPANY_CONTROL_CHAR", "客户字段包含换行或制表符")
    elif company.startswith(_SENTENCE_PREFIXES):
        add_issue("company", "COMPANY_SENTENCE", "客户字段看起来是说明句而非公司名称")
    elif _COMPANY_PREFIX_POLLUTION_RE.search(company):
        add_issue("company", "COMPANY_CODE_MIXED", "客户字段混入项目/证书类型或客户编号")
    elif re.search(r"(?:不能|不可|不得|不要).{0,8}(?:其它|其他)?公司", company):
        add_issue("company", "COMPANY_FORM_NOTE", "客户字段是表单说明句，不是真实公司名称")

    project = str(row.get("项目", "") or "").strip()
    if not project:
        add_issue("program", "FIELD_MISSING", "服务项目为空")
    elif not _PROJECT_TYPE_RE.search(project):
        add_issue("program", "PROJECT_UNSUPPORTED", "服务项目未包含可识别的业务类型")

    request = str(row.get("需求", "") or "").strip()
    if request not in {"注册", "新增", "撤单"}:
        add_issue("request", "INTENT_UNCERTAIN", "需求不是注册、新增或撤单中的明确动作")

    return issues


def apply_row_guard(row: Dict) -> List[dict]:
    """标记人工复核，但绝不自动改写原始提取值。"""
    issues = inspect_row(row)
    if not issues:
        row["业务规则校验状态"] = "通过"
        row["业务规则问题"] = ""
        row["业务规则问题编号"] = ""
        return []

    row["业务规则校验状态"] = "需人工确认"
    row["业务规则问题"] = "；".join(
        f"{item['field']}: {item['reason']}" for item in issues
    )
    row["业务规则问题编号"] = "、".join(dict.fromkeys(
        str(item.get("code", "OTHER")) for item in issues
    ))
    row["置信度"] = "需人工确认"
    existing = str(row.get("人工复核提示", "") or "").strip()
    hint = "确定性业务校验: " + row["业务规则问题"]
    row["人工复核提示"] = "；".join(x for x in (existing, hint) if x)
    return issues
