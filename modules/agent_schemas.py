"""LLM Agent 的严格结构化输出协议。

Pydantic 只负责验证模型输出的结构、类型和枚举值；公司名称是否真实、
项目关系是否合理等业务语义由后续确定性规则和复检 Agent 处理。
"""
from typing import List, Literal, Optional

from pydantic import (
    AliasChoices,
    BaseModel,
    ConfigDict,
    Field,
    ValidationInfo,
    model_validator,
)


Confidence = Literal["high", "medium", "low"]
Intent = Literal["注册", "新增", "撤单", ""]
EmailType = Literal["注册类", "账单", "合同", "咨询", "下证", "购买", "其他"]
ReviewStatus = Literal["valid", "uncertain", "invalid"]
ReviewField = Literal["company", "agent", "country", "program", "request", "other"]
ISSUE_CODES = {
    "COMPANY_PLACEHOLDER", "COMPANY_CODE_MIXED", "COMPANY_SOURCE_CONFLICT",
    "COMPANY_EMAIL", "COMPANY_FILENAME", "COMPANY_CONTROL_CHAR", "COMPANY_SENTENCE",
    "PROJECT_COVERAGE_MISSING", "PROJECT_UNSUPPORTED", "ENTITY_PROJECT_MAPPING",
    "AGENT_UNMATCHED", "INTENT_UNCERTAIN", "FIELD_MISSING", "AGENT_FAILED", "OTHER",
}


class StrictAgentModel(BaseModel):
    """所有 Agent 输出共用的拒绝式协议。"""

    model_config = ConfigDict(
        extra="forbid",
        strict=True,
        str_strip_whitespace=True,
        populate_by_name=True,
    )


class PreflightOutput(StrictAgentModel):
    ready: bool


class IntentOutput(StrictAgentModel):
    is_target: bool
    intent: Intent
    email_type: EmailType
    confidence: Confidence
    reason: str


class ExtractionRecord(StrictAgentModel):
    company: str = Field(alias="客户")
    projects: List[str] = Field(alias="项目")
    confidence: Confidence


class ExtractionOutput(StrictAgentModel):
    agent: str = Field(alias="代理")
    company: str = Field(alias="客户")
    customer_code: str = Field(default="", alias="客户编号")
    projects: List[str] = Field(alias="项目")
    records: List[ExtractionRecord] = Field(
        validation_alias=AliasChoices("记录", "records"),
        serialization_alias="记录",
    )
    request: Intent = Field(alias="需求")
    confidence: Confidence


class ReviewIssue(StrictAgentModel):
    # code 允许旧模型省略（缺省 OTHER），新 Prompt 会要求输出标准问题编号。
    # 其余字段用于工作台展示“为什么错、依据是什么、建议改成什么”。
    code: str = Field(default="OTHER", max_length=64)
    field: ReviewField
    reason: str = Field(default="", max_length=240)
    evidence: str = Field(default="", max_length=500)
    current_value: str = Field(default="", max_length=300)
    suggested_value: str = Field(default="", max_length=300)
    suggestion: str = Field(default="", max_length=120)


class ReviewSuggestions(StrictAgentModel):
    company: str = ""
    agent: str = ""
    country: str = ""
    program: str = ""
    request: str = ""


class ReviewOutput(StrictAgentModel):
    status: ReviewStatus
    confidence: Confidence
    issues: List[ReviewIssue]
    suggestions: ReviewSuggestions
    reason: str
    selected_company_candidate_id: Optional[str] = Field(default=None, max_length=100)


class ReviewBatchItem(ReviewOutput):
    id: str = Field(min_length=1)


class ReviewBatchOutput(StrictAgentModel):
    results: List[ReviewBatchItem]

    @model_validator(mode="after")
    def validate_result_ids(self, info: ValidationInfo):
        """批量输出必须与本次请求 ID 一一对应，不能缺失、重复或越界。"""
        ids = [item.id for item in self.results]
        if len(ids) != len(set(ids)):
            raise ValueError("results contains duplicate ids")

        context = info.context or {}
        expected = context.get("expected_ids")
        if expected is not None and set(ids) != {str(x) for x in expected}:
            raise ValueError("results ids do not exactly match requested ids")
        return self
