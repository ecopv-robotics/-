"""LLM 意图识别模块 — DeepSeek API 辅助邮件分类与字段提取

两个核心功能:
1. classify_email: 对每封拉取到的邮件做全量意图识别，判断是否为新注册询单
2. extract_fields_llm: 对规则提取失败的字段做 LLM 补充提取
"""
import json
import os
import re
import time
from datetime import datetime
from typing import Dict, List, Optional, Type
from urllib.parse import urlparse

from openai import OpenAI
from pydantic import BaseModel, ValidationError
from utils.runtime_paths import APP_ROOT

from modules.agent_schemas import (
    ExtractionOutput,
    ISSUE_CODES,
    IntentOutput,
    PreflightOutput,
    ReviewBatchOutput,
    ReviewOutput,
)
from modules.agent_workflow import AgentWorkflow


SYSTEM_PROMPT_CLASSIFY = """你是一个邮件分类助手，服务于一家处理欧盟环保合规注册的公司（业务包括 WEEE、电池法、包装法、EPR、一次性塑料等）。

你的任务是判断邮件是否属于以下类型之一：
- "注册类"：客户/代理发来要求注册、新增、撤销环保合规项目
- "账单"：关于费用账单、发票
- "合同"：合同相关
- "咨询"：业务咨询、问题反馈
- "下证"：证书下发通知
- "购买"：购买担保、购买回收费
- "其他"：不属于以上类型

判断标准：
- 邮件明确要求新注册、新增、撤单环保合规项目（WEEE/电池法/包装法/EPR等）→ 注册类，is_target=true
- 仅修改已注册主体的公司名称、公司地址、证书抬头、注册地址或其他资料，哪怕同时出现 WEEE/电池法/注册 → 其他或下证，is_target=false
- 证书/下号/完成通知、续费、报价、付款、发票、合同、咨询、资料补交、状态查询 → 对应非注册类型，is_target=false
- 同一封邮件同时包含“新注册项目”和“公司改名/地址变更”时，只要正文明确存在新的注册项目，is_target=true；reason 中说明新项目和存量变更分别是什么
- 主题与正文冲突时，以正文中明确的实际办理意图为准；不要因为主题含有“注册”或项目名称就直接判为目标邮件
- 主题中的公司改名、地址变更、证书抬头变更、主体信息维护，如果正文没有明确的新项目注册，必须判为非目标；
  如果同一封邮件同时出现新项目和存量变更，只保留新项目意图，并在原因中分开说明
- 正文、附件名称或附件文本中明确的办理动作优先于主题；只有主题出现业务词而正文没有办理证据时，不得凭项目词猜测为注册
- 无法确定时 confidence=low，并将 is_target=false，reason 中明确写“需人工复核”，不能猜测为新询单
- 邮件提到"账单""invoice""费用"→ 账单
- 邮件提到"合同""协议"→ 合同
- 邮件提到"咨询""问题""反馈"→ 咨询
- 邮件提到"下证""证书""下号"→ 下证
- 邮件提到"购买""担保""回收费"→ 购买
- 其他情况 → 其他

你必须只返回 JSON，不要输出任何其他文字，也不要增加约定以外的字段。
reason 只写一句简短依据，最多 80 个汉字，不要复述整封邮件。返回结构：
{"is_target":true,"intent":"注册|新增|撤单|","email_type":"注册类|账单|合同|咨询|下证|购买|其他","confidence":"high|medium|low","reason":"简要原因"}
注意：布尔值必须使用 JSON 的 true/false，不能使用字符串。"""


SYSTEM_PROMPT_EXTRACT = """你是一个信息提取助手，服务于一家处理欧盟环保合规注册的公司。

从邮件内容中提取以下字段。邮件可能同时包含多家公司和多个项目，绝不能只返回第一家公司或第一个项目。
1. "代理"：处理注册的中间代理机构名称（不是客户公司名，是代理/中介公司名）。如果邮件中没有明确提到代理名称，返回空字符串。
2. "客户"：单客户兼容字段；如果有多家公司，返回第一家公司，同时必须把全部公司写入“记录”。
3. "客户编号"：EG3164、BG1945 等客户/案件编号；没有明确编号时返回空字符串，不能把编号拼入公司名称。
4. "项目"：邮件中提到的全部环保合规项目名称数组。标准格式为“国家+项目类型”，例如“德国WEEE”“法国电池法”“波兰包装法”“荷兰包装法”等。组合项目要拆开，例如“德国WEEE+电池法”应拆为["德国WEEE","德国电池法"]，注意根据上下文补齐国家。
5. "记录"：公司与项目的对应关系数组，每家公司一条，格式为[{"客户":"公司A","项目":["德国WEEE"],"confidence":"high/medium/low"}]。如果邮件只明确了一个公司，仍返回一条记录；如果同一批项目无法确定属于哪家公司，为每家公司保留全部可能项目，并将 confidence 设为“low”，不要静默漏掉。
6. "需求"：邮件意图类型，"注册"/"新增"/"撤单"之一。如果无法判断，返回空字符串。

实体拆分规则（必须遵守）：
- “德国WEEE”“法国电池法”等是证书/服务项目，不是公司名称；“EG3164”“BG1945”等是客户编号，也不是公司名称。
- 例如“德国WEEE EG0001 Example Trading LIMITED”应拆为：项目=德国WEEE，客户编号=EG0001，客户=Example Trading LIMITED；客户编号必须单独返回，不能塞进客户字段。
- 英文主体如“SED5632 xxxxxxxx B.V.”中，空格左侧的“SED5632”是客户/案件编号，真正公司名是右侧“xxxxxxxx B.V.”；输出客户时去掉该编号，并将编号单独放入“客户编号”。公司名通常不应只是数字或编号，看到“一家公司/2家公司”等数量说明时客户必须留空。
- 正文明确写出的公司名称优先于主题；主题中的代理名、证书类型和编号不能污染客户字段。
- 公司名可能同时出现在主题、正文和附件（尤其是 xlsx 表格）中；必须用至少两个独立来源互相印证。来源冲突或只能看到占位词时不要猜，客户返回空并交人工复核。
- 保护括号内的规格文本，不要把“（大型设备+小型设备）”拆成两个项目或两个公司。
- 不能确定公司与项目对应关系时，不要编造；在“记录”中保留候选关系并将 confidence 设为 low，确保项目不漏提。
- 客户字段只允许填写邮件/附件中明确出现的公司或法人主体名称；“POA法人职务/Legal positions”、
  “公司注册资金/Registration Capital”、“签字地点/Place of signature”、“签字时间/Signing time”、
  “Name of legal person”、地址、电话、邮箱、店铺链接、表单说明和表头标签都不是公司名。
- “东莞市虎门钦泓阁五金经营部（个体工商户）”这类个体工商户是合法客户主体，即使没有“有限公司”后缀也要保留完整名称；“非中国公司”只是 VAT 表单说明，不是客户名称。
- “不能与其它公司申请EPR的邮箱一致”“中国公司/外国公司/非中国公司”等是表单说明或类别文字，不是客户公司；看到这类值必须返回空并标记人工复核。
- 如果正文、主题和附件都没有清晰的公司名称，或只能看到上述字段标签，客户必须返回空字符串，
  不要根据联系人姓名、法人姓名、代理名、文件名或上下文猜测；宁可留空并交人工复核。
- 所有建议和不确定说明只写简短关键词，不超过 80 个汉字。

只返回合法 JSON，不要输出任何其他文字，也不要增加约定以外的字段。
返回结构：
{"代理":"","客户":"","客户编号":"","项目":[],"记录":[{"客户":"","项目":[],"confidence":"high|medium|low"}],"需求":"注册|新增|撤单|","confidence":"high|medium|low"}
注意：“项目”和“记录”必须始终是 JSON 数组；没有值时返回空数组。"""


SYSTEM_PROMPT_VALIDATE = """你是一个邮件询单字段质量审核助手，服务于欧盟环保合规注册团队。

请审核程序从一封邮件中抽取出的字段是否可信，重点检查：
1. 客户公司：是否像真实公司/法人主体名称，而不是人名、邮箱、部门、地址、完整句子、文件名残片或乱码；个体工商户（如“经营部/商店（个体工商户）”）也是合法主体；主题和正文冲突时，以正文明确写出的客户名称为优先；若提供“正文原文”和“附件证据”，还要与附件表格/单元格中的公司名互相印证。
2. 代理：是否像代理/中介机构名称；为空可以判为 uncertain，不要猜测。
3. 国家与服务项目：国家是否属于项目所在国家，WEEE/电池法/包装法/EPR 等项目是否自洽，组合项目是否漏拆。
4. 需求：是否为注册/新增/撤单；为空或与正文不符要标记。

复核优先级：正文明确内容 > 附件正文/附件名称 > 主题 > 历史参考。历史参考只能辅助，不能覆盖当前邮件证据。
特别检查：证书/服务项目、EG/BG/SED 等客户编号、法定公司名称必须分开；例如“SED5632 xxxxxxxx B.V.”应检查为编号=SED5632、公司=xxxxxxxx B.V.；“一家公司”“2家公司”“非中国公司”“不能与其它公司”、文件名、说明句不是合法公司名。公司名可能在标题、正文、附件表格中重复出现，优先使用正文/附件的真实主体并标记来源冲突。

每个问题必须选择一个问题编号：
COMPANY_PLACEHOLDER（占位公司）、COMPANY_CODE_MIXED（公司混入编号/项目）、
COMPANY_SOURCE_CONFLICT（主题正文公司冲突）、PROJECT_COVERAGE_MISSING（项目疑似漏提）、
ENTITY_PROJECT_MAPPING（多公司多项目关系不明）、AGENT_UNMATCHED（代理未命中）、
INTENT_UNCERTAIN（意图不确定）、FIELD_MISSING（字段缺失）、OTHER（其他）。
建议必须简短，最多 80 个汉字；只给出一个明确动作，不要写长篇分析。

不要因为名称是中文、英文、繁体或含有公司后缀差异就判错。不要根据常识编造公司名称。
只能返回合法 JSON，不要增加约定以外的字段：
{
  "status": "valid|uncertain|invalid",
  "confidence": "high|medium|low",
  "issues": [{"code":"问题编号","field": "company|agent|country|program|request|other","reason": "问题（不超过 120 字）","evidence":"原文短证据","current_value":"程序当前值","suggested_value":"建议值","suggestion": "一句话建议（不超过 80 字）"}],
  "suggestions": {"company": "", "agent": "", "country": "", "program": "", "request": ""},
  "reason": "总体判断"
}
"""

SYSTEM_PROMPT_VALIDATE_BATCH = SYSTEM_PROMPT_VALIDATE + """

本次输入包含多条记录。请逐条审核，必须保留每条记录的 id，不要合并记录。

重要：一条记录只承载一个服务项目，同一封邮件会被拆成多行，所以每行都带有
"同封邮件全部项目"（该封邮件拆分出的完整项目清单，含本行）。
判定第 3 条"组合项目是否漏拆"时以此为准：
- 若本行的服务项目已出现在"同封邮件全部项目"中，说明它只是拆分结果的一部分，
  不得因为"只抽到一个项目"判为漏拆；
- 只有当正文明确提到的项目不在“同封邮件全部项目”里时，才算真的漏拆。
- 没有“同封邮件全部项目”字段时，按普通单条记录审核。

输入中可能附带“正文原文”和“附件证据”（含 xlsx 结构化行/单元格快照）。
公司名若在附件中出现，应优先使用附件中明确的法定主体，并与主题、正文交叉核对；
附件证据缺失或来源冲突时不要猜测，标记 COMPANY_SOURCE_CONFLICT/uncertain。
若输入含“公司候选（仅可选已有ID）”，请仅在给定候选 ID 中选择一个最符合该申请主体的值，
并把对应 ID 写入 selected_company_candidate_id；不得改写或输出新公司名。证据冲突、无法判定时留空，
将 status 设为 uncertain。候选来自同一申请表字段时，结合字段标签、表格值及邮件证据综合判断。

只返回：{"results":[{"id":"原id","status":"valid|uncertain|invalid","confidence":"high|medium|low","issues":[{"code":"问题编号","field":"company|agent|country|program|request|other","reason":"问题","evidence":"原文短证据","current_value":"程序当前值","suggested_value":"建议值","suggestion":"一句话建议"}],"suggestions":{"company":"","agent":"","country":"","program":"","request":""},"selected_company_candidate_id":"候选ID或空字符串","reason":"一句话总体判断"}]}。
"""


SYSTEM_PROMPT_WEEE_REVIEW = """你是德国 ElektroG/WEEE 品牌与品类字段复检助手。

请复核程序根据现有规则从邮件标题、正文和附件表格中提取出的品牌/原始品类项目。
只检查字段提取，不负责决定德国第1至第6类；类别编号和分类映射由独立规则及人工确认完成。

复检要求：
1. 每个品牌和原始品类必须能在当前邮件或附件证据中找到，不能根据常识补写。
2. 检查品牌与品类是否错配、漏拆、重复，是否把公司名、代理名、表头、示例、数量、尺寸或模板说明当成字段。
3. 附件表格中的“品牌/类别”结构化行优先；正文/标题只有明确字段或明确的“第几类”才可作为补充。
4. 没有明确证据时标记 uncertain 或 invalid，不要猜测。
5. 初步项目为空但原文明确出现品牌/类别时，必须标记 invalid，并指出需要重新提取。

只返回 JSON，不要 Markdown 或额外文字：
{"status":"valid|uncertain|invalid","confidence":"high|medium|low",
"issues":[{"code":"MISSING|MISPAIRED|DUPLICATE|TEMPLATE|UNSUPPORTED|OTHER",
"field":"brand|category|pair|other","reason":"不超过120字",
"evidence":"原文短证据","current_value":"当前值","suggestion":"不超过80字"}],
"reason":"不超过120字"}
"""


SYSTEM_PROMPT_WEEE_REEXTRACT = """你是德国 ElektroG/WEEE 品牌与原始品类字段提取助手。

请参考程序的既有提取规则，对邮件标题、正文和附件表格重新提取全部真实出现的品牌/原始品类项目。

提取规则：
1. 附件表格中明确的“品牌/品牌名称/Brand/Marke”和“品类/类别/产品类别/产品分类/Category/Product Type”列优先。
2. 邮件正文或标题只读取明确字段（例如“品牌：X”“品类：Y”“新增品牌X”“第六类”）；普通句子中的公司名、代理名和项目名不能当品牌。
3. 每个品牌与品类按同一行、同一字段组配对；不能跨公司或跨行猜测。
4. 忽略表头、示例、说明、占位词、数量、重量、尺寸和空值。例如 Apple、Li、10g、3500个若只在“例”中出现，不是业务项目。
5. “小型设备”和“单边尺寸小于50cm”应作为同一条原始品类；孤立尺寸不能单独成项。
6. 只返回原文/附件中真实出现的文字，不要翻译、标准化或根据产品常识补写品牌和品类。
7. 如果无法确认，保留原始文字并降低 confidence；没有证据的字段留空。

只返回 JSON，不要 Markdown 或额外文字：
{"items":[{"item_id":"原项目编号或序号","brand":"","category":"",
"source":"邮件正文/标题或附件表格来源","evidence":"原文短证据",
"confidence":"high|medium|low"}],"reason":"不超过120字"}
"""


class LLMIntentClient:
    """OpenAI 兼容 LLM 客户端（支持百炼千问等服务）。"""

    def __init__(self, config: dict, logger=None):
        """
        config: {
            "api_key": str,
            "base_url": str,
            "model": str,
            "timeout": int,
            "max_tokens": int,
        }
        """
        # 配置 api_key_env 时只接受该环境变量，避免误把另一家服务商的旧 Key
        # 发往新端点。未配置 api_key_env 的旧部署仍保留 config.api_key 兼容逻辑。
        self.api_key_env = str(config.get("api_key_env", "")).strip()
        env_api_key = os.getenv(self.api_key_env, "") if self.api_key_env else ""
        self.api_key = env_api_key if self.api_key_env else config.get("api_key", "")
        self.credential_source = (
            f"env:{self.api_key_env}" if env_api_key else
            ("config.api_key" if self.api_key else "未配置")
        )
        self.base_url = config.get("base_url", "https://api.deepseek.com")
        self.model = config.get("model", "deepseek-chat")
        self.timeout = config.get("timeout", 30)
        self.max_tokens = config.get("max_tokens", 1000)
        self.temperature = config.get("temperature", 0.1)
        self.thinking_mode = str(config.get("thinking_mode", "")).strip().lower()
        self.enabled = bool(self.api_key)
        self.logger = logger
        self.endpoint_host = urlparse(self.base_url).netloc or "unknown"
        self.usage = {
            "requests_started": 0,
            "requests_succeeded": 0,
            "requests_failed": 0,
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "total_tokens": 0,
            "output_validation_failures": 0,
            "repair_attempts": 0,
            "repair_successes": 0,
            "manual_review_fallbacks": 0,
        }
        default_log_dir = os.path.join(
            str(APP_ROOT), "logs"
        )
        self.audit_log_path = config.get("audit_log_path") or os.path.join(
            default_log_dir, "llm_calls.jsonl"
        )

        if self.enabled:
            self.client = OpenAI(
                api_key=self.api_key,
                base_url=self.base_url,
                timeout=self.timeout,
            )
        else:
            self.client = None
        self.agent_workflow = AgentWorkflow(self)

    def _log(self, msg, level="info"):
        # 兼容离线测试/轻量构造的客户端对象：这类对象可能没有 logger 属性。
        logger = getattr(self, "logger", None)
        if logger:
            getattr(logger, level)(msg)

    @staticmethod
    def _agent_label(purpose: str) -> str:
        """将内部 purpose 映射为日志中便于排查的 Agent 名称。"""
        value = str(purpose or "")
        if "startup_preflight" in value:
            return "LLM连通性预检"
        if "intent" in value:
            return "意图识别Agent"
        if "field_extraction" in value:
            return "字段抽取Agent"
        if "semantic_field_validation" in value:
            return "语义复检Agent"
        return "LLMAgent"

    def _write_audit_event(self, event: dict) -> None:
        """写入无内容、无密钥的本地调用审计记录；写失败不影响主流程。"""
        safe_event = {
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "endpoint_host": self.endpoint_host,
            "configured_model": self.model,
            **event,
        }
        try:
            audit_dir = os.path.dirname(self.audit_log_path)
            if audit_dir:
                os.makedirs(audit_dir, exist_ok=True)
            with open(self.audit_log_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(safe_event, ensure_ascii=False) + "\n")
        except OSError as e:
            self._log(f"LLM 审计日志写入失败: {type(e).__name__}", "warning")

    @staticmethod
    def _usage_values(response) -> dict:
        usage = getattr(response, "usage", None)
        return {
            "prompt_tokens": int(getattr(usage, "prompt_tokens", 0) or 0),
            "completion_tokens": int(getattr(usage, "completion_tokens", 0) or 0),
            "total_tokens": int(getattr(usage, "total_tokens", 0) or 0),
        }

    def get_usage_summary(self) -> dict:
        """返回本轮的可展示用统计；不包含邮件内容、密钥或提示词。"""
        return {
            **self.usage,
            "configured_model": self.model,
            "endpoint_host": self.endpoint_host,
            "credential_source": self.credential_source,
            "audit_log_path": self.audit_log_path,
        }

    @staticmethod
    def _loads_lenient(raw: str) -> dict:
        """尽力把模型输出解析成 dict。

        模型即使被要求只返回 JSON，也常会带上 ```json 围栏、前后解释文字或
        多余的尾逗号；直接 json.loads 会整条丢弃（此前一轮就出现过这种失败）。
        这里依次尝试原始文本 → 围栏内文本 → 第一个花括号块 → 去掉尾逗号。
        """
        text = (raw or "").strip()
        if not text:
            raise json.JSONDecodeError("empty response", text, 0)
        fenced = re.search(r"```(?:json)?\s*(.+?)```", text, flags=re.S | re.I)
        if fenced:
            text = fenced.group(1).strip()
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            start, end = text.find("{"), text.rfind("}")
            if start == -1 or end <= start:
                raise
            snippet = text[start:end + 1]
            try:
                return json.loads(snippet)
            except json.JSONDecodeError:
                return json.loads(re.sub(r",\s*([}\]])", r"\1", snippet))

    @staticmethod
    def _validation_error_summary(error: Exception) -> str:
        """仅保留字段位置和错误类型，防止正文/公司名进入审计日志。"""
        if isinstance(error, ValidationError):
            parts = []
            for item in error.errors(include_input=False, include_url=False):
                location = ".".join(str(x) for x in item.get("loc", ())) or "root"
                parts.append(f"{location}: {item.get('type', 'validation_error')}")
            return "；".join(parts[:20])
        return type(error).__name__

    def _get_agent_workflow(self) -> AgentWorkflow:
        """兼容测试中通过 ``__new__`` 构造的轻量客户端。"""
        workflow = getattr(self, "agent_workflow", None)
        if workflow is None:
            workflow = AgentWorkflow(self)
            self.agent_workflow = workflow
        return workflow

    def _call_api(self, system_prompt: str, user_content: str,
                  purpose: str, max_tokens: Optional[int] = None,
                  allow_retry: bool = True,
                  output_model: Optional[Type[BaseModel]] = None,
                  validation_context: Optional[dict] = None) -> Optional[dict]:
        """调用 API 并执行严格输出校验；失败时最多请求模型修复一次。"""
        if not self.enabled:
            return None

        self.usage["requests_started"] += 1
        started = time.perf_counter()
        raw = ""
        # 即使模型输出未通过 JSON/Pydantic 校验，也要把本次响应的
        # prompt/completion/total usage 写入审计日志，保证逐条审计与总账一致。
        usage = {
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "total_tokens": 0,
        }
        agent_label = self._agent_label(purpose)
        self._log(
            f"[Agent状态] {agent_label} 开始调用: purpose={purpose}, model={self.model}"
        )
        try:
            request_kwargs = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_content},
                ],
                "response_format": {"type": "json_object"},
                "temperature": self.temperature,
                "max_tokens": max_tokens or self.max_tokens,
            }
            # DeepSeek Chat Completions 的 thinking 参数需通过 extra_body 传入。
            # 对未配置该项的其他 OpenAI 兼容服务不发送此厂商扩展参数。
            if self.thinking_mode in {"enabled", "disabled"}:
                request_kwargs["extra_body"] = {
                    "thinking": {"type": self.thinking_mode}
                }
            response = self.client.chat.completions.create(
                **request_kwargs,
            )
            raw = response.choices[0].message.content.strip()
            usage = self._usage_values(response)
            for key, value in usage.items():
                self.usage[key] += value
            if not raw:
                self.usage["requests_failed"] += 1
                self._write_audit_event({
                    "purpose": purpose,
                    "status": "empty_response",
                    "duration_ms": round((time.perf_counter() - started) * 1000),
                    "response_model": getattr(response, "model", "") or self.model,
                    "request_id": getattr(response, "_request_id", "") or "",
                    **usage,
                })
                self._log(
                    f"[Agent状态] {agent_label} 返回空内容: purpose={purpose}",
                    "warning",
                )
                return None
            result = self._loads_lenient(raw)
            if output_model is not None:
                validated = output_model.model_validate(
                    result,
                    context=validation_context or {},
                )
                result = validated.model_dump(by_alias=True)
            self.usage["requests_succeeded"] += 1
            event = {
                "purpose": purpose,
                "status": "success",
                "duration_ms": round((time.perf_counter() - started) * 1000),
                "response_model": getattr(response, "model", "") or self.model,
                "request_id": getattr(response, "_request_id", "") or "",
                **usage,
            }
            self._write_audit_event(event)
            self._log(
                f"[Agent状态] {agent_label} 调用成功: "
                f"purpose={purpose}, model={event['response_model']}, "
                f"tokens={usage['total_tokens']}, duration_ms={event['duration_ms']}, "
                f"request_id={event['request_id'] or '未返回'}"
            )
            return result
        except (json.JSONDecodeError, ValidationError) as e:
            self.usage["requests_failed"] += 1
            self.usage["output_validation_failures"] += 1
            status = "invalid_json" if isinstance(e, json.JSONDecodeError) else "schema_validation_error"
            error_summary = self._validation_error_summary(e)
            self._write_audit_event({
                "purpose": purpose,
                "status": status,
                "duration_ms": round((time.perf_counter() - started) * 1000),
                "error_type": type(e).__name__,
                "validation_error_count": len(error_summary.split("；")) if error_summary else 0,
                **usage,
            })
            self._log(
                f"[Agent状态] {agent_label} 输出校验失败: purpose={purpose}, status={status}, "
                f"errors={error_summary or type(e).__name__}；准备重试={allow_retry}",
                "warning",
            )
            if allow_retry:
                self.usage["repair_attempts"] += 1
                self._log(
                    f"[Agent状态] {agent_label} 开始结构化输出修复（第1/1次）",
                    "warning",
                )
                previous_output = raw[:6000] if raw else "<空输出或无法解析>"
                repair_content = (
                    "请重新完成下面的原始任务，并只返回符合约定的 JSON 对象。\n\n"
                    "【原始任务】\n" + user_content +
                    "\n\n【上一次不合格输出】\n" + previous_output +
                    "\n\n【结构校验问题】\n" + error_summary +
                    "\n\n不得输出解释、注释、Markdown 围栏或额外字段。"
                )
                repaired = self._call_api(
                    system_prompt,
                    repair_content,
                    purpose=f"{purpose}_output_repair",
                    max_tokens=max_tokens,
                    allow_retry=False,
                    output_model=output_model,
                    validation_context=validation_context,
                )
                if repaired is not None:
                    self.usage["repair_successes"] += 1
                    return repaired
            else:
                self.usage["manual_review_fallbacks"] += 1
            return None
        except Exception as e:
            self.usage["requests_failed"] += 1
            self._write_audit_event({
                "purpose": purpose,
                "status": "api_error",
                "duration_ms": round((time.perf_counter() - started) * 1000),
                "error_type": type(e).__name__,
                **usage,
            })
            self._log(f"[Agent状态] {agent_label} API调用失败: {e}", "error")
            return None

    def preflight(self) -> bool:
        """执行一次无邮件内容的连通性预检，失败时避免批量请求。"""
        if not self.enabled:
            return False
        result = self._call_api(
            "你是 API 连通性检测器，只能返回 JSON。",
            '返回 {"ready": true}。',
            purpose="startup_preflight",
            max_tokens=32,
            output_model=PreflightOutput,
        )
        if result and result.get("ready") is True:
            self._log(
                f"LLM 预检成功: model={self.model}, endpoint={self.endpoint_host}"
            )
            return True
        self.enabled = False
        self._log(
            "LLM 预检失败，已禁用本次运行的 LLM 调用；"
            "过滤候选将保留为人工复核项。",
            "warning",
        )
        return False

    @staticmethod
    def _build_email_content(subject: str, body: str, attachments: list) -> str:
        """构造发给 LLM 的邮件内容文本"""
        parts = [f"【邮件主题】\n{subject}"]
        if body:
            truncated_body = body[:2000] if len(body) > 2000 else body
            parts.append(f"【邮件正文】\n{truncated_body}")

        att_texts = []
        for att in attachments:
            fname = att.get("filename", "")
            text = att.get("text_content", "")
            if text:
                truncated = text[:1000] if len(text) > 1000 else text
                att_texts.append(f"[{fname}] {truncated}")
        if att_texts:
            parts.append(f"【附件内容】\n" + "\n".join(att_texts))

        return "\n\n".join(parts)

    def classify_email(
        self, subject: str, body: str, attachments: list
    ) -> Optional[dict]:
        """
        对单封邮件做意图分类
        返回: {
            "is_target": bool,      # 是否注册类邮件
            "intent": str,          # "注册"/"新增"/"撤单"/""
            "email_type": str,      # "注册类"/"账单"/"合同"/"咨询"/"下证"/"购买"/"其他"
            "reason": str,          # 简要原因
        }
        失败返回 None
        """
        self._log(
            f"[Agent任务] 意图识别Agent: 主题={str(subject or '')[:80]}, "
            f"附件数={len(attachments or [])}"
        )
        content = self._build_email_content(subject, body, attachments)
        step = self._get_agent_workflow().run(
            step="intent_classification",
            system_prompt=SYSTEM_PROMPT_CLASSIFY,
            user_content=(
                content + '\n\n请分析以上邮件并严格按约定结构返回 JSON。'
            ),
            output_model=IntentOutput,
        )
        if not step.accepted:
            self._log("[Agent状态] 意图识别Agent 结果未接受，转人工/规则兜底", "warning")
            return None
        self._log(
            f"[Agent状态] 意图识别Agent 结果已接受: "
            f"intent={step.payload.get('intent', '')}, "
            f"is_target={step.payload.get('is_target')}"
        )
        return step.payload

    def batch_classify(self, mails: List[dict]) -> Dict[str, dict]:
        """
        批量分类邮件
        mails: [{uid, subject, body_text, attachments, ...}, ...]
        返回: {uid: classify_result}
        """
        results = {}
        total = len(mails)
        for idx, mail in enumerate(mails, 1):
            uid = mail.get("uid", str(idx))
            subject = mail.get("subject", "") or ""
            body = mail.get("body_text", "") or ""
            attachments = mail.get("attachments", [])

            self._log(f"[Agent进度] 意图识别Agent {idx}/{total}: {subject[:40]}")
            result = self.classify_email(subject, body, attachments)
            if result:
                results[uid] = result
            else:
                results[uid] = {"is_target": None, "intent": "", "email_type": "LLM失败", "reason": "API调用失败"}

            time.sleep(0.5)

        return results

    def extract_fields_llm(
        self, subject: str, body: str, attachments: list, sender_email: str = ""
    ) -> Optional[dict]:
        """
        LLM 提取结构化字段
        返回: {
            "代理": str,
            "客户": str,
            "客户编号": str,
            "项目": [str, ...],
            "记录": [{"客户": str, "项目": [str, ...], "confidence": str}, ...],
            "需求": str,
            "confidence": str,
        }
        失败返回 None
        """
        self._log(
            f"[Agent任务] 字段抽取Agent: 发件人={sender_email or '未知'}, "
            f"主题={str(subject or '')[:80]}, 附件数={len(attachments or [])}"
        )
        content = self._build_email_content(subject, body, attachments)
        if sender_email:
            content = f"【发件人邮箱】{sender_email}\n\n" + content

        step = self._get_agent_workflow().run(
            step="field_extraction",
            system_prompt=SYSTEM_PROMPT_EXTRACT,
            user_content=content + '\n\n请提取以上邮件中的字段并严格按约定结构返回 JSON。',
            output_model=ExtractionOutput,
        )
        if not step.accepted:
            self._log("[Agent状态] 字段抽取Agent 结果未接受，保留规则值并转人工", "warning")
            return None
        self._log(
            f"[Agent状态] 字段抽取Agent 结果已接受: "
            f"records={len(step.payload.get('记录') or [])}, "
            f"projects={len(step.payload.get('项目') or [])}, "
            f"confidence={step.payload.get('confidence', '')}"
        )
        return step.payload

    def classify_weee_categories(
        self, items: List[dict], category_names: Optional[dict] = None
    ) -> List[dict]:
        """为规则无法唯一归类的德国 WEEE 项目提供候选类别。

        这是辅助建议，不是自动确认：调用方必须保留原始品牌/品类证据，
        并等待操作人员逐项选择后才可进入工单核对。返回值只接受 1--6 类。
        """
        if not self.enabled or not items:
            return []
        names = category_names or {
            "1": "热交换设备", "2": "屏幕和显示设备", "3": "灯具和光源",
            "4": "大型设备", "5": "小型设备", "6": "小型信息和电信设备",
        }
        payload = {
            "类别定义": names,
            "待判断项目": [
                {
                    "item_id": str(item.get("item_id") or index),
                    "品牌": str(item.get("brand") or ""),
                    "原始类别": str(item.get("category") or item.get("category_original") or ""),
                    "证据": str(item.get("evidence") or "")[:1200],
                }
                for index, item in enumerate(items)
                if isinstance(item, dict)
            ],
        }
        system_prompt = (
            "你是德国 ElektroG/WEEE 品类辅助分类器。只能依据给出的原始类别和证据，"
            "在第1至第6类中提出候选；不能把品牌、公司名、表头、示例文字、数量或尺寸"
            "单独当作类别。信息不足时 category_class 必须为空。只返回 JSON："
            '{"results":[{"item_id":"","category_class":"1-6或空","confidence":"high|medium|low",'
            '"reason":"不超过120字"}]}。不得输出 Markdown 或额外字段。'
        )
        result = self._call_api(
            system_prompt,
            json.dumps(payload, ensure_ascii=False) + "\n请只返回约定 JSON。",
            purpose="weee_category_suggestion",
            max_tokens=700,
        )
        if not isinstance(result, dict):
            return []
        raw_results = result.get("results")
        if not isinstance(raw_results, list):
            return []
        accepted: List[dict] = []
        for raw in raw_results:
            if not isinstance(raw, dict):
                continue
            category_class = str(raw.get("category_class") or "").strip()
            if category_class not in names:
                category_class = ""
            item_id = str(raw.get("item_id") or "").strip()
            if not item_id:
                continue
            confidence = str(raw.get("confidence") or "low").strip().lower()
            if confidence not in {"high", "medium", "low"}:
                confidence = "low"
            accepted.append({
                "item_id": item_id,
                "category_class": category_class,
                "category_class_name": names.get(category_class, ""),
                "confidence": confidence,
                "reason": str(raw.get("reason") or "").strip()[:240],
            })
        return accepted

    @staticmethod
    def _compact_weee_attachments(attachments: Optional[list]) -> List[dict]:
        """为 WEEE 复检/重提取构造有限大小的附件证据快照。

        不把附件二进制或绝对路径发送给模型，只保留文件名、结构化行和表格预览。
        这样既能复用既有字段规则，也避免把整本工作簿无界地塞进提示词。
        """
        compact: List[dict] = []
        for attachment in attachments or []:
            if not isinstance(attachment, dict):
                continue
            record = {
                "filename": str(attachment.get("filename") or "")[:240],
            }
            structured = []
            for row in (attachment.get("structured_records") or [])[:80]:
                if not isinstance(row, dict):
                    continue
                structured.append({
                    "sheet_name": str(row.get("sheet_name") or "")[:120],
                    "row_number": row.get("row_number") or "",
                    "brand": str(row.get("brand") or "")[:180],
                    "category": str(row.get("category") or "")[:240],
                    "raw_text": str(row.get("raw_text") or "")[:900],
                })
            if structured:
                record["structured_records"] = structured
            previews = []
            for sheet in (attachment.get("sheets") or [])[:20]:
                if not isinstance(sheet, dict):
                    continue
                rows = []
                for row in (sheet.get("preview_rows") or sheet.get("rows") or [])[:80]:
                    if isinstance(row, dict):
                        rows.append({
                            "row_number": row.get("row_number") or row.get("index") or "",
                            "cells": [str(cell or "")[:240] for cell in (row.get("cells") or row.get("values") or [])[:24]],
                        })
                    elif isinstance(row, (list, tuple)):
                        rows.append({"row_number": "", "cells": [str(cell or "")[:240] for cell in row[:24]]})
                if rows:
                    previews.append({"sheet_name": str(sheet.get("sheet_name") or "工作表")[:120], "rows": rows})
            if previews:
                record["sheets"] = previews
            text_content = str(attachment.get("text_content") or attachment.get("text") or "")
            if text_content:
                record["text_content"] = text_content[:5000]
            compact.append(record)
        return compact

    def review_weee_extraction(
        self,
        subject: str,
        body: str,
        attachments: Optional[list],
        project: str,
        items: List[dict],
    ) -> Optional[dict]:
        """让 LLM 复核规则提取结果，不直接改写结果。"""
        if not self.enabled:
            return None
        payload = {
            "项目": project or "",
            "邮件标题": subject or "",
            "邮件正文": (body or "")[:8000],
            "附件证据": self._compact_weee_attachments(attachments),
            "程序初步提取": [
                {
                    "item_id": str(item.get("item_id") or index),
                    "品牌": str(item.get("brand") or ""),
                    "原始品类": str(item.get("category") or item.get("category_original") or ""),
                    "来源": str(item.get("source") or ";".join(item.get("sources") or [])),
                    "证据": str(item.get("evidence") or ";".join(item.get("evidences") or ""))[:1200],
                }
                for index, item in enumerate(items or [])
                if isinstance(item, dict)
            ],
        }
        result = self._call_api(
            SYSTEM_PROMPT_WEEE_REVIEW,
            json.dumps(payload, ensure_ascii=False) + "\n请只返回约定 JSON。",
            purpose="weee_field_review",
            max_tokens=1800,
        )
        if not isinstance(result, dict):
            return None
        status = str(result.get("status") or "uncertain").strip().lower()
        if status not in {"valid", "uncertain", "invalid"}:
            status = "uncertain"
        confidence = str(result.get("confidence") or "low").strip().lower()
        if confidence not in {"high", "medium", "low"}:
            confidence = "low"
        issues = []
        for issue in result.get("issues") or []:
            if not isinstance(issue, dict):
                continue
            issues.append({
                "code": str(issue.get("code") or "OTHER").strip().upper()[:40],
                "field": str(issue.get("field") or "other").strip()[:30],
                "reason": str(issue.get("reason") or "").strip()[:240],
                "evidence": str(issue.get("evidence") or "").strip()[:500],
                "current_value": str(issue.get("current_value") or "").strip()[:240],
                "suggestion": str(issue.get("suggestion") or "").strip()[:160],
            })
        return {
            "status": status,
            "confidence": confidence,
            "issues": issues[:20],
            "reason": str(result.get("reason") or "").strip()[:240],
        }

    def extract_weee_fields(
        self,
        subject: str,
        body: str,
        attachments: Optional[list],
        project: str,
        initial_items: Optional[List[dict]] = None,
        review: Optional[dict] = None,
    ) -> Optional[dict]:
        """复检不通过时，按既有规则让 LLM 重新提取 WEEE 字段。"""
        if not self.enabled:
            return None
        payload = {
            "项目": project or "",
            "邮件标题": subject or "",
            "邮件正文": (body or "")[:10000],
            "附件证据": self._compact_weee_attachments(attachments),
            "程序初步提取": [
                {
                    "item_id": str(item.get("item_id") or index),
                    "品牌": str(item.get("brand") or ""),
                    "原始品类": str(item.get("category") or item.get("category_original") or ""),
                    "证据": str(item.get("evidence") or ";".join(item.get("evidences") or ""))[:1200],
                }
                for index, item in enumerate(initial_items or [])
                if isinstance(item, dict)
            ],
            "复检结果": review or {},
        }
        result = self._call_api(
            SYSTEM_PROMPT_WEEE_REEXTRACT,
            json.dumps(payload, ensure_ascii=False) + "\n请只返回约定 JSON。",
            purpose="weee_field_reextract",
            max_tokens=2400,
        )
        if not isinstance(result, dict) or not isinstance(result.get("items"), list):
            return None
        accepted: List[dict] = []
        for index, raw in enumerate(result.get("items") or []):
            if not isinstance(raw, dict):
                continue
            brand = str(raw.get("brand") or "").strip()[:180]
            category = str(raw.get("category") or "").strip()[:240]
            if not brand and not category:
                continue
            confidence = str(raw.get("confidence") or "low").strip().lower()
            if confidence not in {"high", "medium", "low"}:
                confidence = "low"
            accepted.append({
                "item_id": str(raw.get("item_id") or index).strip(),
                "brand": brand,
                "category": category,
                "source": str(raw.get("source") or "LLM重新提取").strip()[:240],
                "evidence": str(raw.get("evidence") or "").strip()[:1200],
                "confidence": confidence,
            })
        return {"items": accepted, "reason": str(result.get("reason") or "").strip()[:240]}

    def validate_extracted_fields(
        self,
        subject: str,
        body: str,
        fields: dict,
        sender_email: str = "",
        attachments: Optional[list] = None,
    ) -> Optional[dict]:
        """复核已抽取字段，不直接改写字段，只返回问题和建议。"""
        payload = {
            "发件人邮箱": sender_email,
            "主题": subject or "",
            "正文": (body or "")[:2500],
            "附件名称": [
                str(a.get("filename", "")) for a in (attachments or [])
                if isinstance(a, dict) and a.get("filename")
            ],
            "程序抽取字段": {
                "代理": fields.get("agent", ""),
                "客户公司": fields.get("company", ""),
                "国家": fields.get("country", ""),
                "服务项目": fields.get("program", ""),
                "需求": fields.get("request", ""),
            },
        }
        step = self._get_agent_workflow().run(
            step="semantic_field_validation",
            system_prompt=SYSTEM_PROMPT_VALIDATE,
            user_content=json.dumps(payload, ensure_ascii=False) + "\n\n请审核以上字段并只返回约定 JSON。",
            output_model=ReviewOutput,
            max_tokens=700,
        )
        if not step.accepted:
            return None
        return step.payload

    def validate_extracted_fields_batch(self, items: List[dict]) -> Dict[str, dict]:
        """批量复核多条抽取结果，减少逐行调用造成的延迟和 Token 成本。"""
        if not items:
            return {}
        payload = []
        for item in items:
            fields = item.get("fields") or {}
            record = {
                "id": str(item.get("id", "")),
                "发件人邮箱": item.get("sender_email", ""),
                "主题": item.get("subject", ""),
                "正文": (item.get("body", "") or "")[:1800],
                "附件名称": item.get("attachments", []) or [],
                "程序抽取字段": {
                    "代理": fields.get("agent", ""),
                    "客户公司": fields.get("company", ""),
                    "国家": fields.get("country", ""),
                    "服务项目": fields.get("program", ""),
                    "需求": fields.get("request", ""),
                },
            }
            if item.get("body_original"):
                record["正文原文"] = str(item.get("body_original"))[:6000]
            if item.get("attachment_evidence"):
                record["附件证据"] = str(item.get("attachment_evidence"))[:9000]
            company_candidates = item.get("company_candidates") or []
            if company_candidates:
                record["公司候选（仅可选已有ID）"] = [
                    {
                        "id": str(candidate.get("id", "")),
                        "value": str(candidate.get("value", "")),
                        "field_type": str(candidate.get("field_type", "")),
                        "source": str(candidate.get("source", "")),
                        "evidence": str(candidate.get("evidence", ""))[:300],
                    }
                    for candidate in company_candidates
                    if isinstance(candidate, dict) and candidate.get("id")
                ]
            same_mail = item.get("same_mail_programs")
            if same_mail:
                record["同封邮件全部项目"] = list(same_mail)
            rule_risks = item.get("rule_risks")
            if rule_risks:
                record["确定性规则风险提示"] = str(rule_risks)
            payload.append(record)
        expected_ids = [str(item.get("id", "")) for item in items]
        self._log(
            f"[Agent任务] 语义复检Agent: 批量={len(items)} 条, "
            f"ids={','.join(expected_ids[:12])}{'…' if len(expected_ids) > 12 else ''}"
        )
        step = self._get_agent_workflow().run(
            step="semantic_field_validation_batch",
            system_prompt=SYSTEM_PROMPT_VALIDATE_BATCH,
            user_content=json.dumps(payload, ensure_ascii=False) + "\n\n请逐条审核并返回 results 数组。",
            output_model=ReviewBatchOutput,
            max_tokens=min(4000, max(900, 260 * len(payload))),
            validation_context={"expected_ids": expected_ids},
        )
        if not step.accepted:
            self._log(
                f"[Agent状态] 语义复检Agent 批次未接受: ids={','.join(expected_ids[:12])}",
                "warning",
            )
            return {}
        result = step.payload
        raw_results = result.get("results") if isinstance(result, dict) else None
        if not isinstance(raw_results, list):
            return {}
        output: Dict[str, dict] = {}
        for raw in raw_results:
            if not isinstance(raw, dict):
                continue
            rid = str(raw.get("id", "")).strip()
            if not rid:
                continue
            normalized_issues = []
            for issue in raw.get("issues", []) or []:
                if not isinstance(issue, dict):
                    continue
                item = dict(issue)
                code = str(item.get("code", "OTHER") or "OTHER").strip().upper()
                item["code"] = code if code in ISSUE_CODES else "OTHER"
                # 工作台主区域只需要短建议；详细理由和证据单独保存。
                item["suggestion"] = str(item.get("suggestion", "") or "").strip()[:120]
                normalized_issues.append(item)
            output[rid] = {
                "status": raw["status"],
                "confidence": raw["confidence"],
                "issues": normalized_issues,
                "suggestions": raw.get("suggestions", {}),
                "selected_company_candidate_id": str(raw.get("selected_company_candidate_id") or "").strip(),
                "reason": raw.get("reason", ""),
            }
        self._log(
            f"[Agent状态] 语义复检Agent 批次已接受: 返回={len(output)}/{len(expected_ids)} 条"
        )
        return output
