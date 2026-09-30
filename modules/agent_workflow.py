"""固定的本地 Agent Workflow 编排层。

该层不依赖 Dify/Coze。它把每个模型步骤统一成：模型调用 → Pydantic
结构校验 → 最多一次修复 → 接受或转人工，并保留原有 dict 返回契约。
"""
from dataclasses import dataclass
from typing import Any, Dict, Optional, Type

from pydantic import BaseModel, ValidationError


@dataclass(frozen=True)
class WorkflowStepResult:
    step: str
    status: str
    payload: Optional[Dict[str, Any]] = None
    reason: str = ""

    @property
    def accepted(self) -> bool:
        return self.status == "accepted" and self.payload is not None


class AgentWorkflow:
    """调用底层客户端并把不合格输出统一导向人工复核。"""

    def __init__(self, client):
        self.client = client
        self.last_result: Optional[WorkflowStepResult] = None

    def run(
        self,
        *,
        step: str,
        system_prompt: str,
        user_content: str,
        output_model: Type[BaseModel],
        max_tokens: Optional[int] = None,
        validation_context: Optional[dict] = None,
    ) -> WorkflowStepResult:
        payload = self.client._call_api(
            system_prompt,
            user_content,
            purpose=step,
            max_tokens=max_tokens,
            output_model=output_model,
            validation_context=validation_context,
        )
        if payload is None:
            result = WorkflowStepResult(
                step=step,
                status="manual_review",
                reason="模型调用失败或输出连续两次未通过结构校验",
            )
        else:
            # 正常网络路径已经在 _call_api 中校验；这里再守一次边界，确保测试
            # Mock、未来替换的模型适配器也不可能绕过 Schema Gate。
            try:
                payload = output_model.model_validate(
                    payload,
                    context=validation_context or {},
                ).model_dump(by_alias=True)
            except (ValidationError, TypeError, ValueError):
                result = WorkflowStepResult(
                    step=step,
                    status="manual_review",
                    reason="模型适配器返回值未通过最终结构校验",
                )
                self.last_result = result
                return result
            result = WorkflowStepResult(
                step=step,
                status="accepted",
                payload=payload,
            )
        self.last_result = result
        return result
