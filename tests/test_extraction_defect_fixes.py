# -*- coding: utf-8 -*-
"""提取逻辑确定性缺陷的回归断言（2026-10-16）。

覆盖五类"补齐缺失"型修复，每一类都同时给出**正向断言**（缺陷已被拦住）
与**防误伤断言**（正常数据仍照旧通过），避免修一处坏一处：

1. 国家清单共 5 份副本全部漏「斯洛伐克」
2. 公司名噪声闸门只认「填写」不认「必填」
3. 代理字段没有任何占位词闸门（TBA 一路进到业务明细）
4. 工单编号后的公司名边界缺「一次性塑料」，也缺「斯洛伐克」
5. llm_intent.batch_classify 为无调用方的死代码，已移除
"""
import re
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from modules.business_validator import (  # noqa: E402
    _PLACEHOLDER_AGENTS,
    inspect_row,
)
from modules.field_extractor import FieldExtractor  # noqa: E402
from modules.mail_filter import KEYWORDS_BUSINESS  # noqa: E402
from modules.project_normalizer import (  # noqa: E402
    COUNTRIES,
    COUNTRY_ALIASES_EN,
)
from utils.epr_form_parser import _DRAWING_FALLBACK_COUNTRIES  # noqa: E402

SLOVAK = "斯洛伐克"


# --------------------------------------------------------------------------- #
# 1. 斯洛伐克：国家清单的 5 份副本
# --------------------------------------------------------------------------- #

def test_project_normalizer_countries_include_slovakia():
    assert SLOVAK in COUNTRIES


def test_project_normalizer_english_alias_covers_slovakia():
    assert COUNTRY_ALIASES_EN.get("slovakia") == SLOVAK


def test_epr_form_drawing_fallback_includes_slovakia():
    # 这份兜底直接决定 EPR 申请表里的国家勾选能否被识别，缺一个国家就会让
    # 该国的申请表走到次优路径。
    assert SLOVAK in _DRAWING_FALLBACK_COUNTRIES


def test_mail_filter_business_keywords_include_slovakia():
    # 业务词表缺国家名时，只写“斯洛伐克xx法”的邮件会被判成“无业务词”。
    assert SLOVAK in KEYWORDS_BUSINESS


def test_english_country_code_map_includes_sk():
    """二字码映射 sk -> 斯洛伐克（附件里常写 SK）。"""
    source = (ROOT / "modules" / "field_extractor.py").read_text(encoding="utf-8")
    assert re.search(r'"sk"\s*:\s*"' + SLOVAK + r'"', source), "field_extractor 缺 sk 二字码映射"


# --------------------------------------------------------------------------- #
# 2. 公司名噪声：必填必须被拒
# --------------------------------------------------------------------------- #

def test_customer_noise_rejects_form_placeholder_bibian():
    assert FieldExtractor._customer_noise_reason("必填") != ""


@pytest.mark.parametrize(
    "name",
    [
        "深圳市美特欧科技有限公司",
        "READYWISE UK LTD",
        "香港景卓迪科技有限公司",
        "深圳市鸿盛达数码科技有限公司",
    ],
)
def test_customer_noise_still_accepts_real_companies(name):
    assert FieldExtractor._customer_noise_reason(name) == ""


# --------------------------------------------------------------------------- #
# 3. 代理占位词闸门
# --------------------------------------------------------------------------- #

def _row(agent: str, company: str = "深圳市美特欧科技有限公司"):
    return {
        "代理": agent,
        "客户": company,
        "项目": "德国WEEE",
        "需求": "注册",
    }


@pytest.mark.parametrize("agent", ["TBA", "tba", "TBD", "N/A", "待定", "未知", "未提供"])
def test_agent_placeholder_is_flagged(agent):
    codes = {item["code"] for item in inspect_row(_row(agent))}
    assert "AGENT_PLACEHOLDER" in codes


@pytest.mark.parametrize(
    "agent",
    ["华之星", "海税通", "恒诚越海", "基芯", "古道", "巨齿象", "跨规盈", "深圳华之星检测有限公司"],
)
def test_real_agent_name_is_not_flagged(agent):
    codes = {item["code"] for item in inspect_row(_row(agent))}
    assert "AGENT_PLACEHOLDER" not in codes


def test_agent_placeholder_does_not_rewrite_value():
    """占位词只标记人工复核，绝不改写原始提取值。"""
    row = _row("TBA")
    inspect_row(row)
    assert row["代理"] == "TBA"


def test_placeholder_agent_set_has_no_empty_key():
    """空字符串必须留给“代理为空”分支，不能混进占位词集合。"""
    assert "" not in _PLACEHOLDER_AGENTS


# --------------------------------------------------------------------------- #
# 4. 工单编号后的公司名边界
# --------------------------------------------------------------------------- #

def _loose_company(subject):
    from workbench_server import _subject_loose_company

    return _subject_loose_company(subject)


def test_loose_company_stops_at_single_use_plastic():
    # 修复前 boundary 只认到“注册”，候选变成
    # “READYWISE UK LTD 一次性塑料”，英文全匹配失败 → 返回空。
    assert _loose_company("AB123 READYWISE UK LTD 一次性塑料注册") == "READYWISE UK LTD"


def test_loose_company_stops_at_slovakia():
    assert _loose_company("AB123 READYWISE UK LTD 斯洛伐克包装法") == "READYWISE UK LTD"


@pytest.mark.parametrize(
    "subject,expected",
    [
        ("AB123 READYWISE UK LTD 德国WEEE", "READYWISE UK LTD"),
        ("AB123 READYWISE UK LTD 波兰包装法", "READYWISE UK LTD"),
        ("AB123 READYWISE UK LTD 爱尔兰包装法", "READYWISE UK LTD"),
    ],
)
def test_loose_company_existing_boundaries_unchanged(subject, expected):
    """既有边界词必须照旧生效，不能被本次补词影响。"""
    assert _loose_company(subject) == expected


def test_loose_company_plain_english_name_still_returned():
    assert _loose_company("AB123 READYWISE UK LTD") == "READYWISE UK LTD"


# --------------------------------------------------------------------------- #
# 5. 死代码与批大小
# --------------------------------------------------------------------------- #

def test_batch_classify_dead_code_removed():
    from modules.llm_intent import LLMIntentClient

    assert not hasattr(LLMIntentClient, "batch_classify"), (
        "batch_classify 无任何调用方，保留会误导后续维护"
    )
