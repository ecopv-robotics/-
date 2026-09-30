"""模糊匹配工具模块 — 处理错别字/简繁体差异"""
import re
from typing import List, Tuple

from rapidfuzz import fuzz

try:
    from opencc import OpenCC
    _cc = OpenCC("s2t")
    _cc_reverse = OpenCC("t2s")
    _has_opencc = True
except ImportError:
    _has_opencc = False


def to_simplified(text: str) -> str:
    if not text:
        return ""
    if _has_opencc:
        return _cc_reverse.convert(text)
    return text


def to_traditional(text: str) -> str:
    if not text:
        return ""
    if _has_opencc:
        return _cc.convert(text)
    return text


def normalize_text(text: str) -> str:
    if not text:
        return ""
    result = to_simplified(text)
    result = re.sub(r"\s+", "", result)
    result = result.replace("（", "(").replace("）", ")")
    result = result.replace("，", ",").replace("。", ".")
    return result.lower()


def fuzzy_search(
    query: str,
    candidates: List[Tuple[str, any]],
    high_threshold: int = 90,
    medium_threshold: int = 80,
) -> dict:
    """
    模糊搜索
    query: 待匹配文本
    candidates: [(文本, 关联数据), ...]
    返回: {
        'matched': bool,
        'certain': bool,
        'results': [{'text': str, 'data': any, 'score': int}, ...],
    }
    """
    if not query or not candidates:
        return {"matched": False, "certain": False, "results": []}

    query_norm = normalize_text(query)
    scored = []

    for text, data in candidates:
        text_norm = normalize_text(text)
        score = int(fuzz.ratio(query_norm, text_norm))
        if score >= medium_threshold:
            scored.append({"text": text, "data": data, "score": score})

    scored.sort(key=lambda x: x["score"], reverse=True)

    if not scored:
        return {"matched": False, "certain": False, "results": []}

    top_score = scored[0]["score"]
    if top_score >= high_threshold:
        return {"matched": True, "certain": True, "results": scored}

    return {"matched": True, "certain": False, "results": scored}


def fuzzy_match_pair(text_a: str, text_b: str, threshold: int = 85) -> bool:
    """判断两个文本是否模糊匹配"""
    if not text_a or not text_b:
        return False
    a_norm = normalize_text(text_a)
    b_norm = normalize_text(text_b)
    if not a_norm or not b_norm:
        return False
    # 子串包含匹配，例如公司简称包含于完整名称中。
    if a_norm in b_norm or b_norm in a_norm:
        return True
    score = fuzz.ratio(a_norm, b_norm)
    if score >= threshold:
        return True
    if len(a_norm) >= 2 and len(b_norm) >= 2:
        partial = fuzz.partial_ratio(a_norm, b_norm)
        if partial >= 90:
            return True
    return False
