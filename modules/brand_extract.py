"""品牌名统一抽取。

WEEE 与电池法过去各写一套品牌正则，结果各自缺一半：
* WEEE 只认「品牌/新增品牌/品牌名称」，遇到「增加品牌 Preciva」「原主体
  注册品牌：…」「注销1品牌：…」就提不出来；
* 电池法对动作词宽容，但值没有边界，会把后面的「第五类」「HST011」甚至
  整段正文一起吞掉，还会在「Co., Ltd.」的逗号处把品牌截断。

这里只保留一套规则：标签词表取并集，值与标签/类目/编号/换行处截断，
清洗后按顿号拆分多值。两个模块只保留各自的「类别判定」差异。
"""

from __future__ import annotations

import re
from typing import List

# 标签词表（并集）；顺序很重要，长词先匹配，避免「品牌」抢走「原主体注册品牌」。
BRAND_LABELS = (
    "原主体注册品牌", "原注册品牌", "注册品牌",
    "新增品牌", "新品牌", "增加品牌", "添加品牌",
    "品牌名称", "品牌名", "品牌",
    "brand name", "brand", "marke",
)

# 「注销2个品牌」「取消 1 款品牌」这类带动作/数量的写法，等同品牌标签。
_ACTION_BRAND_LABEL = re.compile(
    r"(?:注销|註銷|取消|撤销|撤銷|变更|變更)\s*\d+\s*(?:个|個|款|项|項|种|種)?\s*品牌",
    re.I,
)

_LABEL_PATTERN = "|".join(re.escape(label) for label in sorted(BRAND_LABELS, key=len, reverse=True))
_LABEL_RE = re.compile(rf"(?P<label>{_LABEL_PATTERN})\s*[:：=＝]?\s*", re.I)

# 值到这些词为止：下一个标签、电池类型、类目描述、第 N 类、客户编号、换行。
_STOP_RE = re.compile(
    rf"(?=\s*(?:{_LABEL_PATTERN}|电池(?:类型|类别|品类|分类)|类目|品类|类别|型号|"
    r"第\s*[1-6一二三四五六]\s*类|[A-Z]{2,6}\d{3,}))",
    re.I,
)

_URL_NOISE_RE = re.compile(
    r"https?://\S+|www\.\S+"
    r"|(?<![\w@.])[A-Za-z0-9][A-Za-z0-9-]*(?:\.[A-Za-z0-9-]+)*"
    r"\.(?:com|de|cn|net|eu|org|shop|top|site|online)(?:/[^\s,，;；|]*)?",
    re.I,
)

_PLACEHOLDER = re.compile(r"^(?:无|沒有|没有|未知|待确认|待確認|待定|暂无|n/?a|none|nil|不限)$", re.I)

# 申请表模板会把必填标记和填写说明夹在标签与值之间：
#   品牌 *（多个品牌，请用逗号隔开） KDTEND
# 旧逻辑直接在说明里的全角逗号处截断，得到 "*（多个" 这种碎片，品牌永远提不出来
# （反馈：申请表里明明写了 KDTEND，界面品牌是空的）。这里只剥掉**开头**的
# 括号说明，且必须是填写说明（含「请/多个/必填」等提示词），避免把
# 「（株）ABC」这类品牌自身的括号一起吃掉。
_LEADING_NOTE_RE = re.compile(
    r"^[\s*＊:：=＝]*[（(【\[][^）)】\]]{1,120}[）)】\]][\s*＊:：=＝]*"
)
_NOTE_HINT_RE = re.compile(r"请|請|多个|多個|必填|选填|選填|注[:：]|示例|格式|如有|separat", re.I)


def _strip_leading_note(value: str) -> str:
    """剥掉标签后紧跟的「*（多个品牌，请用逗号隔开）」类填写说明。"""
    text = value
    for _ in range(3):
        match = _LEADING_NOTE_RE.match(text)
        if not match or not _NOTE_HINT_RE.search(match.group(0)):
            break
        text = text[match.end():]
    return text


def _clean_brand(value: str) -> str:
    text = _URL_NOISE_RE.sub(" ", str(value or ""))
    text = re.sub(r"&(?:nbsp|amp);", " ", text, flags=re.I)
    # 不剥半角句点：Co., Ltd. 的点是品牌名的一部分。
    text = re.sub(r"\s{2,}", " ", text).strip(" \t:：=＝,，;；|+＋-—_。*＊")
    if not text or len(text) > 90:
        return ""
    # 整串就是个括号说明（"（日本法人）"）→ 不是品牌。
    if re.fullmatch(r"[（(【\[][^）)】\]]*[）)】\]]", text):
        return ""
    if _PLACEHOLDER.match(text):
        return ""
    return text


def _cut_value(tail: str) -> str:
    """从标签之后截出品牌值：到换行、分号或下一个字段为止。"""
    tail = tail.split("\n")[0].split("\r")[0]
    # 先剥掉「*（多个品牌，请用逗号隔开）」这类填写说明，再做分隔符切分：
    # 说明文字自带全角逗号，先切会得到 "*（多个" 这样的碎片。
    tail = _strip_leading_note(tail)
    # 全角逗号与加号永远是分隔符；半角逗号要留给「Co., Ltd.」。
    tail = re.split(r"[;；|，＋+]", tail, maxsplit=1)[0]
    stop = _STOP_RE.search(tail)
    if stop:
        tail = tail[: stop.start()]
    return tail


def extract_brand_values(text: str, multi_split: bool = True) -> List[str]:
    """从文本里抽取品牌名，覆盖冒号/无冒号/动作词/多值等写法。"""
    raw = str(text or "")
    if not raw.strip():
        return []

    spans = []
    for match in _LABEL_RE.finditer(raw):
        spans.append((match.start(), match.end(), match.group("label")))
    for match in _ACTION_BRAND_LABEL.finditer(raw):
        spans.append((match.start(), match.end(), match.group(0)))
    spans.sort()

    values: List[str] = []
    for _, end, label in spans:
        value = _cut_value(raw[end:])
        if not value:
            continue
        parts = [value]
        if multi_split and re.search(r"、", value):
            parts = [part for part in value.split("、") if part.strip()]
        elif multi_split and re.search(r"\d+\s*(?:个|個|款|项|項|种|種)", label):
            parts = [part for part in re.split(r"[、,，]", value) if part.strip()]
        for part in parts:
            cleaned = _clean_brand(part)
            if cleaned and cleaned not in values:
                values.append(cleaned)
    return values
