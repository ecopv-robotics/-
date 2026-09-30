"""M4 项目标准化与拆分模块 — 将邮件中的项目简写/组合拆分为标准项目名称"""
import re
from typing import List, Dict

from utils.fuzzy_match import normalize_text, fuzzy_search


# 国家列表(从附件四提取)
COUNTRIES = [
    "德国", "荷兰", "爱尔兰", "意大利", "比利时", "波兰", "丹麦",
    "法国", "捷克", "葡萄牙", "瑞典", "西班牙", "卢森堡", "奥地利",
    "匈牙利", "芬兰", "罗马尼亚", "挪威", "爱沙尼亚", "瑞士",
    "希腊", "英国", "加拿大", "拉脱维亚",
]

# 业务类型
BUSINESS_TYPES = ["WEEE", "电池法", "包装法", "一次性塑料"]

# 英文国家名 → 中文。离线规则偶尔会产出 "Sweden WEEE" / "Ireland WEEE"
# 这类英文写法，直接按空格切分会把整串当成国家。
COUNTRY_ALIASES_EN = {
    "germany": "德国", "netherlands": "荷兰", "holland": "荷兰",
    "ireland": "爱尔兰", "italy": "意大利", "belgium": "比利时",
    "poland": "波兰", "denmark": "丹麦", "france": "法国",
    "czech": "捷克", "portugal": "葡萄牙", "sweden": "瑞典",
    "spain": "西班牙", "luxembourg": "卢森堡", "austria": "奥地利",
    "hungary": "匈牙利", "finland": "芬兰", "romania": "罗马尼亚",
    "norway": "挪威", "estonia": "爱沙尼亚", "switzerland": "瑞士",
    "greece": "希腊", "united kingdom": "英国", "canada": "加拿大",
    "latvia": "拉脱维亚",
}


# 「国家+EPR」统称，例如 "奥地利EPR" / "奥地利 EPR"。
# EPR 是生产者责任延伸的统称，代理商口语里常拿它代指该国 WEEE/电池法/包装法 三项，
# 项目名称表里并没有这种名字（表里只有"奥地利WEEE/电池法/包装法"）。
# 抽取出这种裸统称时，如果同一封邮件里已经有该国的具体项目，它就是冗余行。
_BARE_EPR_RE = re.compile(r"^\s*([\u4e00-\u9fa5]{2,6})\s*EPR\s*$", re.IGNORECASE)


def bare_epr_country(project: str) -> str:
    """识别「国家+EPR」统称，返回其中的国家；不是统称则返回空串。"""
    m = _BARE_EPR_RE.match(str(project or "").strip())
    if not m:
        return ""
    country = m.group(1)
    return country if country in COUNTRIES else ""


def prune_umbrella_epr(names: List[str]):
    """同一封邮件内剔除被同国具体项目覆盖的「国家+EPR」统称行。

    ``names`` 是同一封邮件拆分后的全部项目名。返回 ``(drop_idx, flag_idx)``
    两个下标集合：

    - ``drop_idx``：该国已有具体项目（如 奥地利WEEE），此统称行冗余，应丢弃；
    - ``flag_idx``：整封邮件只写了统称、没写任何具体项，保留但需人工确认。

    信息不会丢：只有"同国具体项已存在"时才删；否则保留并标注。
    """
    umbrella = {}  # 国家 -> 该项目名的下标列表（同国可能重复出现多次）
    for i, name in enumerate(names):
        country = bare_epr_country(name)
        if country:
            umbrella.setdefault(country, []).append(i)
    if not umbrella:
        return set(), set()

    covered = set()
    for name in names:
        if bare_epr_country(name):
            continue
        for country in umbrella:
            if country in str(name):
                covered.add(country)

    drop_idx, flag_idx = set(), set()
    for country, idx_list in umbrella.items():
        target = drop_idx if country in covered else flag_idx
        target.update(idx_list)
    return drop_idx, flag_idx


def country_of_project(project: str) -> str:
    """从标准化项目名里取出国家，取不到返回空串。

    项目名的标准写法是「国家+业务」连排或不连排，例如
    ``奥地利WEEE`` / ``德国 WEEE`` / ``Sweden WEEE``。
    旧实现按第一个空格切分，遇到无空格写法会把整串当成国家
    （国家列显示成 "奥地利WEEE"），这里改为按国家名做前缀匹配。
    """
    text = (project or "").strip()
    if not text:
        return ""
    for country in sorted(COUNTRIES, key=len, reverse=True):
        if text.startswith(country):
            return country
    for alias in sorted(COUNTRY_ALIASES_EN, key=len, reverse=True):
        if text.lower().startswith(alias):
            return COUNTRY_ALIASES_EN[alias]
    # 国家不在开头（如 "WEEE德国"）时退化为包含匹配
    for country in sorted(COUNTRIES, key=len, reverse=True):
        if country in text:
            return country
    return ""

# 组合项目映射 (附件四中以11/12/13/14结尾的组合项目)
COMBO_MAP = {
    "荷兰WEEE+包装法": ["荷兰WEEE", "荷兰包装法"],
    "荷兰电池法+包装法": ["荷兰电池法", "荷兰包装法"],
    "荷兰WEEE+电池法+包装法": ["荷兰WEEE", "荷兰电池法", "荷兰包装法"],
    "爱尔兰WEEE+电池法": ["爱尔兰WEEE", "爱尔兰电池法"],
    "意大利WEEE+电池法": ["意大利WEEE", "意大利电池法"],
    "波兰WEEE+波兰包装法": ["波兰WEEE", "波兰包装法"],
    "波兰电池法+波兰包装法": ["波兰电池法", "波兰包装法"],
    "波兰WEEE+波兰电池法+波兰包装法": ["波兰WEEE", "波兰电池法", "波兰包装法"],
    "捷克WEEE+电池法": ["捷克WEEE", "捷克电池法"],
    "匈牙利WEEE+匈牙利电池法": ["匈牙利WEEE", "匈牙利电池法"],
    "匈牙利WEEE+匈牙利包装法": ["匈牙利WEEE", "匈牙利包装法"],
    "匈牙利电池法+匈牙利包装法": ["匈牙利电池法", "匈牙利包装法"],
    "匈牙利WEEE+匈牙利电池法+匈牙利包装法": ["匈牙利WEEE", "匈牙利电池法", "匈牙利包装法"],
    "挪威WEEE/电池法/包装法": ["挪威WEEE", "挪威电池法", "挪威包装法"],
}


class ProjectNormalizer:
    def __init__(self, project_table: List[Dict], logger=None):
        """
        project_table: 附件四完整列表
        支持中文key和英文key:
        [{"项目编号": str, "项目名称": str, "国家": str, "业务类型": str}, ...]
        或 [{"project_name": str, "country": str, "type": str}, ...]
        """
        self.project_table = project_table

        def _get_name(p):
            return p.get("项目名称") or p.get("project_name") or ""

        def _get_country(p):
            return p.get("国家") or p.get("country") or ""

        def _get_type(p):
            return p.get("业务类型") or p.get("type") or ""

        self.standard_names = [_get_name(p) for p in project_table if _get_name(p)]
        self.project_lookup = {_get_name(p): p for p in project_table if _get_name(p)}
        self.logger = logger

        self.project_candidates = [
            (name, name) for name in self.standard_names
        ]

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    def normalize_and_split(self, raw_value: str, subject: str = "", body: str = "") -> List[Dict]:
        """
        标准化并拆分项目
        raw_value: 邮件中提取的项目原始文本
        subject/body: 用于推断国家上下文
        返回: [{"raw_value": str, "standard_name": str, "matched": bool}, ...]
        """
        if not raw_value:
            return []

        # 如果已经是标准名称，直接返回
        if raw_value in self.project_lookup:
            return [{"raw_value": raw_value, "standard_name": raw_value, "matched": True}]

        # 尝试组合项目映射
        combo_result = self._check_combo(raw_value)
        if combo_result:
            return combo_result

        # 尝试分隔符拆分
        split_parts = self._split_by_delimiters(raw_value)
        if len(split_parts) > 1:
            results = []
            # 用原始值作为额外上下文, 这样 "荷兰" 段能从 "荷兰、瑞典、波兰、比利时4国包装法" 知道业务类型是"包装法"
            full_context = raw_value + " " + subject + " " + body
            for part in split_parts:
                results.extend(self._normalize_single(part, full_context, body))
            if results:
                return results

        # 单个值标准化
        results = self._normalize_single(raw_value, subject, body)
        if results:
            return results

        # 模糊匹配
        match = fuzzy_search(raw_value, self.project_candidates, high_threshold=85, medium_threshold=75)
        if match["matched"]:
            return [{
                "raw_value": raw_value,
                "standard_name": match["results"][0]["text"],
                "matched": match["certain"],
            }]

        # 无法匹配
        self._log(f"项目标准化失败: {raw_value}", "warning")
        return [{"raw_value": raw_value, "standard_name": raw_value, "matched": False}]

    def _check_combo(self, raw_value: str) -> List[Dict]:
        """检查是否是已知的组合项目"""
        norm = normalize_text(raw_value)
        for combo_name, sub_items in COMBO_MAP.items():
            if normalize_text(combo_name) == norm:
                return [
                    {"raw_value": raw_value, "standard_name": item, "matched": True}
                    for item in sub_items
                ]
        return []

    def _split_by_delimiters(self, raw_value: str) -> List[str]:
        """按分隔符拆分"""
        # 按多种分隔符拆分
        parts = re.split(r"[+＋/／、，,]", raw_value)
        parts = [p.strip() for p in parts if p.strip()]

        if len(parts) > 1:
            return parts

        # 尝试国家组合模式: "波兰荷兰包装法" → ["波兰包装法", "荷兰包装法"]
        # 或 "荷兰、瑞典、波兰、比利时4国包装法" → 4条
        countries_found = []
        for c in COUNTRIES:
            if c in raw_value and c not in countries_found:
                countries_found.append(c)

        # 去掉国家后剩余部分作为业务类型
        if len(countries_found) > 1:
            remaining = raw_value
            for c in countries_found:
                remaining = remaining.replace(c, "")
            # 去除"N国"模式
            remaining = re.sub(r"\d+国", "", remaining)
            remaining = remaining.strip()

            if remaining and any(bt in remaining for bt in BUSINESS_TYPES):
                results = []
                for c in countries_found:
                    candidate = c + remaining
                    # 检查是否是标准名称
                    if candidate in self.project_lookup:
                        results.append(candidate)
                    else:
                        # 模糊匹配
                        match = fuzzy_search(
                            candidate, self.project_candidates,
                            high_threshold=85, medium_threshold=75
                        )
                        if match["matched"] and match["certain"]:
                            results.append(match["results"][0]["text"])
                        else:
                            results.append(candidate)

                if len(results) > 1:
                    return results

        return parts

    def _normalize_single(self, value: str, subject: str = "", body: str = "") -> List[Dict]:
        """标准化单个项目值"""
        value = value.strip()
        if not value:
            return []

        # 精确匹配
        if value in self.project_lookup:
            return [{"raw_value": value, "standard_name": value, "matched": True}]

        # 模糊匹配
        match = fuzzy_search(value, self.project_candidates, high_threshold=90, medium_threshold=80)
        if match["matched"] and match["certain"]:
            return [{
                "raw_value": value,
                "standard_name": match["results"][0]["text"],
                "matched": True,
            }]

        # 检查是否是纯业务类型 (如"电池"、"包装"、"weee")，需要补国家前缀
        context = subject + " " + body
        value_lower = value.lower()

        # 判断 value 是否是某个业务类型(或其简写)
        bt_aliases = {
            "WEEE": ["weee", "weee"],
            "电池法": ["电池", "电池法", "battery"],
            "包装法": ["包装", "包装法", "packaging"],
            "一次性塑料": ["一次性塑料", "塑料"],
        }

        matched_bt = None
        for bt, aliases in bt_aliases.items():
            for alias in aliases:
                if alias.lower() == value_lower:
                    matched_bt = bt
                    break
            if matched_bt:
                break

        if matched_bt:
            # 从上下文中找国家
            countries_in_context = [c for c in COUNTRIES if c in context]
            if not countries_in_context:
                # 从 value 自身找国家 (如 "德国电池" 中的 "德国")
                for c in COUNTRIES:
                    if c in value:
                        countries_in_context = [c]
                        break

            results = []
            for country in countries_in_context:
                candidate = country + matched_bt
                if candidate in self.project_lookup:
                    results.append({"raw_value": value, "standard_name": candidate, "matched": True})
            if results:
                return results

            # 模糊匹配补全
            for country in countries_in_context:
                candidate = country + matched_bt
                match2 = fuzzy_search(candidate, self.project_candidates, high_threshold=80, medium_threshold=70)
                if match2["matched"]:
                    results.append({
                        "raw_value": value,
                        "standard_name": match2["results"][0]["text"],
                        "matched": match2["certain"],
                    })
            if results:
                return results

        # 特殊处理: value 是纯国家名 (如 "荷兰" "瑞典"), 从上下文推断业务类型
        is_pure_country = False
        for c in COUNTRIES:
            if value == c:
                is_pure_country = True
                break

        if is_pure_country:
            # 从上下文中找业务类型
            context_bt = None
            for bt in BUSINESS_TYPES:
                if bt in context:
                    context_bt = bt
                    break
            if context_bt:
                candidate = value + context_bt
                if candidate in self.project_lookup:
                    return [{"raw_value": value, "standard_name": candidate, "matched": True}]
                match3 = fuzzy_search(candidate, self.project_candidates, high_threshold=80, medium_threshold=70)
                if match3["matched"]:
                    return [{
                        "raw_value": value,
                        "standard_name": match3["results"][0]["text"],
                        "matched": match3["certain"],
                    }]

        # 尝试从主题推断国家 (兼容旧逻辑)
        country = None
        for c in COUNTRIES:
            if c in context:
                country = c
                break

        if country:
            for bt in BUSINESS_TYPES:
                if bt.lower() in value_lower or bt in value:
                    candidate = country + bt
                    if candidate in self.project_lookup:
                        return [{"raw_value": value, "standard_name": candidate, "matched": True}]

        return []
