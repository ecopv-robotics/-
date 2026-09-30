"""德国 WEEE 品牌/品类与注册工单品类明细核对。

这个模块只做两件事：

* 从邮件正文、主题及附件结构化行中读取明确出现的“品牌/品类”；
* 将这些值与注册工单查询结果中明确标注的品类字段逐项比对。

没有明确证据时返回 ``pending``，不根据产品常识或数字猜测品类。
"""

from __future__ import annotations

import json
import re
from typing import Any, Dict, Iterable, List, Sequence


GERMANY_ALIASES = ("德国", "德國", "germany", "german", "deutschland", "de")
WEEE_RE = re.compile(r"(?<![A-Za-z])WEEE(?![A-Za-z])", re.I)
WEEE_RULE_SOURCE = "产品分类表 中文版.docx（德国 ElektroG/WEEE 六类）"

_BRAND_KEYS = (
    "品牌", "品牌名称", "品牌名", "新增品牌", "新品牌", "brand", "brand name", "marke",
)
_CATEGORY_KEYS = (
    "品类", "品類", "类目", "類目", "类别", "類別", "产品类目", "產品類目", "产品类别", "產品類別", "商品类别",
    "商品類別", "产品分类", "產品分類", "设备类别", "設備類別", "注册类别",
    "注册品类", "申报品类", "申報品類", "新增类别", "新增品类", "新类别", "新品类",
    "新增类目", "新增品类", "新类目", "类目名称", "類目名稱",
    "category", "product category",
    "product type", "product name", "product type/name", "warengruppe", "produktkategorie",
)
_WORKORDER_CATEGORY_KEYS = _CATEGORY_KEYS + (
    "品类明细", "品類明細", "类别明细", "類別明細", "分类明细", "分類明細",
    "品类名称", "品類名稱", "类别名称", "類別名稱", "品类名", "类别名",
    "注册类型明细", "产品信息", "商品信息", "产品名称", "商品名称",
    "服务品类", "服务类别", "category name", "product name",
)

# 《产品分类表 中文版.docx》中的德国 ElektroG/WEEE 六类边界。
# 这里保存“能直接从邮件/附件文字判断”的短语，不把产品名称强行猜成类别：
# 命中多个类别或只命中泛化词时，后续结果会标记为 pending 交人工确认。
WEEE_CATEGORY_DEFINITIONS = {
    "1": {
        "name": "热交换设备",
        "aliases": (
            "热交换设备", "温度交换设备", "制冷设备", "冷却设备", "冰箱", "冰柜",
            "冷冻设备", "饮料冷却器", "冷藏柜台", "冷水机", "冷热饮水机", "空调",
            "热泵", "热泵干燥机", "含油散热器", "工业压缩机", "冷却循环器",
            "heat exchange", "temperature exchange", "refrigerator", "freezer",
            "cooling equipment", "air conditioner", "air conditioning", "heat pump",
            "compressor", "chiller", "dehumidifier", "冷却/冷冻",
            "wärmeüberträger", "temperaturüberträger", "kühlgerät", "kühlschrank",
            "gefrierschrank", "klimagerät", "wärmepumpe", "entfeuchter",
        ),
        "excludes": ("无冷却功能", "不带冷却功能", "without cooling"),
    },
    "2": {
        "name": "屏幕和显示设备",
        "aliases": (
            "屏幕", "显示器", "监视器", "电视机", "电视", "电子书阅读器", "阅读器",
            "笔记本电脑", "笔记本", "平板电脑", "lcd", "led屏", "oled", "触摸屏",
            "投影屏幕", "视频显示器", "图形显示器", "相框", "screen", "display",
            "monitor", "television", "tv", "laptop", "notebook", "tablet", "e-reader",
            "touchscreen", "projector screen",
            "bildschirmgeräte", "bildschirm", "fernseher",
        ),
        "excludes": (),
    },
    "3": {
        "name": "灯具和光源",
        "aliases": (
            "气体放电灯", "放电灯", "荧光灯", "紧凑型荧光灯", "节能灯", "led灯",
            "led灯泡", "led灯丝灯", "金属卤化物灯", "霓虹灯", "汞蒸汽灯", "紫外线灯",
            "灯具", "照明器具", "路灯", "照明灯", "灯泡", "lamp", "luminaire",
            "lighting", "fluorescent", "discharge lamp", "led lamp", "bulb", "neon",
            "street light", "uv lamp",
            "lampen", "leuchte", "entladungslampe", "led-lampe",
        ),
        # docx 明确说明传统白炽灯/卤素灯不在 ElektroG 第 3 类范围。
        "excludes": ("白炽灯", "卤素灯", "incandescent", "halogen"),
    },
    "4": {
        "name": "大型设备",
        "aliases": (
            "大型设备", "大件设备", "大型家电", "大型电气设备", "大型光伏电池板",
            "大型非光伏设备", "大型非光伏", "光伏电池板", "太阳能电池板", "工业设备", "商业厨房设备", "电动医院病床",
            "大型医疗设备", "自动售货机", "取款机", "充电站", "电动汽车充电柱",
            "电机", "发电机", "打印机", "复印机", "大型打印机", "大型it设备", "large equipment",
            "large appliance", "industrial equipment", "pv panel", "solar panel",
            "vending machine", "atm", "charging station", "large printer",
            "großgeräte", "großgerät", "großes gerät",
        ),
        "excludes": ("小型", "small", "不超过50厘米", "<=50cm", "50 cm or less"),
    },
    "5": {
        "name": "小型设备",
        "aliases": (
            "小型设备", "小家电", "小型家电", "小型电气设备", "小型光伏电池板",
            "小型非光伏设备", "小型非光伏", "小型工具", "吸尘器", "咖啡机", "微波炉", "风扇", "加湿器", "电动工具",
            "玩具", "耳机", "扬声器", "摄像机", "照相机", "无人机", "智能手表",
            "体温计", "电动牙刷", "充电器", "插座", "延长线", "电缆", "电源适配器",
            "small equipment", "small appliance", "vacuum cleaner", "microwave",
            "fan", "humidifier", "power tool", "toy", "headphone", "speaker",
            "camera", "drone", "smartwatch", "thermometer", "charger", "socket",
            "extension cable", "adapter", "cable",
            "kleingeräte", "kleingerät", "kleine geräte",
        ),
        "excludes": ("屏幕面积大于100", "大于100平方厘米", ">100 cm2", "large equipment"),
    },
    "6": {
        "name": "小型信息和电信设备",
        "aliases": (
            "小型信息和电信设备", "小型it设备", "信息技术设备", "电信设备", "电脑",
            "个人电脑", "计算机", "服务器", "打印机", "扫描仪", "路由器", "交换机",
            "网络设备", "收银机", "读卡器", "电话", "手机", "移动电话", "智能手机",
            "传真机", "硬盘", "键盘", "鼠标", "内存卡", "usb", "网线", "电话线",
            "显示端口电缆", "hdmi电缆", "小型itk", "small it", "it equipment",
            "telecommunication", "computer", "pc", "server", "printer", "scanner",
            "router", "switch", "cash register", "card reader", "telephone", "mobile phone",
            "smartphone", "fax", "hard drive", "keyboard", "mouse", "memory card",
            "usb cable", "network cable", "ethernet", "itk",
            "kleine it", "kleine informations- und telekommunikationsgeräte",
            "telekommunikationsgerät",
        ),
        # 第 6 类是“外部尺寸不超过 50 厘米”的 IT/通信设备；大 IT 归第 4 类，
        # 大于 100 cm² 的主要显示设备归第 2 类。
        "excludes": ("大型it", "大于100平方厘米", ">100 cm2", "屏幕面积大于100", "large it"),
    },
}

_CHINESE_CATEGORY_NUMBERS = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5", "六": "6"}
_CATEGORY_LABEL_RE = re.compile(
    r"(?:第\s*([1-6一二三四五六])\s*类|(?:category|class|klasse)\s*([1-6])|类别\s*([1-6一二三四五六])|分类\s*([1-6一二三四五六]))",
    re.I,
)

# 申请表中常见的说明/示例不能作为真实品牌或品类。尤其是“例：设备电池，
# Li，Apple，10g，3500个”会被旧版按分隔符拆成多条假品类。
_EXAMPLE_MARKER_RE = re.compile(r"(?:例\s*[:：]|例如|示例|example|e\.g\.?|for\s+example)", re.I)
_MEASUREMENT_ONLY_RE = re.compile(
    r"^\s*\d+(?:\.\d+)?\s*(?:g|kg|克|千克|个|件|台|套|cm|mm|m|厘米|毫米|米|平方厘米|cm2)\s*$",
    re.I,
)
_DIMENSION_ONLY_RE = re.compile(
    r"^\s*(?:(?:单边|外部|最大|最小)?\s*尺寸\s*)?(?:小于|不超过|不大于|≤|<=|低于|少于)?\s*"
    r"\d+(?:\.\d+)?\s*(?:cm|厘米|mm|毫米|m|米)(?:以下|以内)?\s*$",
    re.I,
)
_PLACEHOLDER_RE = re.compile(
    r"^(?:对应的?品牌(?:logo)?|品牌\s*\d*|品牌名|品牌名称|产品类型|产品类别|产品分类|商品类别|"
    r"电池类型(?:\s*type)?|对应品牌|marke|type|类别|品类|包装材质类别|必填|选填)"
    r"(?:\s*[（(].*)?$",
    re.I,
)
_TEMPLATE_LABEL_RE = re.compile(
    r"(?:信息在下列填写|需要的服务|注册信息|申请表|公司英文名称|公司中文名称|法人名字|"
    r"营业执照号码|vat号|德国\s*vat|邮政编码|签字时间)",
    re.I,
)


def _text(value: Any) -> str:
    return str(value or "").replace("\u3000", " ").strip()


def _norm(value: Any) -> str:
    text = _text(value).casefold()
    return re.sub(r"[\s\-_/.,，。:：;；|+＋()（）\[\]【】{}]+", "", text)


def _norm_product(value: Any) -> str:
    """产品分类匹配用的规范化文本。

    保留数字（如 100 cm²）用于排除规则，去掉 HTML 空格和常见分隔符，
    但不做词干化或同义词臆测。
    """
    text = _text(value)
    text = re.sub(r"&(?:nbsp|amp);", " ", text, flags=re.I)
    text = text.replace("平方厘米", "cm2").replace("厘米", "cm")
    text = text.replace("㎝", "cm").replace("²", "2")
    return re.sub(r"[\s\-_/.,，。:：;；|+＋()（）\[\]【】{}'\"“”‘’]+", "", text.casefold())


def _category_id_from_label(value: Any) -> str:
    text = _text(value)
    match = _CATEGORY_LABEL_RE.search(text)
    if not match:
        return ""
    raw = next((group for group in match.groups() if group), "")
    return _CHINESE_CATEGORY_NUMBERS.get(raw, raw)


def _category_alias_hits(value: Any) -> List[Dict[str, Any]]:
    """根据产品/品类文字返回六类候选及命中证据。"""
    raw = _text(value)
    normalized = _norm_product(raw)
    if not normalized:
        return []
    explicit = _category_id_from_label(raw)
    if explicit:
        return [{"category_class": explicit, "category_class_name": WEEE_CATEGORY_DEFINITIONS[explicit]["name"],
                 "score": 1000, "evidence": raw, "reason": "命中分类表中的明确类别编号"}]

    hits: Dict[str, Dict[str, Any]] = {}
    for class_id, definition in WEEE_CATEGORY_DEFINITIONS.items():
        if any(_norm_product(ex) and _norm_product(ex) in normalized for ex in definition.get("excludes", ())):
            continue
        for alias in definition["aliases"]:
            alias_norm = _norm_product(alias)
            if not alias_norm or alias_norm not in normalized:
                continue
            # 单字/极泛的词（如“灯”“电话”）证据弱；完整短语证据强。
            score = len(alias_norm) * 3
            if len(alias_norm) <= 1:
                score = 2
            elif len(alias_norm) <= 2:
                # “手机/电脑/空调/热泵”等二字产品词是分类表中的明确条目；
                # 只有单字泛词才保持弱命中。
                score = 12 if alias_norm == normalized else 6
            elif len(alias_norm) <= 4:
                score = max(score, 12)
            hit = hits.setdefault(class_id, {
                "category_class": class_id,
                "category_class_name": definition["name"],
                "score": 0,
                "evidence": [],
                "reason": "产品分类表关键词命中",
            })
            if score > hit["score"]:
                hit["score"] = score
            if alias not in hit["evidence"]:
                hit["evidence"].append(alias)
    return sorted(hits.values(), key=lambda item: (-item["score"], item["category_class"]))


def classify_weee_product(value: Any) -> Dict[str, Any]:
    """按产品分类表给单个品牌/品类文字归类。

    返回 ``matched``、``ambiguous`` 或 ``unmatched``。只有唯一高置信命中才
    进入自动比对；跨类别命中或只剩弱词命中都保持待人工确认。
    """
    raw = _clean_candidate(value)
    if not raw:
        return {"status": "unmatched", "category_class": "", "category_class_name": "", "candidates": [], "evidence": []}
    hits = _category_alias_hits(raw)
    if not hits:
        return {"status": "unmatched", "category_class": "", "category_class_name": "", "candidates": [], "evidence": []}
    top = hits[0]
    # 明确类别编号或足够长的完整词组可自动确认；若第二候选相近则不自动猜。
    if top["score"] >= 1000:
        status = "matched"
    elif top["score"] >= 12 and (len(hits) == 1 or top["score"] > hits[1]["score"] * 1.35):
        status = "matched"
    elif top["score"] >= 8:
        status = "ambiguous"
    else:
        status = "unmatched"
    return {
        "status": status,
        "category_class": top["category_class"] if status == "matched" else "",
        "category_class_name": top["category_class_name"] if status == "matched" else "",
        "candidates": [{
            "category_class": hit["category_class"],
            "category_class_name": hit["category_class_name"],
            "score": hit["score"],
        } for hit in hits],
        "evidence": top.get("evidence", []),
    }


def decompose_weee_category(value: Any) -> Dict[str, Any]:
    """拆解德国 WEEE 邮件中的复合品类描述，保留原文并提取结构化组成。

    例如 ``电气与电子设备废料-大型设备(光伏面板除外)`` 会拆成：
    ``category_family=电气与电子设备废料``、``equipment_size=大型设备``、
    ``exclusions=[光伏面板]``。拆解只用于证据展示和审计，最终第 1–6 类仍
    由 :func:`classify_weee_product` 根据原始完整文本决定。
    """
    raw = _clean_candidate(value)
    if not raw:
        return {}
    components: Dict[str, Any] = {}
    family_match = re.search(
        r"(?:电气与电子设备废料|电气和电子设备废料|电子电气设备废料|电子与电气设备废料)",
        raw,
        re.I,
    )
    if family_match:
        components["category_family"] = family_match.group(0)
    size_match = re.search(r"(大型设备|小型设备|大型非光伏设备|小型非光伏设备)", raw, re.I)
    if size_match:
        components["equipment_size"] = size_match.group(1)
    exclusions: List[str] = []
    for match in re.finditer(r"([^()（）,，;；]+?)\s*除外", raw):
        excluded = _clean_candidate(match.group(1))
        if excluded and excluded not in exclusions:
            exclusions.append(excluded)
    if exclusions:
        components["exclusions"] = exclusions
    return components


def _flatten_values(value: Any) -> List[str]:
    """展开工单品类字段里常见的 list/dict/JSON 结构。"""
    if value is None:
        return []
    if isinstance(value, dict):
        values: List[str] = []
        for key, child in value.items():
            if key in {"status", "reason", "score"}:
                continue
            values.extend(_flatten_values(child))
        return values
    if isinstance(value, (list, tuple, set)):
        values: List[str] = []
        for child in value:
            values.extend(_flatten_values(child))
        return values
    return _split_values(value)


def _nested_alias_values(value: Any, aliases: Sequence[str]) -> List[str]:
    """从品类明细的嵌套记录中只取品类字段，避免把编号/品牌当品类。"""
    if isinstance(value, dict):
        values: List[str] = []
        for key, child in value.items():
            key_norm = re.sub(r"\s+", "", _text(key).casefold())
            if any(re.sub(r"\s+", "", alias.casefold()) in key_norm for alias in aliases):
                values.extend(_flatten_values(child))
            elif key_norm in {"data", "rows", "records", "items", "品类明细", "品類明細"}:
                values.extend(_nested_alias_values(child, aliases))
        return values
    if isinstance(value, (list, tuple, set)):
        values: List[str] = []
        for child in value:
            values.extend(_nested_alias_values(child, aliases))
        return values
    return _split_values(value)


def _split_values(value: Any) -> List[str]:
    text = _text(value)
    if not text:
        return []
    values = re.split(r"[\n\r,，;；、|/／]+", text)
    return [item.strip(" \t:：-—") for item in values if item.strip(" \t:：-—")]


def is_germany_weee(project: Any = "", subject: Any = "", body: Any = "", attachments: Any = None) -> bool:
    """只在明确出现德国 WEEE 服务的邮件/项目上启用专项规则。

    申请表模板经常固定包含“WEEE 产品信息”字样；附件正文命中这段
    模板说明不能证明本封邮件申请的是 WEEE。已标准化的非 WEEE 项目
    （例如“德国包装法”）优先作为排他信号，避免在包装法卡片上显示
    “德国 WEEE 品类待核对”。
    """
    project_text = _text(project)
    subject_text = _text(subject)
    body_text = _text(body)
    # 项目字段一旦明确且不含 WEEE，就不因申请表模板中的通用文字开启专项。
    if project_text and not WEEE_RE.search(project_text):
        return False

    # 附件文件名可作为服务声明证据；附件 text_content 不纳入启用判断，
    # 因为其中往往是整张申请表模板，固定包含 WEEE 标签。
    text_parts = [project_text, subject_text, body_text]
    for attachment in attachments or []:
        if isinstance(attachment, dict):
            text_parts.append(_text(attachment.get("filename")))
    text = " ".join(text_parts)
    has_germany = any(
        bool(re.search(r"(?<![A-Za-z])DE(?![A-Za-z])", text, re.I))
        if alias.casefold() == "de"
        else alias.casefold() in text.casefold()
        for alias in GERMANY_ALIASES
    )
    return bool(has_germany and WEEE_RE.search(text))


def _label_value_pairs(text: str, labels: Sequence[str]) -> List[tuple[str, str]]:
    """读取“品牌：xxx / 类别：yyy”这类明确标签，不抓整段自然语言。"""
    if not text:
        return []
    label_pattern = "|".join(re.escape(label) for label in sorted(labels, key=len, reverse=True))
    stop_pattern = "|".join(re.escape(label) for label in sorted(set(labels) | set(_BRAND_KEYS) | set(_CATEGORY_KEYS), key=len, reverse=True))
    # 值到下一个标签、换行或分隔符为止；允许英文品牌包含空格和点号。
    pattern = re.compile(
        rf"(?:^|[\n\r;；|\s])\s*(?P<label>{label_pattern})\s*[:：=＝]\s*"
        rf"(?P<value>[^\n\r;；|]+?)(?=\s*(?:{stop_pattern})\s*[:：=＝]|\s*(?:[;；|]|$))", re.I,
    )
    return [(_text(match.group("label")), _text(match.group("value"))) for match in pattern.finditer(text)]


_INLINE_BRAND_MARKERS = ("新增品牌", "新品牌", "品牌名称", "品牌名", "品牌")
_INLINE_CATEGORY_MARKERS = (
    "新增类别", "新增品类", "新类别", "新品类", "产品类别", "产品分类", "注册类别",
    "注册品类", "申报品类", "新增类目", "新类目", "产品类目", "类目", "品类", "类别",
)


def _inline_marker_pairs(text: str, labels: Sequence[str]) -> List[tuple[str, str]]:
    """读取不带冒号的邮件写法，例如“新增品牌AUZONIMICS 第六类”。"""
    if not text:
        return []
    label_pattern = "|".join(re.escape(label) for label in sorted(labels, key=len, reverse=True))
    all_markers = "|".join(re.escape(label) for label in sorted(
        set(_INLINE_BRAND_MARKERS) | set(_INLINE_CATEGORY_MARKERS), key=len, reverse=True,
    ))
    pattern = re.compile(
        rf"(?<![\u4e00-\u9fffA-Za-z0-9])(?P<label>{label_pattern})[ \t]*(?:[:：=＝][ \t]*)?(?P<value>[^\n\r;；|]+)", re.I,
    )
    pairs: List[tuple[str, str]] = []
    for match in pattern.finditer(text):
        # 主题里常把“新增类别”作为普通标题词而不是字段，例如“新增类别 示例代理己…”。
        # 无冒号的类别字段交给“第几类”或同一行自然语言分类逻辑处理，避免把整段主题当品类。
        has_separator = bool(re.match(r"[ \t]*[:：=＝]", text[match.end("label"):]))
        if set(labels).issubset(set(_INLINE_CATEGORY_MARKERS)) and not has_separator:
            continue
        value = match.group("value")
        # 一个片段里可能连续写成“新增类别：小型非光伏（品牌：PURELLEL）”或
        # “新增品牌AUZONIMICS 第六类”，先在下一个字段/类别编号处截断。
        value = re.split(
            rf"(?=\s*(?:{all_markers})\s*(?:[:：=＝]|(?=[\u4e00-\u9fffA-Za-z]))|"
            rf"\s*第\s*[1-6一二三四五六]\s*类)",
            value, maxsplit=1, flags=re.I,
        )[0]
        # 无冒号写法“品牌 X 冷却设备”中，品牌值和自然语言品类相邻；
        # 品类命中分类表后把它从品牌值中拆出，避免把两个字段粘成一个。
        if not has_separator and set(labels).issubset(set(_INLINE_BRAND_MARKERS)):
            context_classification = classify_weee_product(value)
            if context_classification.get("status") == "matched":
                for evidence in context_classification.get("evidence", []):
                    index = value.casefold().find(_text(evidence).casefold())
                    if index > 0:
                        value = value[:index]
                        break
        value = _clean_candidate(value)
        if value:
            pairs.append((_text(match.group("label")), value))
    return pairs


def _category_class_pairs(text: str) -> List[tuple[str, str]]:
    """将正文/标题中的“第 1 类/第六类”保留为可追溯的原始品类。"""
    if not text:
        return []
    pairs: List[tuple[str, str]] = []
    seen = set()
    for match in re.finditer(r"第\s*[1-6一二三四五六]\s*类", text, re.I):
        value = _clean_candidate(match.group(0))
        key = _norm(value)
        if value and key not in seen:
            seen.add(key)
            pairs.append(("分类编号", value))
    return pairs


def _context_category_values(text: str) -> List[str]:
    """在带有品牌标记的同一行补抓自然语言品类，如“品牌 X 冷却设备”。"""
    if not text:
        return []
    brand_pattern = "|".join(re.escape(label) for label in _INLINE_BRAND_MARKERS)
    values: List[str] = []
    for line in re.split(r"[\r\n]+", text):
        if not re.search(rf"(?:{brand_pattern})", line, re.I):
            continue
        hits = _category_alias_hits(line)
        if not hits:
            continue
        top = hits[0]
        if top.get("score", 0) < 12 or (len(hits) > 1 and top["score"] <= hits[1]["score"] * 1.35):
            continue
        values.extend(str(value) for value in top.get("evidence", []) if _clean_candidate(value))
    return _dedupe(values)


def _clean_candidate(value: Any) -> str:
    # 先去除字段分隔符；成对括号属于原始品类描述（例如“光伏面板除外”），
    # 不能像普通标点一样把结尾括号剥掉，否则证据文本会被截断。
    text = _text(value).strip(" \t:：-—_，,;；")
    bracket_pairs = (("(", ")"), ("（", "）"), ("[", "]"), ("【", "】"))
    for opening, closing in bracket_pairs:
        if text.endswith(closing) and text.count(opening) < text.count(closing):
            text = text.rstrip(closing)
    text = text.strip("()（）[]【】") if text.startswith(("(", "（", "[", "【")) and not text.endswith((")", "）", "]", "】")) else text
    if not text or len(text) > 180:
        return ""
    # 不能把业务标签、数量或整句说明当品牌/品类。
    if re.fullmatch(r"\d+(?:\.\d+)?", text) or _MEASUREMENT_ONLY_RE.fullmatch(text):
        return ""
    if re.search(r"^(?:无|未知|待确认|待定|暂无|n/?a|none|nil)$", text, re.I):
        return ""
    if text in {"·", "•", "-", "—", "_"} or _PLACEHOLDER_RE.fullmatch(text):
        return ""
    # 这些是模板字段名，不是客户填写的产品数据；保留含有真实分类词的复合值，
    # 例如“小型设备（必填）”仍可由分类表识别。
    if _TEMPLATE_LABEL_RE.search(text) and not _category_alias_hits(text) and not _category_id_from_label(text):
        return ""
    return text


def _is_dimension_only(value: Any) -> bool:
    """判断是否只有尺寸限制，没有独立的产品类别。"""
    return bool(_DIMENSION_ONLY_RE.fullmatch(_text(value)))


def _is_example_record(value: Any) -> bool:
    return bool(_EXAMPLE_MARKER_RE.search(_text(value)))


def _combine_dimension_categories(values: Iterable[str]) -> List[str]:
    """把同一单元格拆出的“类别 + 尺寸”恢复为一个类别，丢弃孤立尺寸。"""
    result: List[str] = []
    for value in _dedupe(values):
        if _is_dimension_only(value):
            if result:
                result[-1] = f"{result[-1]}，{value}"
            continue
        result.append(value)
    return result


def _record_values(record: Dict[str, Any], keys: Sequence[str]) -> List[str]:
    values: List[str] = []
    lowered = {re.sub(r"\s+", "", str(key).casefold()): value for key, value in record.items()}
    for alias in keys:
        alias_key = re.sub(r"\s+", "", alias.casefold())
        for key, value in lowered.items():
            if key == alias_key or alias_key in key:
                values.extend(_split_values(value))
    return [_clean_candidate(value) for value in values if _clean_candidate(value)]


def _record_category_values(record: Dict[str, Any]) -> List[str]:
    """读取结构化行的原始品类，不把自然语言逗号误当成多条品类。

    Excel 的德文类别常写成 ``Kleingeräte, die ...``；通用字段拆分器会把
    逗号拆开，随后把“品牌 × 类别片段”展开成多条卡片。结构化行已经有
    明确的 category 字段，此处只把真正的行/列分隔符 ``|`` 拆开，保留
    原始品类描述供证据和人工复核使用。
    """
    record_type = _text(record.get("record_type"))
    raw_category = _text(record.get("category"))
    if record_type in {"entity", "weee_catalog"} and raw_category:
        values = [value.strip() for value in re.split(r"[\n\r|]+", raw_category) if value.strip()]
        return [_clean_candidate(value) for value in values if _clean_candidate(value)]
    return _record_values(record, _CATEGORY_KEYS)


def _dedupe(values: Iterable[str]) -> List[str]:
    result: List[str] = []
    seen = set()
    for value in values:
        cleaned = _clean_candidate(value)
        key = _norm(cleaned)
        if cleaned and key and key not in seen:
            seen.add(key)
            result.append(cleaned)
    return result


def _sheet_preview_records(attachment: Dict[str, Any]) -> List[Dict[str, Any]]:
    """从阶段一保存的 xlsx 预览行恢复 WEEE 品牌/品类记录。

    阶段一的新结构化解析通常已经提供 ``structured_records``，但历史缓存、
    压缩包内表格和旧版工作簿有时只保留 ``sheets.preview_rows``。这里仅在
    同一张表中找到品牌列和品类/产品描述列时读取，避免把申请表的普通字段
    或表格说明文字误当成 WEEE 项目。
    """
    records: List[Dict[str, Any]] = []
    for sheet in attachment.get("sheets") or []:
        if not isinstance(sheet, dict):
            continue
        rows = sheet.get("preview_rows") or sheet.get("rows") or []
        normalized_rows: List[tuple[int, List[str]]] = []
        for row in rows:
            if isinstance(row, dict):
                row_number = row.get("row_number") or row.get("index") or "?"
                cells = row.get("cells") or row.get("values") or []
            elif isinstance(row, (list, tuple)):
                row_number, cells = "?", row
            else:
                continue
            if not isinstance(cells, (list, tuple)):
                continue
            normalized_rows.append((int(row_number) if str(row_number).isdigit() else row_number, [_text(cell) for cell in cells]))

        header_index = -1
        brand_columns: List[int] = []
        category_columns: List[int] = []
        weee_catalog_header = False
        for index, (_row_number, cells) in enumerate(normalized_rows[:60]):
            brands = [column for column, cell in enumerate(cells) if _is_header_alias(cell, _BRAND_KEYS)]
            categories = [column for column, cell in enumerate(cells) if _is_header_alias(cell, _CATEGORY_KEYS)]
            if brands and categories:
                header_index = index
                brand_columns = brands
                category_columns = categories
                weee_catalog_header = any(
                    re.search(r"(?<![A-Za-z])WEEE(?![A-Za-z])", _text(cell), re.I)
                    for cell in cells
                )
                break
        if header_index < 0:
            continue

        records_started = False
        previous_row_number: int | None = None
        for row_number, cells in normalized_rows[header_index + 1:]:
            if (
                weee_catalog_header
                and records_started
                and isinstance(row_number, int)
                and previous_row_number is not None
                and row_number > previous_row_number + 1
            ):
                # preview_rows 只保存非空行，空白分隔行表现为行号跳跃。
                break
            if not any(cells):
                if weee_catalog_header and records_started:
                    # 申请表的 WEEE 区块后紧接着还有电池/包装区块，
                    # 后者的“类别/品牌”说明不能作为 WEEE 项目。
                    break
                continue
            brands = [_clean_candidate(cells[column]) for column in brand_columns if column < len(cells)]
            categories = [_clean_candidate(cells[column]) for column in category_columns if column < len(cells)]
            brands = [value for value in brands if value]
            categories = [value for value in categories if value]
            if not brands and not categories:
                continue
            records_started = True
            if isinstance(row_number, int):
                previous_row_number = row_number
            records.append({
                "sheet_name": _text(sheet.get("sheet_name")) or "工作表",
                "row_number": row_number,
                "cells": cells[:24],
                "brand": " | ".join(brands),
                "category": " | ".join(categories),
                "raw_text": " | ".join(cell for cell in cells if cell),
                "record_type": "weee_catalog_preview",
            })
    return records


def _record_mentions_company(record: Dict[str, Any], company: str, filename: str = "") -> bool:
    """判断附件结构化行是否属于当前业务明细主体。

    WEEE 品牌/品类提取按邮件调用时容易把整张表的记录复制到每个主体。
    这里只使用行内已有的公司字段、单元格文本或附件文件名做保守匹配；
    没有证据时不猜测归属。
    """
    target = _norm_product(company)
    if not target:
        return True
    parts = [
        record.get("customer"),
        record.get("company"),
        record.get("raw_text"),
        record.get("brand"),
        record.get("category"),
        filename,
    ]
    parts.extend(record.get("cells") or [])
    haystack = _norm_product(" ".join(_text(part) for part in parts if _text(part)))
    return bool(haystack and target in haystack)


def _scope_body_to_company(body: str, company: str) -> str:
    """保留正文中明确包含当前主体的行，避免多主体邮件交叉污染。"""
    target = _norm_product(company)
    if not target:
        return _text(body)
    raw = re.sub(r"&(?:nbsp|amp);", " ", _text(body), flags=re.I)
    lines = [line.strip() for line in re.split(r"[\r\n]+", raw) if line.strip()]
    matched = [line for line in lines if target in _norm_product(line)]
    return "\n".join(matched)


def _is_header_alias(value: Any, aliases: Sequence[str]) -> bool:
    normalized = re.sub(r"\s+", "", _text(value).casefold())
    if not normalized:
        return False
    return any(
        normalized == re.sub(r"\s+", "", alias.casefold())
        or re.sub(r"\s+", "", alias.casefold()) in normalized
        for alias in aliases
    )


def _llm_source_text(subject: Any, body: Any, attachments: Iterable[Any]) -> str:
    """构造本地证据串，用于校验 LLM 重提取字段是否真实出现。"""
    parts = [_text(subject), _text(body)]
    for attachment in attachments or []:
        if not isinstance(attachment, dict):
            continue
        parts.append(_text(attachment.get("filename")))
        parts.append(_text(attachment.get("text_content") or attachment.get("text")))
        for record in attachment.get("structured_records") or []:
            if not isinstance(record, dict):
                continue
            parts.extend(_text(record.get(key)) for key in ("brand", "category", "raw_text"))
            parts.extend(_text(cell) for cell in record.get("cells") or [])
        for sheet in attachment.get("sheets") or []:
            if not isinstance(sheet, dict):
                continue
            for row in sheet.get("preview_rows") or sheet.get("rows") or []:
                if isinstance(row, dict):
                    parts.extend(_text(cell) for cell in (row.get("cells") or row.get("values") or []))
                elif isinstance(row, (list, tuple)):
                    parts.extend(_text(cell) for cell in row)
    return "\n".join(part for part in parts if part)


def _llm_item_has_evidence(item: Dict[str, Any], source_text: str) -> bool:
    """拒绝 LLM 编造的品牌/品类，只接收能在当前证据中定位的字段。"""
    brand = _clean_candidate(item.get("brand"))
    category = _clean_candidate(item.get("category"))
    if not brand and not category:
        return False
    haystack = _norm_product(source_text)
    if not haystack:
        return False
    if brand and _norm_product(brand) not in haystack:
        return False
    if category:
        # 组合品类可能包含尺寸说明；类别主体出现即可，尺寸不要求再次拼成整句。
        category_values = [category]
        category_values.extend(_combine_dimension_categories(_split_values(category)))
        category_values.extend(re.findall(r"第\s*[1-6一二三四五六]\s*类", category))
        if not any(_norm_product(value) in haystack for value in category_values if _clean_candidate(value)):
            return False
    return True


def extract_weee_items(
    subject: Any = "",
    body: Any = "",
    attachments: Any = None,
    project: Any = "",
    llm_client: Any = None,
    company: Any = "",
) -> Dict[str, Any]:
    """提取德国 WEEE 的品牌/品类证据。

    附件结构化行优先；正文/标题只有出现显式标签时才采用。
    ``items`` 每项至少包含 ``category`` 或被标记为待人工确认的证据。
    """
    attachments = attachments or []
    if not is_germany_weee(project, subject, body, attachments):
        return {"enabled": False, "items": [], "status": "not_applicable", "sources": []}

    scoped_company = _text(company)
    scoped_attachments: List[Dict[str, Any]] = []
    has_scoped_attachment_records = False
    for attachment in attachments:
        if not isinstance(attachment, dict):
            continue
        filename = _text(attachment.get("filename")) or "附件"
        records = list(attachment.get("structured_records") or [])
        # 阶段一同时保存 structured_records 和 sheets.preview_rows 时，二者
        # 描述的是同一物理行。优先使用结构化行；只有历史缓存没有结构化
        # 行时才回退到预览行，避免一行被拆成两份并交叉组合品牌/品类。
        if not records:
            records.extend(_sheet_preview_records(attachment))
        if not scoped_company or not records:
            scoped_attachments.append(attachment)
            continue
        matched_records = [
            record for record in records
            if isinstance(record, dict)
            and _record_mentions_company(record, scoped_company, filename)
        ]
        if matched_records:
            narrowed = dict(attachment)
            narrowed["structured_records"] = matched_records
            narrowed["sheets"] = []
            # 附件解析结果还会保留整张表的 text_content。若只缩小
            # structured_records 而不缩小这段文本，LLM 复检/重提取仍会看到
            # 其他主体的品牌，随后把整张表重新提回当前明细。
            scoped_text = []
            for record in matched_records:
                evidence = _text(record.get("raw_text")) or " | ".join(
                    _text(cell) for cell in record.get("cells") or []
                )
                if evidence:
                    scoped_text.append(evidence)
            if scoped_text:
                narrowed["text_content"] = "\n".join(scoped_text)
                if "text" in narrowed:
                    narrowed["text"] = narrowed["text_content"]
            scoped_attachments.append(narrowed)
            has_scoped_attachment_records = True
    if scoped_company:
        attachments = scoped_attachments

    # 有主体对应的附件表格时，附件行是当前明细的权威证据，正文只会把
    # 全邮件品牌再次复制进来；丢弃全局正文，避免“主体数 × 全部品牌”的笛卡尔积。
    body_for_scope = _text(body)
    subject_for_scope = _text(subject)
    if scoped_company and has_scoped_attachment_records:
        body_for_scope = ""
        subject_for_scope = _scope_body_to_company(subject_for_scope, scoped_company)
    elif scoped_company:
        body_for_scope = _scope_body_to_company(body_for_scope, scoped_company)
        subject_for_scope = _scope_body_to_company(subject_for_scope, scoped_company)
        if not body_for_scope and not subject_for_scope:
            body_for_scope = _text(body)
            subject_for_scope = _text(subject)

    items: List[Dict[str, Any]] = []
    sources: List[str] = []
    for attachment in attachments:
        if not isinstance(attachment, dict):
            continue
        filename = _text(attachment.get("filename")) or "附件"
        attachment_records = list(attachment.get("structured_records") or [])
        if not attachment_records:
            attachment_records.extend(_sheet_preview_records(attachment))
        for record in attachment_records:
            if not isinstance(record, dict):
                continue
            brands = _record_values(record, _BRAND_KEYS)
            categories = _combine_dimension_categories(_record_category_values(record))
            evidence = _text(record.get("raw_text")) or " | ".join(_text(c) for c in record.get("cells") or [])
            # 例示行只能作为模板说明。若其中没有真实分类表关键词，整行跳过，
            # 避免把 Apple、10g、3500个等示例拆成业务项目。
            if _is_example_record(evidence) and not any(
                _category_alias_hits(value) or _category_id_from_label(value)
                for value in categories
            ):
                continue
            if not brands and not categories:
                continue
            source = (
                f"附件表格：{filename} / {_text(record.get('sheet_name')) or '工作表'} "
                f"第{_text(record.get('row_number')) or '?'}行"
            )
            for brand in brands or [""]:
                for category in categories or [""]:
                    items.append({
                        "brand": brand,
                        "category": category,
                        "source": source,
                        "evidence": evidence,
                        "confidence": "high" if category else "medium",
                    })
            sources.append(source)

    mail_text = _text(subject_for_scope) + "\n" + _text(body_for_scope)

    def unique_pairs(pairs: Iterable[tuple[str, str]]) -> List[tuple[str, str]]:
        result: List[tuple[str, str]] = []
        seen = set()
        for label, value in pairs:
            cleaned = _clean_candidate(value)
            key = (_norm(label), _norm(cleaned))
            if cleaned and key not in seen:
                seen.add(key)
                result.append((_text(label), cleaned))
        return result

    # 邮件里既有“品牌：X/品类：Y”，也常见“新增品牌X 第六类”这种没有冒号的写法。
    brand_pairs = unique_pairs(
        _label_value_pairs(mail_text, _BRAND_KEYS)
        + _inline_marker_pairs(mail_text, _INLINE_BRAND_MARKERS)
    )
    category_pairs = unique_pairs(
        _label_value_pairs(mail_text, _CATEGORY_KEYS)
        + _inline_marker_pairs(mail_text, _INLINE_CATEGORY_MARKERS)
    )
    # 没有显式“类别/品类”标签时，保留“第 1 类/第六类”作为原始品类证据。
    if not category_pairs:
        category_pairs = _category_class_pairs(mail_text)
    # “品牌 X 冷却设备”没有字段标签时，只有同一行存在品牌标记且分类表能
    # 唯一归类，才把自然语言品类补出来，避免扫描模板正文造成大量误报。
    if not category_pairs and brand_pairs:
        category_pairs = [("产品分类表命中", value) for value in _context_category_values(mail_text)]

    if brand_pairs and category_pairs:
        # 同一封邮件通常按“品牌1/品类1；品牌2/品类2”排列；按出现顺序配对，
        # 数量不一致时只配已有项，其余仍保留为待人工核对，不跨行猜测。
        pair_count = max(len(brand_pairs), len(category_pairs))
        for index in range(pair_count):
            brand_label, brand_value = brand_pairs[index] if index < len(brand_pairs) else ("", "")
            category_label, category_value = category_pairs[index] if index < len(category_pairs) else ("", "")
            items.append({
                "brand": _clean_candidate(brand_value),
                "category": _clean_candidate(category_value),
                "source": "邮件正文/标题",
                "evidence": "；".join(part for part in (
                    f"{brand_label}：{brand_value}" if brand_label else "",
                    f"{category_label}：{category_value}" if category_label else "",
                ) if part),
                "confidence": "medium",
            })
            sources.append("邮件正文/标题")
    else:
        for label, value in brand_pairs:
            items.append({"brand": _clean_candidate(value), "category": "", "source": "邮件正文/标题", "evidence": f"{label}：{value}", "confidence": "medium"})
            sources.append("邮件正文/标题")
        for label, value in category_pairs:
            items.append({"brand": "", "category": _clean_candidate(value), "source": "邮件正文/标题", "evidence": f"{label}：{value}", "confidence": "medium"})
            sources.append("邮件正文/标题")

    # 先让 LLM 复检规则提取结果；复检不通过时再按同一套字段规则重提取。
    # LLM 结果必须经过本地证据校验，不能直接覆盖成模型臆测值。
    llm_review: Dict[str, Any] = {}
    llm_reextract: Dict[str, Any] = {}
    extraction_method = "规则提取"
    if llm_client is not None and getattr(llm_client, "enabled", False):
        for index, item in enumerate(items):
            if isinstance(item, dict):
                item.setdefault("item_id", str(index))
        try:
            review_result = llm_client.review_weee_extraction(
                subject=_text(subject_for_scope),
                body=_text(body_for_scope),
                attachments=attachments,
                project=_text(project),
                items=items,
            )
        except Exception:
            review_result = None
        if isinstance(review_result, dict):
            llm_review = review_result
        review_ok = bool(
            llm_review.get("status") == "valid"
            and llm_review.get("confidence") in {"high", "medium"}
        )
        if review_ok:
            extraction_method = "规则提取·LLM复检通过"
        else:
            try:
                reextract_result = llm_client.extract_weee_fields(
                    subject=_text(subject_for_scope),
                    body=_text(body_for_scope),
                    attachments=attachments,
                    project=_text(project),
                    initial_items=items,
                    review=llm_review,
                )
            except Exception:
                reextract_result = None
            if isinstance(reextract_result, dict):
                llm_reextract = reextract_result
            source_text = _llm_source_text(subject_for_scope, body_for_scope, attachments)
            candidates = [
                candidate for candidate in (llm_reextract.get("items") or [])
                if isinstance(candidate, dict) and _llm_item_has_evidence(candidate, source_text)
            ]
            if candidates:
                items = candidates
                extraction_method = "LLM重新提取"
            else:
                extraction_method = "规则提取·LLM复检未通过"

    # 按品牌/品类证据合并，避免同一行被附件和正文重复展示。
    merged: Dict[tuple[str, str], Dict[str, Any]] = {}
    for item in items:
        brand = _clean_candidate(item.get("brand"))
        category = _clean_candidate(item.get("category"))
        if _is_dimension_only(category):
            # 没有同一记录中的主体类别时，尺寸限制本身不构成 WEEE 类别。
            continue
        if not brand and not category:
            continue
        # “小型设备”与“ 小型设备，单边尺寸小于50cm ”是同一业务项目；
        # 尺寸只作为证据保留，不再生成第二张卡片。
        category_class = classify_weee_product(category).get("category_class", "")
        category_key = (
            f"class:{category_class}"
            if category_class
            else _norm(re.sub(
                r"[，,;；|]?\s*(?:(?:单边|外部|最大|最小)?\s*尺寸\s*)?(?:小于|不超过|不大于|≤|<=|低于|少于)?\s*"
                r"\d+(?:\.\d+)?\s*(?:cm|厘米|mm|毫米|m|米)(?:以下|以内)?",
                "",
                category,
                flags=re.I,
            ))
        )
        key = (_norm(brand), category_key or _norm(category))
        current = merged.setdefault(key, {"brand": brand, "category": category, "sources": [], "evidences": [], "confidence": item.get("confidence", "medium")})
        if brand and not current.get("brand"):
            current["brand"] = brand
        if category and len(category) > len(current.get("category") or ""):
            current["category"] = category
        for name in ("source", "evidence"):
            value = _text(item.get(name))
            target_key = "sources" if name == "source" else "evidences"
            if value and value not in current[target_key]:
                current[target_key].append(value)

    result_items = list(merged.values())
    for item in result_items:
        item["evidence"] = "；".join(item.get("evidences") or [])
        if llm_client is not None and getattr(llm_client, "enabled", False):
            item["extraction_method"] = extraction_method
        if llm_review:
            item["llm_review_status"] = _text(llm_review.get("status"))
            item["llm_review_confidence"] = _text(llm_review.get("confidence"))
        if llm_reextract:
            item["llm_reextract_reason"] = _text(llm_reextract.get("reason"))
        # 不覆盖原始品类文字；分类表归类只是一个可追溯的派生字段。
        # 品牌不能代替品类参与分类。品牌存在而品类为空时必须保持待人工，
        # 否则会把品牌名中恰好出现的产品词误判成 WEEE 类别。
        category_text = item.get("category") or ""
        classification = classify_weee_product(category_text)
        item["category_original"] = category_text
        item["category_class"] = classification.get("category_class", "")
        item["category_class_name"] = classification.get("category_class_name", "")
        item["category_class_status"] = classification.get("status", "unmatched")
        item["category_candidates"] = classification.get("candidates", [])
        item["category_class_evidence"] = classification.get("evidence", [])
        item["category_rule_source"] = WEEE_RULE_SOURCE
        components = decompose_weee_category(category_text)
        if components:
            item["category_components"] = components

    # 规则没有唯一结果时才请求 LLM。LLM 仅写入候选建议，不把项目标记为
    # matched；操作人员仍需对每一个品牌/品类项目单独选择并确认。
    if llm_client is not None and getattr(llm_client, "enabled", False):
        uncertain = [
            item for item in result_items
            if item.get("category_class_status") in {"unmatched", "ambiguous"}
            and (item.get("brand") or item.get("category"))
        ]
        if uncertain:
            for index, item in enumerate(result_items):
                item.setdefault("item_id", str(index))
            try:
                suggestions = llm_client.classify_weee_categories(uncertain)
            except Exception:
                suggestions = []
            by_id = {str(item.get("item_id")): item for item in uncertain}
            for suggestion in suggestions or []:
                target = by_id.get(str(suggestion.get("item_id")))
                category_class = _text(suggestion.get("category_class"))
                if not target or category_class not in WEEE_CATEGORY_DEFINITIONS:
                    continue
                candidate = {
                    "category_class": category_class,
                    "category_class_name": WEEE_CATEGORY_DEFINITIONS[category_class]["name"],
                    "score": 0,
                    "source": "LLM辅助建议",
                }
                if not any(
                    str(existing.get("category_class")) == category_class
                    for existing in target.get("category_candidates") or []
                    if isinstance(existing, dict)
                ):
                    target.setdefault("category_candidates", []).append(candidate)
                target["llm_category_suggestion"] = {
                    "category_class": category_class,
                    "category_class_name": candidate["category_class_name"],
                    "confidence": _text(suggestion.get("confidence")) or "low",
                    "reason": _text(suggestion.get("reason")),
                }
    if not result_items:
        status = "pending"
    elif any(
        not item.get("category")
        or item.get("category_class_status") != "matched"
        for item in result_items
    ):
        status = "pending"
    else:
        status = "ready"
    return {
        "enabled": True,
        "items": result_items,
        "status": status,
        "sources": _dedupe(sources),
        "extraction_method": extraction_method,
        "llm_review": llm_review,
        "llm_reextract": llm_reextract,
    }


def _json_items(value: Any) -> List[Dict[str, Any]]:
    if isinstance(value, list):
        return [item for item in value if isinstance(item, dict)]
    text = _text(value)
    if not text:
        return []
    try:
        parsed = json.loads(text)
    except (TypeError, ValueError, json.JSONDecodeError):
        return []
    return [item for item in parsed if isinstance(item, dict)] if isinstance(parsed, list) else []


def _workorder_values(row: Dict[str, Any], aliases: Sequence[str]) -> List[str]:
    values: List[str] = []
    for key, value in row.items():
        key_norm = re.sub(r"\s+", "", _text(key).casefold())
        if any(re.sub(r"\s+", "", alias.casefold()) in key_norm for alias in aliases):
            if isinstance(value, (dict, list, tuple, set)):
                nested = _nested_alias_values(value, aliases)
                values.extend(nested or _flatten_values(value))
            else:
                values.extend(_split_values(value))
    return _dedupe(values)


def compare_weee_categories(items: Any, workorders: Iterable[Dict[str, Any]] = ()) -> Dict[str, Any]:
    """逐项比较邮件品类和注册工单品类。"""
    mail_items = _json_items(items)
    orders = [row for row in (workorders or []) if isinstance(row, dict)]
    order_categories = _dedupe(value for row in orders for value in _workorder_values(row, _WORKORDER_CATEGORY_KEYS))
    order_brands = _dedupe(value for row in orders for value in _workorder_values(row, _BRAND_KEYS))
    order_classifications = []
    for value in order_categories:
        classification = classify_weee_product(value)
        if classification.get("status") == "matched":
            order_classifications.append({
                "raw": value,
                "category_class": classification.get("category_class"),
                "category_class_name": classification.get("category_class_name"),
            })
    checked: List[Dict[str, Any]] = []
    for item in mail_items:
        brand = _clean_candidate(item.get("brand"))
        category = _clean_candidate(item.get("category"))
        if not category:
            checked.append({**item, "status": "pending", "reason": "邮件未提取到明确品类"})
            continue
        # LLM 只是候选建议；只有规则命中或人工保存为 matched 的项目才可
        # 参与工单类别自动比对，避免未经确认的建议被当成最终结果。
        mail_class = item.get("category_class") if item.get("category_class_status") == "matched" else ""
        if not mail_class:
            inferred = classify_weee_product(category)
            if inferred.get("status") == "matched":
                mail_class = inferred.get("category_class")
        category_hit = any(
            _norm(category) == _norm(value)
            or _norm(category) in _norm(value)
            or _norm(value) in _norm(category)
            for value in order_categories
        )
        class_hit = bool(mail_class and any(item2.get("category_class") == mail_class for item2 in order_classifications))
        category_hit = category_hit or class_hit
        brand_hit = not brand or not order_brands or any(_norm(brand) == _norm(value) or _norm(brand) in _norm(value) or _norm(value) in _norm(brand) for value in order_brands)
        if not order_categories:
            status, reason = "pending", "注册工单未返回品类明细字段"
        elif category_hit and brand_hit:
            status, reason = "matched", "工单品类明细已找到"
        elif category_hit:
            status, reason = "pending", "品类已找到但品牌需人工核对"
        else:
            status, reason = "missing", "注册工单品类明细中未找到对应品类"
        checked.append({**item, "status": status, "reason": reason})
    if not checked:
        overall = "pending"
        reason = "邮件未提取到明确品牌/品类"
    elif any(item["status"] == "missing" for item in checked):
        overall, reason = "missing", "至少一项邮件品类未出现在注册工单品类明细"
    elif any(item["status"] == "pending" for item in checked):
        overall, reason = "pending", "品类明细不完整，需人工核对"
    else:
        overall, reason = "matched", "邮件品类与注册工单品类明细全部对应"
    return {
        "enabled": bool(mail_items),
        "status": overall,
        "reason": reason,
        "items": checked,
        "workorder_categories": order_categories,
        "workorder_category_classes": order_classifications,
        "workorder_brands": order_brands,
        "workorder_count": len(orders),
    }


def audit_row(row: Dict[str, Any], workorders: Iterable[Dict[str, Any]] = ()) -> Dict[str, Any]:
    """从阶段一/阶段二行读取专项字段，便于 Excel 和工作台复用。"""
    project = row.get("项目") or row.get("标准化项目名称") or row.get("标准化项目")
    subject = row.get("subject") or row.get("邮件主题") or row.get("主题")
    body = row.get("body_text") or row.get("邮件正文原文") or row.get("正文(精简)")
    attachments = row.get("_attachments") or []
    extracted = {"enabled": _text(row.get("德国WEEE专项")) == "是", "items": _json_items(row.get("德国WEEE品类明细")), "status": "pending"}
    if not extracted["enabled"]:
        extracted = extract_weee_items(subject, body, attachments, project)
    return compare_weee_categories(extracted.get("items") or [], workorders) if extracted.get("enabled") else {"enabled": False, "status": "not_applicable", "items": []}
