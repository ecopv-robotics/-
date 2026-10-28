"""German battery evidence extraction and independent item-by-item review."""
import hashlib
import json
import re

from modules.brand_extract import extract_brand_values


BATTERY_CLASSES = ("设备电池", "轻型交通工具电池", "工业电池", "启动电池", "汽车电池")
_BATTERY_CONTEXT = re.compile(r"电池|batter(?:y|ies|ie|ien)|BattG|BattVO|BattDG", re.I)
_WEEE_CONTEXT = re.compile(r"WEEE|ElektroG|电子电气", re.I)
_BRAND_LABEL = re.compile(r"(?:新增品牌|品牌名称|品牌名|品牌|brand(?:\s+name)?|marke)\s*[:：]?\s*([^,，;；|\n\r+＋]{1,90})", re.I)
_TYPE_LABEL = re.compile(r"(?:电池(?:类型|类别|品类|分类)|原始品类|注册类别|battery\s*(?:type|category)|batteriekategorie)\s*[:：]?\s*([^,，;；|\n\r+＋]{1,90})", re.I)
_PLACEHOLDER = re.compile(r"^(?:品牌|品牌名称|电池类型|电池类别|示例|例如|例|待填写|请选择|必填|选填|brand|category|type)(?:\s*[:：].*)?$", re.I)
# 品牌/品类候选不允许携带链接或裸域名（反馈 R10：品牌混入亚马逊链接）。
_URL_NOISE_RE = re.compile(
    r"https?://\S+"
    r"|www\.\S+"
    r"|(?<![\w@.])[A-Za-z0-9][A-Za-z0-9-]*(?:\.[A-Za-z0-9-]+)*\.(?:com|de|cn|net|eu|org|shop|top|site|online)(?:/[^\s,，;；|]*)?",
    re.I,
)


def _strip_url_noise(value):
    """剥离品牌/品类候选中的链接与裸域名，剩余部分交给 _clean。"""
    text = _URL_NOISE_RE.sub(" ", str(value or ""))
    return re.sub(r"\s{2,}", " ", text).strip(" \t:：,，;；|+＋-—_")


def _clean_field_value(value):
    """品牌/品类专用清洗：先剥链接噪声，再做通用清洗。"""
    return _clean(_strip_url_noise(value))
_CLASS_PATTERNS = (
    ("汽车电池", re.compile(r"电动汽车电池|电动车(?:辆)?电池|电动汽车动力电池|Elektrofahrzeugbatterie|(?:Anlage[\s_-]*)?FZG|electric vehicle batter", re.I)),
    ("轻型交通工具电池", re.compile(r"轻型交通工具电池|轻型交通电池|轻型车辆电池|LV[-\s]?Batter|Leichte\s+Verkehrsmittel|light means of transport batter|(?:Anlage[\s_-]*)?LV(?:[-_.\s]|$)", re.I)),
    ("启动电池", re.compile(r"汽车启动电池|启动电池|起动电池|Starterbatterie|starting.*battery|(?:Anlage[\s_-]*)?STB(?:[-_.\s]|$)", re.I)),
    ("工业电池", re.compile(r"工业电池|Industriebatterie|industrial batter|(?:NB[-_\s]?)?BESS|stationäre[sn]?\s+Batterie|储能电池", re.I)),
    ("设备电池", re.compile(r"设备电池|便携式电池|Gerätebatterie|portable batter", re.I)),
)


def classify_battery_category(value):
    """Return a class only when the same evidence names exactly one category."""
    text = str(value or "").strip()
    hits = [name for name, pattern in _CLASS_PATTERNS if pattern.search(text)]
    return hits[0] if len(hits) == 1 else ""


def _clean(value):
    text = re.sub(r"&nbsp;|\s+", " ", str(value or ""), flags=re.I).strip(" \t:：,，;；|+＋")
    return "" if not text or _PLACEHOLDER.fullmatch(text) else text[:1000]


def _norm(value):
    return re.sub(r"[\s\W_]+", "", str(value or "").casefold())


def _company_matches(record, company, filename):
    if not company:
        return True
    owner = _clean(record.get("customer") or record.get("company"))
    if owner:
        return _norm(owner) == _norm(company) or _norm(owner) in _norm(company) or _norm(company) in _norm(owner)
    # A shared table without an owner must not be copied to every company.
    return bool(_norm(company) and _norm(company) in _norm(filename))


def _sheet_battery_records(attachment, company):
    """Read selected values below battery section labels in EPR worksheets."""
    filename = _clean(attachment.get("filename"))
    for sheet in attachment.get("sheets") or []:
        if not isinstance(sheet, dict):
            continue
        sheet_name = _clean(sheet.get("sheet_name"))
        source_name = _clean(sheet.get("member_path")) or filename
        rows = sheet.get("preview_rows") or sheet.get("rows") or []
        normalized = []
        for row in rows:
            if isinstance(row, dict):
                normalized.append((row.get("row_number") or row.get("index") or "?", [str(cell or "").strip() for cell in row.get("cells") or row.get("values") or []]))
            elif isinstance(row, (list, tuple)):
                normalized.append(("?", [str(cell or "").strip() for cell in row]))
        if not normalized:
            continue
        whole_sheet = " ".join(cell for _, cells in normalized for cell in cells if cell)
        if company and _norm(company) not in _norm(source_name + whole_sheet):
            continue
        active = bool(_BATTERY_CONTEXT.search(sheet_name) and not _WEEE_CONTEXT.search(sheet_name))
        headers = None
        for index, (row_number, cells) in enumerate(normalized):
            row_text = " ".join(cells)
            if re.search(r"WEEE产品信息|包装(?:法)?产品信息|WEEE\s*(?:product|information)", row_text, re.I):
                active, headers = False, None
                continue
            if re.search(r"电池(?:法)?产品信息|电池类型信息|battery\s*(?:product\s*)?information|Batterie.*Information", row_text, re.I):
                active, headers = True, None
                continue
            if not active:
                continue
            brand_cols = [column for column, value in enumerate(cells) if re.fullmatch(r"\*?\s*(?:品牌|品牌名称|brand|brand name|marke)\s*", value, re.I)]
            category_cols = [column for column, value in enumerate(cells) if re.fullmatch(r"\*?\s*(?:类别|品类|电池类型|电池类别|battery type|battery category|batteriekategorie)\s*", value, re.I)]
            if brand_cols or category_cols:
                headers = (brand_cols, category_cols)
                continue
            if not headers:
                continue
            brands = [_clean(cells[column]) for column in headers[0] if column < len(cells)]
            categories = [_clean(cells[column]) for column in headers[1] if column < len(cells)]
            brand = next((value for value in brands if value), "")
            category = next((value for value in categories if value), "")
            if not (brand or category) or re.search(r"^(?:示例|例如|例\s*[:：]|example|e\.g\.)", row_text, re.I):
                continue
            yield {"brand": brand, "category": category, "sheet_name": sheet_name,
                   "row_number": row_number, "raw_text": row_text,
                   "attachment_name": source_name, "record_type": "battery_section"}


_BRAND_FIELD_LABEL = re.compile(
    r"(?:新增品牌|原主体注册品牌|原注册品牌|注册品牌|品牌名称|品牌名|品牌|brand(?:\s*name)?|marke)",
    re.I,
)
# 字段行必须**以品牌标签开头**：申请表是「品牌 | 值」一行一字段的布局，
# 而长说明文里到处是 brand/品牌 字样。旧写法用 search 任意位置匹配，把
# 「… t in Ireland? … brand …」整段截成了品牌（实测 183 个申请表抽出 80 种
# 取值，绝大多数是表头/说明文）。
_BRAND_FIELD_LINE = re.compile(
    r"^[\s\d.、)）\-—_*＊]*(?P<label>新增品牌|原主体注册品牌|原注册品牌|注册品牌|"
    r"品牌名称|品牌名|品牌|brand\s*name|brand|marke)"
    r"\s*[*＊]?\s*(?:[（(][^）)]{0,80}[）)]\s*)?[:：]?\s*(?P<value>.+)$",
    re.I,
)
# 品牌字段行里的填写说明/清单文字，不是品牌值。
_BRAND_VALUE_NOISE = re.compile(
    r"请|請|提供|填写|填寫|须知|須知|如需|多个|多個|示例|例如|注[:：]|以上|以下|选择|選擇|确认|確認",
    re.I,
)
_FORM_FIELD_WORDS = re.compile(
    r"必填|选填|選填|logo|预估|預估|销量|销量|重量|尺寸|链接|鏈接|型号|型號|规格|規格|备注|備註|"
    r"数量|數量|占比|商标|商標|register|registration|\bnr\b|号码|incoterm|invoice|http|www",
    re.I,
)
# 纯中文候选里的业务名词：品牌通常不叫「单个电池重量」「商标注册编号」。
_BRAND_CN_NOISE = re.compile(
    r"重量|数量|数量|个数|尺寸|时间|编号|商标|注册|德文|英文|中文|汉字|生产|销售|投放|商用|家用|"
    r"预估|类型|类别|品类|名称|地址|电话|邮箱|日期|备注|单位|价格|金额|说明|示例|填写|必填|选填|"
    r"链接|页面|店铺|公司|法人|身份|证件|营业|执照|翻译|市场|国家|"
    r"对应|相应|上述|如下|详见|参见|同上|以下|以上|其他|其它"
)
# 品牌形态：英文/数字型（允许 Co., Ltd. 的空格与半角逗号）或 2~10 个纯中文字。
_BRAND_SHAPE_ASCII = re.compile(r"[A-Za-z0-9][A-Za-z0-9 .,&'’\-]{0,39}")
_BRAND_SHAPE_CN = re.compile(r"[\u4e00-\u9fff]{2,10}")
# 净重/尺寸这类测量值会填在「品牌」列附近，被错认成品牌（实测 32.9g）。
_MEASUREMENT_RE = re.compile(
    r"^\d+(?:[.,]\d+)?\s*(?:g|kg|mg|t|cm|mm|m|l|ml|w|kw|v|%|个|只|件|pcs?)\.?$", re.I
)


def _is_junk_brand(value):
    """明显不是品牌的候选：表头/填写说明/业务名词/测量值。

    「品牌」这两个字在申请表和粘贴正文里到处都是，只看标签会大量误抽：
    `logo（必填） 单边最大尺寸（必填） 产品链接`、`单个电池重量`、
    `Brand register Nr.`、`32.9g`、`品牌1` 这些都不是品牌。
    """
    text = _clean(value)
    if not text:
        return True
    if re.search(r"[（(【\[）)】\]]", text):
        return True
    if _FORM_FIELD_WORDS.search(text) or _BRAND_CN_NOISE.search(text):
        return True
    if _MEASUREMENT_RE.match(text):
        return True
    # 纯日期/纯编号（主题「新增品牌-2026.8.24」会被截成 2026.8.24）不是品牌。
    if re.fullmatch(r"[\d\s./\-:年月日]+", text):
        return True
    if re.fullmatch(r"品牌\s*\d*", text):
        return True
    return False


def _brand_value_ok(value):
    """品牌候选是否可信（排除噪声 + 形态校验）。正文、表格、附件三条路径共用。"""
    text = _clean(value)
    if not text or len(text) > 40:
        return False
    if _BRAND_VALUE_NOISE.search(text) or _is_junk_brand(text):
        return False
    if _BRAND_SHAPE_CN.fullmatch(text):
        return True
    if _BRAND_SHAPE_ASCII.fullmatch(text):
        # 单个字母/纯符号不是品牌（「t」「ts)」这类是说明文被截断的残留）。
        return bool(re.search(r"[A-Za-z]{2,}|\d", text))
    return False


# 申请表品牌栏常写明「多个品牌，请用逗号隔开」，一个格子里会填两个品牌。
# 半角逗号只有在后面不是公司后缀时才当分隔符，否则「TianTian Co., Ltd.」会被拆坏。
_COMPANY_SUFFIX = re.compile(
    r"^(?:co|ltd|limited|inc|llc|gmbh|corp|corporation|sa|s\.a|ag|bv|b\.v|oy|plc)\.?$", re.I
)


def _split_brand_cell(value):
    """把「mesterth,Naxidf」这类多品牌单元格拆开；公司后缀不拆。"""
    text = _clean(value)
    if "," not in text:
        return [text]
    parts = [part.strip(" .") for part in text.split(",")]
    if len(parts) < 2 or any(not part for part in parts):
        return [text]
    if any(_COMPANY_SUFFIX.match(part) for part in parts):
        return [text]
    return parts


def _sheet_brand_field_values(attachment, company):
    """申请表「品牌」字段行 → [(品牌, 证据行)]。

    申请表把品牌填在字段行里（`品牌 *（多个品牌，请用逗号隔开） KDTEND`），
    解析成结构化记录后只剩公司字段，品牌值读不到 —— 表现为「附件里明明写了
    品牌，界面品牌是空的」（反馈：德国电池法 重庆振丰维 KDTEND）。
    这里对工作表文本补一路证据；类别仍留空，由人工在五选一下拉确认。
    """
    values = []
    seen = set()
    company_key = _norm(company)
    # 申请表通常把主体写在文件名里（「XX公司＋电池法申请表.xlsx」），表内也可能写。
    # 两处都没有该公司名时，无法判断这张表属于谁，宁可不采信，避免把别家
    # 申请表的品牌搬到这封邮件上。
    bound_to_company = bool(company_key) and company_key in _norm(attachment.get("filename"))
    for sheet in attachment.get("sheets") or []:
        if not isinstance(sheet, dict):
            continue
        text = str(sheet.get("text") or "")
        if not text:
            continue
        if (company_key and not bound_to_company
                and company_key not in _norm(text)):
            continue
        for line in re.split(r"[\r\n]+", text):
            line = line.strip()
            if not line or len(line) > 160:
                continue
            match = _BRAND_FIELD_LINE.match(line)
            if not match:
                continue
            for brand in extract_brand_values(line):
                if not _brand_value_ok(brand):
                    continue
                for single in _split_brand_cell(brand):
                    key = _norm(single)
                    if not key or key in seen:
                        continue
                    seen.add(key)
                    values.append((single, line[:300]))
    return values


def extract_battery_items(subject="", body="", attachments=None, project="", company=""):
    """Extract battery brand/type from the selected business and attachment evidence.

    Contract codes identify candidates, not a human confirmation. The generic OfH
    contract names all five categories, so it never establishes a specific one.
    """
    if not is_german_battery(project):
        return {"enabled": False, "items": [], "status": "not_applicable"}
    attachments = attachments or []
    items = []
    form_brands = []
    contract_classes = {}
    for attachment in attachments:
        if not isinstance(attachment, dict):
            continue
        form_brands.extend(_sheet_brand_field_values(attachment, company))
        filename = _clean(attachment.get("filename"))
        member_names = [filename]
        member_names.extend(str(sheet.get("member_path") or "") for sheet in attachment.get("sheets") or [] if isinstance(sheet, dict))
        member_names.extend(str(record.get("attachment_name") or "") for record in attachment.get("structured_records") or [] if isinstance(record, dict))
        member_names.extend(re.findall(r"(?m)\[([^\]\r\n]{1,500}\.(?:pdf|docx|xlsx|xls|csv|txt))\]", str(attachment.get("text_content") or ""), re.I))
        for member_name in dict.fromkeys(member_names):
            lower_name = member_name.casefold()
            if re.search(r"(?:vertrag|合同|anlage|grs)", lower_name, re.I) and not re.search(r"(?:^|[/\\_\s-])ofh(?:[/\\_\s.-]|$)", lower_name, re.I):
                category_class = classify_battery_category(member_name)
                if category_class:
                    contract_classes.setdefault(category_class, member_name)
        records = list(attachment.get("structured_records") or [])
        records.extend(_sheet_battery_records(attachment, company))
        for record in records:
            if not isinstance(record, dict):
                continue
            record_filename = _clean(record.get("attachment_name")) or filename
            if record.get("record_type") != "battery_section" and not _company_matches(record, company, record_filename):
                continue
            brand = _clean_field_value(record.get("brand"))
            category = _clean_field_value(record.get("category"))
            context = " ".join(str(record.get(key) or "") for key in ("project", "business", "sheet_name", "record_type")) + " " + record_filename
            # 申请表的电池区块常被解析成通用目录记录（如 record_type=weee_catalog、
            # sheet=Sheet1），上下文里根本没有“电池”字样，旧逻辑会整条丢弃，
            # 即使品牌已经解析出来（反馈：品牌普遍提不出来）。业务本身是德国
            # 电池法时，带品牌的目录记录就是电池证据；类别无法归入五类则留空，
            # 由人工在五选一下拉里确认。
            recognized_class = classify_battery_category(category) or classify_battery_category(brand)
            # “锂电池/储能电池”这类写法也能归进 WEEE 的设备类别，但字面已经
            # 写明是电池，应优先按电池证据处理，不能再被 WEEE 上下文挡掉。
            literal_battery = bool(re.search(r"电池|batter", f"{category} {brand}", re.I))
            if not recognized_class and not literal_battery:
                weee_like = False
                if category:
                    from modules.weee_category_audit import classify_weee_product
                    weee_like = bool(classify_weee_product(category).get("category_class"))
                if weee_like:
                    continue
                if _WEEE_CONTEXT.search(context) and not _BATTERY_CONTEXT.search(context):
                    continue
                # 进入本函数时 project 已确认是德国电池法，因此“品牌 + 类别”齐全、
                # 且类别不属于 WEEE 的表格行本身就是电池证据；分类留空由人工
                # 五选一确认。
                if not _BATTERY_CONTEXT.search(context) and not (brand and category):
                    continue
            if not (brand or category):
                continue
            source = f"附件表格：{record_filename} / {_clean(record.get('sheet_name')) or '工作表'} 第{record.get('row_number') or '?'}行"
            items.append({"brand": brand, "category": category,
                          "category_class": classify_battery_category(category),
                          "source": source, "evidence": _clean(record.get("raw_text")) or " | ".join(map(str, record.get("cells") or [])),
                          "confidence": "high" if category else "medium"})

    # 无附件邮件的正文常把申请表粘成表格：转文本后每个单元格一行，
    # 或同一行用制表符分列。此时“品牌”“电池类型”只是表头列名，值在下一行
    # 同一列，普通“品牌：X”标签匹配读不到，品牌会整列丢失。
    from modules.weee_category_audit import _table_block_pairs
    table_brands, table_categories = _table_block_pairs(str(body or ""))
    if table_brands or table_categories:
        pair_count = max(len(table_brands), len(table_categories))
        for index in range(pair_count):
            brand = table_brands[index] if index < len(table_brands) else ""
            category = table_categories[index] if index < len(table_categories) else ""
            # 表格列错位时，重量/说明文字会落进品牌列（实测 32.9g、
            # 「对应的品牌logo（必填）」），必须和附件路径用同一套校验。
            if brand and not _brand_value_ok(brand):
                brand = ""
            if not (brand or category):
                continue
            items.append({"brand": brand, "category": category,
                          "category_class": classify_battery_category(category),
                          "source": "邮件正文表格",
                          "evidence": f"表格品牌列：{brand}；表格类别列：{category}",
                          "confidence": "medium"})

    # Only battery-specific lines can contribute mail fields in a mixed WEEE mail.
    lines = [line.strip() for line in re.split(r"[\r\n]+", str(subject or "") + "\n" + str(body or "")) if line.strip()]
    for line in lines:
        if company and _norm(company) not in _norm(line):
            continue
        if not _BATTERY_CONTEXT.search(line):
            continue
        # 混合邮件里同一行常同时写 WEEE 与电池法；只要出现「电池法」就按电池
        # 证据处理，不能因为该行没有电池类别而被整行跳过。
        explicit_battery = re.search(r"电池法|BattG|Batteriegesetz", line, re.I)
        if (_WEEE_CONTEXT.search(line) and not classify_battery_category(line)
                and not explicit_battery):
            continue
        brands = extract_brand_values(line)
        categories = [_clean_field_value(match.group(1)) for match in _TYPE_LABEL.finditer(line)]
        # 正文里的「品牌」后面常跟着表头/说明文（logo（必填）…、单个电池重量），
        # 只靠标签会把这些当成品牌名，必须过同一套形态校验。
        brands = [value for value in brands if _brand_value_ok(value)]
        categories = [value for value in categories if value]
        if not categories:
            category_class = classify_battery_category(line)
            if category_class:
                categories = [category_class]
        for index in range(max(len(brands), len(categories))):
            brand = brands[index] if index < len(brands) else ""
            category = categories[index] if index < len(categories) else ""
            if brand or category:
                items.append({"brand": brand, "category": category, "category_class": classify_battery_category(category),
                              "source": "邮件正文/标题", "evidence": line[:1000], "confidence": "medium"})

    # 品牌常单独成行（「新增品牌：Fyrove」「品牌名称：Goffii」），那一行既没有
    # 公司名也没有「电池」字样，会被上面的上下文过滤整行丢掉 —— 现象就是
    # 「正文里明明写了品牌，卡片品牌是空的」。只有在一条品牌都没抽到时，才回退
    # 到「标签带值就采信」；这样多主体邮件不会因放宽而把别家品牌挂过来。
    if not any(_clean(item.get("brand")) for item in items):
        for line in lines:
            if _WEEE_CONTEXT.search(line) and not _BATTERY_CONTEXT.search(line):
                continue
            for value in extract_brand_values(line):
                if not _brand_value_ok(value):
                    continue
                for brand in _split_brand_cell(value):
                    items.append({"brand": brand, "category": "", "category_class": "",
                                  "source": "邮件正文品牌行", "evidence": line[:300],
                                  "confidence": "medium"})

    merged = {}
    for item in items:
        key = (_norm(item["brand"]), _norm(item["category"]))
        if not any(key):
            continue
        old = merged.setdefault(key, item)
        if old is not item and item["source"] not in old["source"]:
            old["source"] += "；" + item["source"]
    result = list(merged.values())
    # 同一品类已经有带品牌的条目时，品牌为空的同品类条目只是同一行被正文和附件
    # 各抽了一次的重复证据（实测 深圳市南幽科技：附件第9行给出 ZCGJERTT，正文又
    # 给了一条空品牌，界面上就成了两条「设备电池」）。去掉空的那条，品牌留下。
    branded_categories = {
        _norm(item.get("category")) for item in result
        if _clean(item.get("brand")) and _clean(item.get("category"))
    }
    if branded_categories:
        result = [
            item for item in result
            if _clean(item.get("brand"))
            or _norm(item.get("category")) not in branded_categories
        ]
    # 已经有真正的电池品类时，「空品牌 + 品类就写着类别名（设备电池）」只是
    # 未找到明确品类时的填充行，留着会在界面上多出一条空品牌条目
    # （实测 香港郡為/香港酷康：3 条真实电池 + 1 条 空品牌/设备电池）。
    class_keys = {_norm(name) for name in BATTERY_CLASSES}
    has_real_category = any(
        _clean(item.get("category")) and _norm(item.get("category")) not in class_keys
        for item in result
    )
    if has_real_category:
        result = [
            item for item in result
            if _clean(item.get("brand"))
            or not _clean(item.get("category"))
            or _norm(item.get("category")) not in class_keys
        ]
    # 同品牌、同电池类别下「便携式电池 内置锂电池」与「便携式电池」是同一条证据的
    # 两种写法（实测 ZENITHFLOW 同品牌两条），保留信息更长的那条。
    deduped = []
    for item in result:
        merged_into = False
        for kept in deduped:
            if (_norm(kept.get("brand")) != _norm(item.get("brand"))
                    or (kept.get("category_class") or "") != (item.get("category_class") or "")):
                continue
            left = _norm(kept.get("category"))
            right = _norm(item.get("category"))
            if not left or not right:
                continue
            if left in right or right in left:
                if len(right) > len(left):
                    kept["category"] = item.get("category")
                merged_into = True
                break
        if not merged_into:
            deduped.append(item)
    result = deduped
    for category_class, filename in contract_classes.items():
        if not any(item.get("category_class") == category_class for item in result):
            result.append({"brand": "", "category": category_class, "category_class": category_class,
                           "source": "附件合同名称", "evidence": filename, "confidence": "medium"})
    if attachments and not contract_classes:
        for item in result:
            if not item.get("category_class"):
                item["category_class"] = "设备电池"
                item["confidence"] = "low"
                item["source"] += "；未见专用类别合同"
    if not result and attachments:
        # Absence of a dedicated contract is only the user's weak equipment-battery heuristic.
        result.append({"brand": "", "category": "", "category_class": "设备电池",
                       "source": "未见专用类别合同", "evidence": "仅作设备电池候选，待人工核对", "confidence": "low"})
    if not result:
        result.append({"brand": "", "category": "", "category_class": "",
                       "source": "邮件/附件未找到明确品类", "evidence": "请人工核对电池类别", "confidence": "low"})
    if form_brands:
        # 申请表品牌字段优先回填到「品牌还空着」的条目，而不是新增重复行：
        # 用户看到的是同一条业务原本就有类别候选，只是品牌缺失。
        empty_items = [item for item in result if not _clean(item.get("brand"))]
        for order, (brand, evidence) in enumerate(form_brands):
            if any(_norm(item.get("brand")) == _norm(brand) for item in result):
                continue
            if order < len(empty_items):
                item = empty_items[order]
                item["brand"] = brand
                item["source"] = "；".join(part for part in (item.get("source"), "申请表品牌字段") if part)
                if evidence:
                    item["evidence"] = evidence
                if item.get("confidence") == "low":
                    item["confidence"] = "medium"
            else:
                result.append({"brand": brand, "category": "", "category_class": "",
                               "source": "申请表品牌字段", "evidence": evidence,
                               "confidence": "medium"})
    for index, item in enumerate(result):
        item["item_id"] = "BAT-" + hashlib.sha256(repr((company, index, item["brand"], item["category"], item["source"])).encode("utf-8")).hexdigest()[:20]
        item["battery_confirmed"] = False
    return {"enabled": True, "items": result, "status": "pending"}


def is_german_battery(program):
    return re.sub(r"\s+", "", str(program or "")) in {"德国电池", "德国电池法"}


def _is_blank_battery_item(item):
    """品牌、品类都没填的占位壳。"""
    return (isinstance(item, dict)
            and not _clean(item.get("brand"))
            and not _clean(item.get("category")))


def battery_review(program, saved, source_items, status):
    items = saved.get("battery_items", source_items)
    # 早期版本的品牌抽取把「logo（必填）…」「单个电池重量」「对应的品牌logo（必填）」
    # 当成品牌存进了人工快照，这些条目会一直跟着这封邮件显示。品牌本身就是
    # 表头/说明文碎片时，整条都是机器误抽的产物，直接丢掉；人工填的真品牌
    # 不会命中这些特征。
    if isinstance(items, list) and items:
        kept = [
            item for item in items
            if not (isinstance(item, dict) and _clean(item.get("brand"))
                    and _is_junk_brand(item.get("brand")))
        ]
        if len(kept) != len(items):
            items = kept if kept else (source_items or [])
    # 快照里可能存着一条「品牌/品类全空」的占位壳。阶段一后来从申请表补出了
    # 品牌，空壳会把它整条盖掉 —— 表现就是「附件里明明写了品牌，界面是空的」。
    # 只有「全部是空壳 + 都未确认 + 新证据确实有内容」三个条件同时成立才切换，
    # 不覆盖任何人工确认结果。
    if (isinstance(items, list) and items
            and all(_is_blank_battery_item(item) for item in items)
            and not any(item.get("battery_confirmed") for item in items if isinstance(item, dict))
            and any(not _is_blank_battery_item(item) for item in (source_items or []))):
        items = source_items
    items = items if isinstance(items, list) else []
    required = "battery_items" in saved or bool(items)
    confirmed = bool(items) and all(item.get("battery_confirmed") for item in items)
    if not required:
        confirmed = status == "confirmed"  # Preserve previously confirmed business.
    if status in {"review", "returned", "needs_info"}:
        confirmed = False
    return {"enabled": is_german_battery(program), "items": items,
            "requires_review": required, "confirmed": confirmed,
            "status": "confirmed" if confirmed else "pending"}


def normalize_battery_items(raw_items, previous_items=(), confirm_id="", delete_id="", confirm_all=False):
    if not isinstance(raw_items, list) or len(raw_items) > 200:
        raise ValueError("电池品牌/品类格式不正确，一次最多200项")
    previous = {item.get("item_id"): item for item in previous_items if isinstance(item, dict)}
    result, seen = [], set()
    keys = ("brand", "category", "category_class")
    for index, raw in enumerate(raw_items):
        if not isinstance(raw, dict):
            raise ValueError("电池品牌/品类项目格式不正确")
        item = {key: str(raw.get(key) or "").strip() for key in keys}
        if any(len(value) > 1000 for value in item.values()):
            raise ValueError("电池品牌/品类字段过长")
        item_id = str(raw.get("item_id") or "BAT-" + hashlib.sha256(
            repr((index, item)).encode("utf-8")).hexdigest()[:20])
        if item_id in seen:
            raise ValueError("电池品牌/品类项目编号重复")
        seen.add(item_id)
        old = previous.get(item_id, {})
        # category is retained as original mail/attachment evidence, not a target
        # field that the reviewer must manually edit. Confirmation is brand + class.
        complete = bool(item["brand"] and item["category_class"])
        # Only an explicit confirmation or unchanged saved confirmation is trusted.
        # confirm_all 用于“整条电池明细一次确认”：与 WEEE 的整体确认对称，
        # 只有 brand + category_class 都齐全的项才被置为已确认。
        item["battery_confirmed"] = bool(complete and (confirm_all or item_id == confirm_id or (
            old.get("battery_confirmed") and all(old.get(key) == item[key] for key in keys))))
        item["item_id"] = item_id
        for key in ("source", "evidence", "confidence"):
            value = str(raw.get(key) or old.get(key) or "").strip()
            if value:
                item[key] = value[:1000]
        if item_id != delete_id:
            result.append(item)
    if confirm_id and confirm_id not in seen or delete_id and delete_id not in seen:
        raise ValueError("电池品牌/品类项目不存在")
    return result


def compare_battery_categories(items, workorders):
    """Compare each brand + battery type against the same category-detail row."""
    if isinstance(items, str):
        try:
            items = json.loads(items)
        except (TypeError, ValueError):
            items = []
    mail_items = [item for item in (items or []) if isinstance(item, dict)]
    orders = [order for order in (workorders or []) if isinstance(order, dict)]
    if not mail_items:
        return {"enabled": False, "status": "pending", "reason": "邮件未提取到明确品牌/电池类型",
                "items": [], "extra_items": [], "workorder_count": len(orders)}
    if not orders:
        checked = [{**item, "status": "missing", "reason": "实时工单列表未找到对应工单"}
                   for item in mail_items]
        return {"enabled": True, "status": "missing", "reason": "实时工单列表未找到对应工单",
                "items": checked, "extra_items": [], "workorder_count": 0}
    if len(orders) != 1:
        checked = [{**item, "status": "pending", "reason": "工单候选不唯一，无法确认品类明细归属"}
                   for item in mail_items]
        return {"enabled": True, "status": "pending", "reason": "工单候选不唯一，需人工复核",
                "items": checked, "extra_items": [], "workorder_count": len(orders)}

    order = orders[0]
    detail_status = str(order.get("德国电池法品类明细读取状态") or order.get("品类明细读取状态") or "").strip()
    category_rows = order.get("品类明细") or []
    if not isinstance(category_rows, list):
        category_rows = []
    category_rows = [row for row in category_rows if isinstance(row, dict)]
    if detail_status in {"failed", "pending", "not_read"}:
        checked = [{**item, "status": "pending", "reason": "工单品类明细读取失败或不完整"}
                   for item in mail_items]
        return {"enabled": True, "status": "pending", "reason": "工单品类明细读取失败或不完整，需人工复核",
                "items": checked, "extra_items": [], "workorder_count": 1}
    if detail_status == "loaded_empty":
        checked = [{**item, "status": "missing", "reason": "工单品类明细已确认为空"}
                   for item in mail_items]
        return {"enabled": True, "status": "missing", "reason": "工单品类明细为空，邮件目标电池类型缺失",
                "items": checked, "extra_items": [], "workorder_count": 1}
    if not category_rows:
        checked = [{**item, "status": "pending", "reason": "未确认工单品类明细是否完整"}
                   for item in mail_items]
        return {"enabled": True, "status": "pending", "reason": "未确认工单品类明细是否完整，需人工复核",
                "items": checked, "extra_items": [], "workorder_count": 1}

    brand_words = ("品牌", "brand", "marke")
    category_words = ("品类", "类别", "分类", "电池类型", "category", "battery type", "batteriekategorie")
    def has_field(row, words):
        return any(any(word in str(key).casefold() for word in words) for key in row)
    if not any(has_field(row, brand_words) for row in category_rows) or not any(
            has_field(row, category_words) for row in category_rows):
        checked = [{**item, "status": "pending", "reason": "工单品类明细缺少品牌或电池类型字段"}
                   for item in mail_items]
        return {"enabled": True, "status": "pending", "reason": "工单品类明细字段不清楚，需人工复核",
                "items": checked, "extra_items": [], "workorder_count": 1}

    pairs, order_categories, order_brands = [], [], []
    for row_index, row in enumerate(category_rows):
        brands, categories = [], []
        for key, value in row.items():
            label = str(key).casefold()
            if any(word in label for word in brand_words):
                found = _clean(value)
                if found:
                    brands.append(found)
            if any(word in label for word in category_words):
                found = _clean(value)
                if found:
                    categories.append(found)
        order_brands.extend(brands)
        order_categories.extend(categories)
        for brand in brands:
            for category in categories:
                pairs.append({"brand": brand, "category": category, "row_index": row_index,
                              "category_class": classify_battery_category(category)})

    checked, consumed = [], set()
    for item in mail_items:
        brand = _clean(item.get("brand"))
        category = _clean(item.get("category"))
        category_class = _clean(item.get("category_class")) or classify_battery_category(category)
        if not brand or category_class not in BATTERY_CLASSES:
            checked.append({**item, "status": "pending", "reason": "邮件未提取到完整品牌＋电池类型"})
            continue
        match_index = None
        for index, pair in enumerate(pairs):
            if index in consumed:
                continue
            brand_norm = _norm(brand)
            registered_brand = _norm(pair["brand"])
            brand_hit = bool(brand_norm and registered_brand and brand_norm == registered_brand)
            category_hit = pair["category_class"] == category_class
            if not category_hit and category:
                category_hit = _norm(category) == _norm(pair["category"])
            if brand_hit and category_hit:
                match_index = index
                break
        if match_index is None:
            checked.append({"brand": brand, "category": category, "category_class": category_class,
                            "status": "missing", "reason": "工单品类明细中未找到同一行对应的品牌＋电池类型"})
        else:
            consumed.add(match_index)
            checked.append({"brand": brand, "category": category, "category_class": category_class,
                            "status": "matched", "reason": "同一条工单品类明细中的品牌与电池类型均匹配"})
    extra_items = [{"brand": pair["brand"], "category": pair["category"],
                    "category_class": pair["category_class"]}
                   for index, pair in enumerate(pairs) if index not in consumed]
    if any(item["status"] == "missing" for item in checked):
        overall, reason = "missing", "至少一项邮件目标品牌＋电池类型缺失或不符"
    elif any(item["status"] == "pending" for item in checked):
        overall, reason = "pending", "邮件或工单电池字段不完整，需人工复核"
    else:
        overall, reason = "matched", "邮件目标品牌＋电池类型全部匹配"
        if extra_items:
            reason += f"；工单另有 {len(extra_items)} 项额外电池类型"
    return {"enabled": True, "status": overall, "reason": reason, "items": checked,
            "extra_items": extra_items, "workorder_categories": list(dict.fromkeys(order_categories)),
            "workorder_brands": list(dict.fromkeys(order_brands)), "workorder_count": 1}
