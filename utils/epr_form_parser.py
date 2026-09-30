# -*- coding: utf-8 -*-
"""EPR 申请表复选框解析器

背景
----
"泛欧 EPR 申请表" 系列模板(xlsx)中, 客户勾选的国家与业务**不在单元格里**,
而是存放在 xlsx 的**表单控件**(Form Control / CheckBox)中 —— openpyxl 完全读不到,
必须直接解 zip 读原始 XML:

  xl/worksheets/sheetN.xml          <controls><control name="Check Box 39" r:id="rId9">
                                        <from><xdr:col>1</xdr:col><xdr:row>5</xdr:row></from>
  xl/worksheets/_rels/sheetN.xml.rels   rId9 -> ../ctrlProps/ctrlPropX.xml
  xl/ctrlProps/ctrlPropX.xml         checked="Checked"  → 已勾选
  xl/drawings/drawingN.xml           shape name="Check Box 39" 内含 <a:t>西班牙（ES）</a:t>

即: 控件名(Check Box 39) ↔ 位置(行/列) 在 sheetN.xml; 勾选状态在 ctrlProps;
    控件名 → 显示文本(国家名) 在 drawingN.xml。

模板版本差异(均已实测)
----------------------
  旧版 泛欧EPR申请表-该国XX法.xlsx : 国家列 B, 业务列 C=WEEE / D=电池法 / E=包装法
  新版 泛欧EPR申请表-8国新版.xlsx  : 国家列 B, 业务列 D=WEEE / E=电池法 / F=包装法
→ 业务列**不写死**, 按表头文本("WEEE"/"电池法"/"包装法")自动定位。

用法
----
    from utils.epr_form_parser import parse_epr_form
    form = parse_epr_form("xx.xlsx")     # 不是可勾选申请表时返回 None
    form["projects"]              # ['西班牙包装法', '荷兰电池法']
    form["unmatched_countries"]   # 国家勾了但业务一个都没勾 → 需人工
    form["orphan_business"]       # 业务勾了但国家没勾 → 需人工
"""
import html
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional

_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

# 国家勾选框的显示文本形如 " 西班牙（ES）" / " 意大利 (IT)"
_COUNTRY_LABEL_RE = re.compile(r"([\u4e00-\u9fa5]{2,6})\s*[（(]\s*([A-Za-z]{2})\s*[)）]")

# 表头文本 → 标准业务名。匹配前会去掉所有空白字符。
_BT_MATCHERS = (
    ("WEEE", lambda t: t.upper() == "WEEE"),
    ("电池法", lambda t: t.startswith("电池法") or t.lower() == "battery"),
    ("包装法", lambda t: t.startswith("包装法") or t.lower() == "packaging"),
    ("一次性塑料", lambda t: t.startswith("一次性塑料") or t.startswith("塑料法")),
)

_HEADER_MAX_LEN = 24      # 表头文本长度上限, 用于排除说明文字
_SCAN_MAX_ROW = 30
_SCAN_MAX_COL = 20

# 单元格: 兼容自闭合 <c .../> 与 <c ...>...</c>
_CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
_R_RE = re.compile(r'\br="([A-Z]+)(\d+)"')


# --------------------------------------------------------------------------- #
# 公共入口
# --------------------------------------------------------------------------- #

def parse_epr_form(filepath: str) -> Optional[Dict]:
    """解析 EPR 申请表的勾选状态。

    返回 None 表示「不是可勾选的 EPR 申请表」(例如德国/意大利旧版无控件模板),
    调用方应回落到文本规则提取。
    """
    if not filepath or not os.path.exists(filepath):
        return None
    try:
        with zipfile.ZipFile(filepath, "r") as zf:
            return parse_epr_form_from_zip(zf)
    except zipfile.BadZipFile:
        return None
    except Exception:
        return None


def parse_epr_form_from_zip(zf: zipfile.ZipFile) -> Optional[Dict]:
    """同 parse_epr_form, 但接收已打开的 ZipFile(便于测试 / 内嵌 xlsx)"""
    names = set(zf.namelist())
    sheet_names = sorted(
        n for n in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)
    )
    for sheet_name in sheet_names:
        sheet_xml = _read(zf, sheet_name)
        if "<controls" not in sheet_xml:
            continue
        result = _parse_sheet(zf, names, sheet_name, sheet_xml)
        if result:
            return result
    return None


# --------------------------------------------------------------------------- #
# 单 sheet 解析
# --------------------------------------------------------------------------- #

def _parse_sheet(zf, names: set, sheet_name: str, sheet_xml: str) -> Optional[Dict]:
    cells = _cell_texts(zf, names, sheet_xml)
    bt_cols, header_row = _find_business_cols(cells)
    if not bt_cols:
        # 有控件但找不到 WEEE 表头 → 不是泛欧 EPR 申请表
        return None

    controls = _parse_controls(zf, names, sheet_name, sheet_xml)
    labels = _drawing_labels(zf, names)

    countries_checked: Dict[int, tuple] = {}   # row -> (国家名, 代码)
    biz_checked: Dict[int, List[str]] = {}     # row -> [业务名]

    for ctl in controls:
        if not ctl["checked"]:
            continue
        row, col = ctl["row"], ctl["col"]
        if row is None or col is None:
            continue
        if col in bt_cols:
            biz_checked.setdefault(row, []).append(bt_cols[col])
            continue
        country = _match_country_label(labels.get(ctl["name"], ""))
        if not country:
            # 兜底: 模板把国家名写成了单元格文本(新版 C 列)
            country = _country_from_row_cells(cells, row, min(bt_cols))
        if country:
            countries_checked[row] = country

    # 未勾选国家行也尝试登记? 不需要 — 只关心勾中的
    projects: List[str] = []
    unmatched: List[str] = []
    orphan: List[str] = []

    for row in sorted(biz_checked):
        country = countries_checked.get(row)
        for bt in biz_checked[row]:
            if country:
                name = f"{country[0]}{bt}"
                if name not in projects:
                    projects.append(name)
            else:
                item = f"{bt}(第{row}行)"
                if item not in orphan:
                    orphan.append(item)

    for row in sorted(countries_checked):
        if not biz_checked.get(row):
            name = countries_checked[row][0]
            if name not in unmatched:
                unmatched.append(name)

    return {
        "is_epr_form": True,
        "sheet": sheet_name,
        "header_row": header_row,
        # {业务名: 列号(0基)}
        "business_cols": {bt: col for col, bt in bt_cols.items()},
        "checked_countries": [countries_checked[r][0] for r in sorted(countries_checked)],
        "projects": projects,
        "unmatched_countries": unmatched,
        "orphan_business": orphan,
        "checked_controls": len(projects) + len(unmatched) + len(orphan),
    }


def _parse_controls(zf, names: set, sheet_name: str, sheet_xml: str) -> List[Dict]:
    """解析 <controls>: 控件名 / 位置 / 勾选状态"""
    rels_name = sheet_name.replace("worksheets/", "worksheets/_rels/") + ".rels"
    rels = _read(zf, rels_name)
    rid2prop = {}
    for tag in re.findall(r"<Relationship\b[^>]*/>", rels):
        if "ctrlProp" not in tag:
            continue
        rid = re.search(r'Id="([^"]+)"', tag)
        tgt = re.search(r'Target="([^"]+)"', tag)
        if rid and tgt:
            rid2prop[rid.group(1)] = _resolve(sheet_name, tgt.group(1))

    try:
        root = ET.fromstring(sheet_xml)
    except ET.ParseError:
        return []

    controls = []
    for el in root.iter():
        if _local(el.tag) != "control":
            continue
        # 注意: <xdr:from> 里的 col/row 是 **0 基**, 而单元格坐标是 1 基。
        # 这里统一 +1 归一化, 后续可安全与 _cell_texts 的 (row, col) 比较。
        row = col = None
        for sub in el.iter():
            if _local(sub.tag) != "from":
                continue
            for ch in sub:
                tag = _local(ch.tag)
                if tag == "col" and ch.text:
                    col = int(ch.text) + 1
                elif tag == "row" and ch.text:
                    row = int(ch.text) + 1
            break

        checked = False
        prop_path = rid2prop.get(el.get(_R_ID) or "")
        if prop_path and prop_path in names:
            checked = 'checked="Checked"' in _read(zf, prop_path)

        controls.append({
            "name": el.get("name") or "",
            "row": row,
            "col": col,
            "checked": checked,
        })
    return controls


def _drawing_labels(zf, names: set) -> Dict[str, str]:
    """drawingN.xml 里 shape 名 → 首个文本(即国家名标签)"""
    labels: Dict[str, str] = {}
    for name in sorted(n for n in names if re.fullmatch(r"xl/drawings/drawing\d+\.xml", n)):
        xml = _read(zf, name)
        for block in re.split(r"(?=<xdr:(?:sp|pic|graphicFrame)\b)", xml):
            text = re.search(r"<a:t>([^<]*)</a:t>", block)
            if not text:
                continue
            sname = re.search(r'<xdr:cNvPr[^>]*name="([^"]*)"', block) or \
                re.search(r'name="([^"]*)"', block)
            if sname:
                labels.setdefault(sname.group(1), text.group(1).strip())
    return labels


# --------------------------------------------------------------------------- #
# 单元格文本 / 表头定位
# --------------------------------------------------------------------------- #

def _cell_texts(zf, names: set, sheet_xml: str) -> Dict[tuple, str]:
    """(row, col) -> 文本(已去掉空白)。row/col 均为 1 基(row 为 Excel 行号)。"""
    shared = _shared_strings(zf, names)
    out: Dict[tuple, str] = {}
    # 注意: 必须同时兼容 <c .../> 自闭合与 <c ...>...</c> 两种写法。
    # 若只用 `<c r="X">(.*?)</c>`, 自闭合单元格会把后续若干单元格一并吞掉,
    # 造成表头丢失(实测 泛欧EPR申请表-荷兰电池法.xlsx 的 WEEE 单元格即因此丢失)。
    for m in _CELL_RE.finditer(sheet_xml):
        attrs = m.group(1) or ""
        inner = m.group(2) or ""
        rm = _R_RE.search(attrs)
        if not rm:
            continue
        row = int(rm.group(2))
        col = _col_index(rm.group(1))
        if row > _SCAN_MAX_ROW or col > _SCAN_MAX_COL:
            continue
        text = ""
        if re.search(r'\bt="s"', attrs):
            v = re.search(r"<v>([^<]*)</v>", inner)
            if v:
                try:
                    text = shared[int(v.group(1))]
                except (ValueError, IndexError):
                    text = ""
        elif "<is>" in inner:
            text = re.sub(r"<[^>]+>", "", inner)
        else:
            v = re.search(r"<v>([^<]*)</v>", inner)
            if v:
                text = v.group(1)
        text = re.sub(r"\s+", "", html.unescape(text))
        if text:
            out[(row, col)] = text
    return out


def _shared_strings(zf, names: set) -> List[str]:
    if "xl/sharedStrings.xml" not in names:
        return []
    xml = _read(zf, "xl/sharedStrings.xml")
    return [
        re.sub(r"\s+", "", html.unescape(re.sub(r"<[^>]+>", "", si)))
        for si in re.findall(r"<si>(.*?)</si>", xml, re.S)
    ]


def _find_business_cols(cells: Dict[tuple, str]) -> tuple:
    """扫表头行定位业务列: 返回 ({列号: 业务名}, 表头行号)"""
    weee_pos = None
    for (row, col), text in sorted(cells.items()):
        if text.upper() == "WEEE" and len(text) <= 12:
            if weee_pos is None or row < weee_pos[0]:
                weee_pos = (row, col)
    if not weee_pos:
        return {}, None

    header_row, weee_col = weee_pos
    bt_cols: Dict[int, str] = {}
    for (row, col), text in cells.items():
        if row != header_row or col < weee_col or len(text) > _HEADER_MAX_LEN:
            continue
        for bt, matcher in _BT_MATCHERS:
            if matcher(text):
                bt_cols[col] = bt
                break
    return bt_cols, header_row


def _match_country_label(label: str) -> Optional[tuple]:
    """' 西班牙（ES）' → ('西班牙', 'ES')"""
    if not label:
        return None
    m = _COUNTRY_LABEL_RE.search(label)
    if not m:
        return None
    return (m.group(1), m.group(2).upper())


def _country_from_row_cells(cells: Dict[tuple, str], row: int, bt_min_col: int) -> Optional[tuple]:
    """兜底: 国家名写在业务列左侧的单元格里"""
    for col in range(1, bt_min_col):
        text = cells.get((row, col), "")
        if text and _COUNTRY_LABEL_RE.search(text):
            return _match_country_label(text)
        if text and any(text == c for c in _DRAWING_FALLBACK_COUNTRIES):
            return (text, "")
    return None


# 仅作兜底时用: 单元格里出现的裸国家名
_DRAWING_FALLBACK_COUNTRIES = (
    "德国", "荷兰", "爱尔兰", "意大利", "比利时", "波兰", "丹麦", "法国", "捷克",
    "葡萄牙", "瑞典", "西班牙", "卢森堡", "奥地利", "匈牙利", "芬兰", "罗马尼亚",
    "挪威", "爱沙尼亚", "瑞士", "希腊", "英国", "加拿大", "拉脱维亚",
)


# --------------------------------------------------------------------------- #
# 工具
# --------------------------------------------------------------------------- #

def _read(zf, name: str) -> str:
    try:
        return zf.read(name).decode("utf-8", "ignore")
    except (KeyError, ValueError):
        return ""


def _resolve(sheet_name: str, target: str) -> str:
    """把 rels 里的相对 Target 解析为 zip 内绝对路径"""
    if target.startswith("/"):
        return target.lstrip("/")
    base = os.path.dirname(sheet_name)
    return os.path.normpath(os.path.join(base, target)).replace("\\", "/")


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _col_index(letters: str) -> int:
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx
