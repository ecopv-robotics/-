"""M5 工单系统比对模块 — Playwright 半自动登录（验证码人工输入），查询并比对

操作流程:
  1. 登录（自动填账号密码，人工输验证码）
  2. 进入注册工单页面
  3. 逐条数据查询:
     a. 填入公司名称（文本输入）
     b. 选择所属代理（下拉框）
     c. 选择国家（下拉框）
     d. 选择服务项目（下拉框）
     e. 选择状态（下拉框）
     f. 按邮件发件日期前后一个日历月填写下单日期
     g. 点击查询
     h. 无结果 → 模糊匹配公司名称重新查询
     i. 读取表格结果
  4. 四维模糊比对；多条候选按邮件日期最近原则筛选
"""
import asyncio
import calendar
import difflib
import json
import os
import re
import time
from datetime import date as date_type, datetime, timedelta
from typing import Callable, List, Dict, Optional, Tuple
from urllib.parse import urljoin

from utils.fuzzy_match import fuzzy_match_pair, normalize_text, to_simplified
from modules.weee_category_audit import compare_weee_categories


REGISTRATION_ORDER_PATH = "/work/workOrderRegister"

# 邮件中的代理名称偶尔会出现同音/形近字，而工单平台只接受下拉框中的
# 正式名称。这个映射只用于网页查询，不改写阶段一输出中的原始字段，
# 便于人工追溯数据来源；配置中的 agent_aliases 可以覆盖或扩展它。
DEFAULT_AGENT_ALIASES = {
    "示例代理乙别名": "示例代理乙",
    # 邮件里偶尔写代理公司全称，而工单平台下拉只认简称。
    "示例f1538bdc有限公司": "示例代理甲",
    "示例8422eccf有限公司": "示例代理甲",
}


# 结果表“抓对了”的最小证据：至少要命中一个业务列名。列名完全对不上说明抓到
# 了页面上别的表格，此时必须报错，而不是产出 0 条假漏单。可用
# config.result_table_columns 覆盖或扩展。
DEFAULT_RESULT_COLUMNS = {
    "工单编号", "订单编号", "客户", "客户名称", "公司", "公司名称", "客户号",
    "所属代理", "代理", "代理名称", "服务项目", "项目", "项目名称", "国家",
    "状态", "工单状态", "下单日期", "创建时间", "提交时间", "客户信息",
    # 工单平台（Ant Design 表格）的实际列名，含合并列
    "主工单编号/客户号", "所属代理/直客", "注册类型", "证书号", "服务商",
    "最近交付进度", "驳回原因",
}


# 工单平台结果表的实际列名与内部约定字段名不一致（合并列尤其明显）。这里做一次
# 别名归并：保留原始列名，同时补一个约定字段名。否则下游按「所属代理」取值时会
# 静默拿到空字符串，代理维度形同失效——不报错、但比对结果不可信。
RESULT_COLUMN_ALIASES = {
    "所属代理/直客": "所属代理",
    "代理/直客": "所属代理",
    "主工单编号/客户号": "工单编号",
    "工单编号/客户号": "工单编号",
    "服务项目": "项目",
    "标准化项目": "项目",
    "公司名称": "客户",
    "状态": "工单状态",
    "下单日期": "日期",
    "创建时间": "提交时间",
}


# 默认选择器（config.yaml 的 workorder.selectors 段可按键覆盖，无需改代码）
DEFAULT_SELECTORS = {
    "login_account": [
        'input[placeholder="账号"]',
        'input[name="username"]',
        'input[name="account"]',
        'input[type="text"]',
        'input:first-of-type',
    ],
    "login_password": [
        'input[type="password"]',
        'input[placeholder="密码"]',
        'input[name="password"]',
    ],
    # 工单平台的菜单是二级结构：先展开“工单管理”，再点“注册工单”。
    # 禁止把宽泛的“工单”或“注册”文字当成可安全点击的入口。
    "workorder_management": [
        'div:has-text("工单管理") >> nth=0',
        'li:has-text("工单管理") >> nth=0',
    ],
    "dropdown_option": [
        ".el-select-dropdown__item",
        ".ant-select-item",
        ".ant-select-option",
        "li[role='option']",
        "div[role='option']",
        ".dropdown-item",
        ".select-option",
    ],
    "company_input": [
        'input[placeholder="工单编号/客户号/公司名称"]',
        'input[placeholder*="公司"]',
        'input[placeholder*="客户"]',
        'input[placeholder*="名称"]',
        'input[placeholder*="请输入"]',
        'input[name*="company"]',
        'input[name*="customer"]',
        'input[name*="name"]',
    ],
    # 注册工单查询页的四个实际控件。必须优先使用这些精确占位符，
    # 避免把“注册类型”等其他下拉框当成代理/国家/服务项目。
    "agent_input": [
        '.ant-select:has-text("选择所属代理")',
        '.el-select:has-text("选择所属代理")',
        # 查询栏固定顺序：注册类型(0) → 所属代理(1)。
        # 选中值后占位文字会消失，因此保留位置兜底。
        '.ant-select >> nth=1',
        '.el-select >> nth=1',
        'input[placeholder="选择所属代理"]',
        'input[placeholder*="所属代理"]',
        # Ant Design Select 将占位文本渲染到 span，内部搜索 input
        # 通常没有 placeholder 属性。
        '.ant-select:has(.ant-select-selection-placeholder:has-text("选择所属代理"))',
    ],
    "country_input": [
        '.ant-select:has-text("国家")',
        '.el-select:has-text("国家")',
        '.ant-select >> nth=2',
        '.el-select >> nth=2',
        'input[placeholder="国家"]',
        'input[placeholder*="国家"]',
        '.ant-select:has(.ant-select-selection-placeholder:has-text("国家"))',
    ],
    "service_item_input": [
        '.ant-select:has-text("服务项目")',
        '.el-select:has-text("服务项目")',
        '.ant-select >> nth=3',
        '.el-select >> nth=3',
        'input[placeholder="服务项目"]',
        'input[placeholder*="服务项目"]',
        '.ant-select:has(.ant-select-selection-placeholder:has-text("服务项目"))',
    ],
    # 页面查询栏中的下单日期控件。日期筛选是按每封邮件的发件日期动态计算的，
    # 不能使用固定的全局日期，否则同一批邮件跨月份时会互相污染。
    "order_date_from_input": [
        'input[placeholder="下单开始日期"]',
        'input[placeholder*="下单开始"]',
        'input[aria-label*="下单开始"]',
        'input[placeholder*="开始日期"]',
    ],
    "order_date_to_input": [
        'input[placeholder="下单结束日期"]',
        'input[placeholder*="下单结束"]',
        'input[aria-label*="下单结束"]',
        'input[placeholder*="结束日期"]',
    ],
    "query_button": [
        'button:has-text("查询")',
        'button:has-text("搜索")',
        'button:has-text("查找")',
        'button:has-text("检索")',
        'button[type="submit"]',
        'button.el-button--primary:has-text("查")',
        'button.ant-btn-primary:has-text("查")',
    ],
    # 仅作为“页面上可能存在结果表格”的参考列表；结果解析已改为按结构读取
    # （见 TABLE_EXTRACT_JS），不再按这个列表逐个试选择器。
    "table": [
        ".el-table__body-wrapper table",
        ".el-table table",
        ".ant-table-content table",
        "table.table",
        "table.data-table",
        "table",
    ],
}


# 结果表结构读取脚本（在浏览器里一次性读成纯数据，避免逐单元格往返）。
#
# 表头/表体分离是硬约束：Element UI 与 Ant Design 都把表头和表体渲染成两张独立
# 的 <table>，**表体那张表自身没有 thead**：
#     Element UI : .el-table__header-wrapper table / .el-table__body-wrapper table
#     Ant Design : .ant-table-header > table       / .ant-table-body > table
# 早期实现把“第一个 tr”当成表头，抓到的其实是表体的第一行数据，于是表头全是空
# 字符串、每行 key 全部塌缩成 ''，下游 match_records 取不到客户/项目/代理，整批
# 邮件被误判成“漏单”。所以这里必须把表头和表体彻底分开读，再按列索引映射。
#
# 另外必须把“结构占位行”排除掉，否则空结果的表格也会被读成 1 行：
#     ant-table-measure-row（0 高计量行）、ant-table-placeholder（暂无数据行）、
#     ant-table-expanded-row / ant-table-summary（展开行、汇总行）。
TABLE_EXTRACT_JS = r"""
() => {
  const isVisible = (el) => {
    if (!el) return false;
    const r = el.getBoundingClientRect();
    if (r.width <= 0 || r.height <= 0) return false;
    const s = getComputedStyle(el);
    return s.display !== 'none' && s.visibility !== 'hidden';
  };
  const cellText = (cell) => (cell.innerText || cell.textContent || '')
    .replace(/\s+/g, ' ').trim();
  // colSpan 会让列索引错位：跨列单元格后面补占位，保证下标仍与表头一一对应。
  const readCells = (row, selector) => {
    const out = [];
    for (const cell of Array.from(row.querySelectorAll(selector))) {
      const span = Math.max(1, parseInt(cell.getAttribute('colspan') || '1', 10) || 1);
      out.push(cellText(cell));
      for (let i = 1; i < span; i++) out.push('');
    }
    return out;
  };
  const readHeaders = (table) => {
    if (!table) return { headers: [], found: false };
    const headRows = Array.from(table.querySelectorAll('thead tr'))
      .filter(r => r.querySelector('th, td'));
    if (headRows.length) {
      const headers = readCells(headRows[headRows.length - 1], 'th, td');
      if (headers.some(t => t)) return { headers, found: true };
    }
    // 没有 thead 的普通表格：只用 th 当表头，绝不用第一行的 td 顶替。
    const first = table.querySelector('tr');
    if (first && first.querySelector('th')) {
      const headers = readCells(first, 'th');
      if (headers.some(t => t)) return { headers, found: true };
    }
    return { headers: [], found: false };
  };
  // 结构占位行（非业务数据行）
  const STRUCTURAL = [
    'el-table__expanded-cell',
    'ant-table-measure-row',
    'ant-table-placeholder',
    'ant-table-expanded-row',
    'ant-table-summary',
    'ant-table-row-expand-icon-cell',
  ];
  const isStructuralRow = (row) => {
    const cls = String(row.className || '');
    for (const k of STRUCTURAL) {
      if (cls.indexOf(k) >= 0) return true;
    }
    return !!row.closest(
      '.el-table__expanded-cell, .ant-table-expanded-row, .ant-table-placeholder'
    );
  };
  const readBody = (table) => {
    if (!table) return [];
    let rows = Array.from(table.querySelectorAll('tbody tr'));
    if (!rows.length) {
      rows = Array.from(table.querySelectorAll('tr')).filter(r => r.querySelector('td'));
    }
    const out = [];
    for (const row of rows) {
      if (isStructuralRow(row)) continue;
      if (!isVisible(row)) continue;
      const cells = readCells(row, 'td');
      if (!cells.length) continue;
      if (!cells.some(t => t)) continue;   // 全空占位行直接丢掉
      out.push(cells);
    }
    return out;
  };

  // 表头表 / 表体表配对。同一组只出一个候选，避免“表头表”和“表体表”被当成
  // 两个表格互相竞争——那样行数多的表体表会赢，结果就是表头为空。
  const pairOf = (table) => {
    const elRoot = table.closest('.el-table');
    if (elRoot) {
      const isHeader = !!table.closest('.el-table__header-wrapper');
      const isBody = !!table.closest('.el-table__body-wrapper');
      if (isHeader || isBody) {
        return {
          kind: 'el-table',
          root: elRoot,
          header: elRoot.querySelector('.el-table__header-wrapper table'),
          body: elRoot.querySelector('.el-table__body-wrapper table'),
          empty: !!elRoot.querySelector('.el-table__empty-block, .el-table__empty-text'),
        };
      }
    }
    const antBox = table.closest('.ant-table-container') || table.closest('.ant-table');
    if (antBox) {
      const content = antBox.querySelector('.ant-table-content > table');
      return {
        kind: 'ant-table',
        root: antBox,
        header: antBox.querySelector('.ant-table-header > table') || content || table,
        body: antBox.querySelector('.ant-table-body > table') || content || table,
        empty: !!antBox.querySelector('.ant-table-placeholder, .ant-empty'),
      };
    }
    return { kind: 'plain', root: table, header: table, body: table, empty: false };
  };

  const tables = Array.from(document.querySelectorAll('table')).filter(isVisible);
  const results = [];
  const done = new Set();
  for (const table of tables) {
    if (done.has(table)) continue;
    const pair = pairOf(table);
    if (pair.kind === 'plain') {
      const h = readHeaders(table);
      let rows = readBody(table);
      const empty = !!table.querySelector('.ant-table-placeholder, .ant-empty');
      // 单表结构（thead + tbody 在同一张表里，如 .ant-table-content）不该产生
      // “表头空但有行”的候选；真读不到表头就如实上报 found=false。
      results.push({
        headers: h.headers, headerFound: h.found, rows: rows,
        kind: 'plain', empty: empty,
      });
      continue;
    }
    if (done.has(pair.root)) continue;
    done.add(pair.root);
    if (pair.header) done.add(pair.header);
    if (pair.body) done.add(pair.body);
    const h = readHeaders(pair.header);
    let rows = readBody(pair.body);
    // 表体表恰好没有行时，表头表自己也带 tbody（少见），兜底读一次。
    if (!rows.length && pair.body !== pair.header) rows = readBody(pair.header);
    results.push({
      headers: h.headers, headerFound: h.found, rows: rows,
      kind: pair.kind, empty: pair.empty,
    });
  }
  return results;
}
"""


# 解析失败时抓现场用：把页面上所有表格 / iframe / 表头类元素的结构摘要成纯数据。
# 整页 HTML 动辄几百 KB，先看摘要通常就能定位“表头为什么读不到”。
TABLE_DEBUG_JS = r"""
() => {
  const desc = (t) => {
    const r = t.getBoundingClientRect();
    return {
      cls: String(t.className || ''),
      parentCls: t.parentElement ? String(t.parentElement.className || '') : '',
      inElTable: !!t.closest('.el-table'),
      inHeaderWrapper: !!t.closest('.el-table__header-wrapper'),
      inBodyWrapper: !!t.closest('.el-table__body-wrapper'),
      theadTr: t.querySelectorAll('thead tr').length,
      thCount: t.querySelectorAll('th').length,
      tbodyTr: t.querySelectorAll('tbody tr').length,
      allTr: t.querySelectorAll('tr').length,
      tdCount: t.querySelectorAll('td').length,
      visible: r.width > 0 && r.height > 0,
      text: (t.innerText || '').replace(/\s+/g, ' ').slice(0, 180),
    };
  };
  return {
    url: location.href,
    title: document.title,
    tableCount: document.querySelectorAll('table').length,
    tables: Array.from(document.querySelectorAll('table')).map(desc),
    iframes: Array.from(document.querySelectorAll('iframe')).map(f => ({
      src: f.src, id: f.id, cls: String(f.className || ''),
      visible: !!(f.offsetWidth || f.offsetHeight),
    })),
    elTable: document.querySelectorAll('.el-table').length,
    elHeaderWrapper: document.querySelectorAll('.el-table__header-wrapper').length,
    elBodyWrapper: document.querySelectorAll('.el-table__body-wrapper').length,
    elEmptyBlock: document.querySelectorAll('.el-table__empty-block').length,
    antTable: document.querySelectorAll('.ant-table').length,
    headerishEls: Array.from(document.querySelectorAll('[class*="header"]'))
      .slice(0, 30)
      .map(e => ({
        tag: e.tagName,
        cls: String(e.className || ''),
        text: (e.innerText || '').replace(/\s+/g, ' ').slice(0, 120),
      })),
  };
}
"""


class WorkOrderChecker:
    # 行政区划词：长的排前面，避免「高新技术开发区」被「开发区」或「区」提前切走。
    # 只用于剥离公司名**开头**的行政区划，不碰名字中间的部分。
    _ADMIN_DIVISION_WORDS = (
        "特别行政区", "自治区", "经济特区", "经济技术开发区",
        "高新技术产业开发区", "高新技术开发区", "经济开发区", "技术开发区",
        "高新技术产业园区", "保税港区", "保税区", "高新区", "开发区",
        "工业区", "产业园", "新区", "街道", "省", "市", "区", "县", "镇", "乡", "旗",
    )
    # 括号匹配。中英文括号都要覆盖，且允许括号内容为空。
    _BRACKET_RE = re.compile(r"[（(【\[][^）)】\]]*[）)】\]]")
    # 只有括号里装的是**主体类型说明**时，才把这一块整块删掉；装地名/国别的
    # 括号一律保留原样。白名单化是刻意的（用户 2026-09-14 定，选严格侧）：
    #   - 漏删：只是少一个候选词，模糊重查少查一轮 → 保守失败，安全；
    #   - 误删：「某某（中国）有限公司」→「某某有限公司」、「深圳某某（上海）
    #     贸易有限公司」→「深圳某某贸易有限公司」，会拼出一个平台上真实存在
    #     的**别家**公司名 → 假「已录单」→ 核对功能实质性失效。
    _BRACKET_TYPE_HINTS = (
        "个体", "独资", "合伙", "分支", "分公司", "一人", "小微",
        "农民专业", "外商投资", "中外合资", "中外合作", "台港澳",
        "自然人", "集团",
    )
    # 匹配「开头 2~10 字 + 一个行政区划词」，非贪婪所以取最短的地名单元。
    # 下限卡在 2 字是刻意的：真实地名前置词至少两字（深圳/上海/宝安/浦东），
    # 而「城市之家」「夜市」这类品牌词的前置词只有一字，用它可以把误伤挡掉。
    _ADMIN_PREFIX_RE = re.compile(
        r"^.{2,10}?(?:" + "|".join(map(re.escape, _ADMIN_DIVISION_WORDS)) + r")"
    )
    # 剥掉区划后剩余部分的下限：剩余过短就不剥，避免把主体名剃成两三个字，
    # 那种词搜出来基本全是无关公司。
    _MIN_STRIPPED_LEN = 4
    # 名字开头最多认几级行政区划（省/市/区县/乡镇）。真实公司名不会超过 3 级，
    # 卡上限是防呆：万一正则被人名/品牌名里的"市/区"一路带跑，别无限切下去。
    _MAX_ADMIN_UNITS = 3
    # 公司后缀（按长到短匹配，避免「有限责任公司」被「公司」提前截断）
    _COMPANY_SUFFIXES = (
        "有限责任公司", "股份有限公司", "有限公司", "股份公司", "集团", "公司",
    )

    def __init__(
        self,
        config: dict,
        logger=None,
        cancel_requested: Optional[Callable[[], bool]] = None,
    ):
        self.url = config["url"]
        self.username = config["username"]
        self.password = config["password"]
        self.timeout = config.get("timeout", 30)
        self.date_tolerance = config.get("date_tolerance_days", 60)
        # 工单查询按邮件发件日期前后一个日历月筛选；旧配置未提供时默认启用。
        self.date_window_months = max(0, int(config.get("date_window_months", 1)))
        self.customer_threshold = config.get("customer_threshold", 85)
        self.registration_order_path = config.get(
            "registration_order_path", REGISTRATION_ORDER_PATH
        )
        self.selectors = {**DEFAULT_SELECTORS, **(config.get("selectors") or {})}
        # 断点续查 / 风控
        self.query_interval = float(config.get("query_interval_seconds", 3))
        self.cache_path = config.get("query_cache_path", "storage/query_cache.json")
        self.cache_ttl_days = int(config.get("query_cache_ttl_days", 7))
        self.max_total_seconds = float(config.get("max_total_seconds", 0))  # 0=不限
        self.max_consecutive_query_failures = max(
            1, int(config.get("max_consecutive_query_failures", 3))
        )
        # 即使查询结果来自缓存，也默认在真实浏览器中重放一次查询条件，
        # 让人工能够看到公司、代理、国家、服务项目的填写过程。
        self.show_cached_inputs = bool(config.get("show_cached_inputs", True))
        self.input_step_delay = max(
            0.0, float(config.get("input_step_delay_seconds", 0.8))
        )
        self.cached_input_hold = max(
            0.0, float(config.get("cached_input_hold_seconds", 1.2))
        )
        # 无头运行：不弹浏览器窗口。
        # 登录页有图形验证码，无头下不可能人工输入，因此无头依赖**登录态复用**：
        # storage/login_state.json 里的 cookie 还有效就能直接进系统；失效时按
        # headless_fallback_to_headed 决定是否临时弹窗（详见 login()）。
        self.headless = bool(config.get("headless", False))
        # Windows 操作人员通常已有 Edge。优先使用系统 Edge 可以避免把约 800 MB
        # 的 Chromium/Headless Chromium 二进制塞进 PyInstaller 发布包；需要完整
        # Chromium 时仍可在 config.yaml 中改成 ``chromium``，旧的便携运行时也可回退。
        self.browser_channel = str(
            config.get("browser_channel")
            or os.environ.get("ECOPV_BROWSER_CHANNEL", "msedge")
        ).strip().lower()
        if self.browser_channel in {"default", "bundled", "none"}:
            self.browser_channel = ""
        self.headless_fallback = bool(
            config.get("headless_fallback_to_headed", True)
        )
        self.login_state_path = str(
            config.get("login_state_path")
            or os.path.join("storage", "login_state.json")
        )
        self.login_wait_seconds = max(
            10, int(float(config.get("login_wait_seconds", 120)))
        )
        # 长时间停留后再次点击查询时，先主动探测一次登录页并刷新保存态，
        # 避免把远端会话过期误报成“查询无数据/页面失效”。远端 Cookie 的实际
        # 有效期仍由工单系统决定；过期后会重新进入可视登录流程，而不是静默失败。
        self.login_refresh_interval_seconds = max(
            60, int(float(config.get("login_refresh_interval_seconds", 900)))
        )
        # 无头下没人看浏览器，凡是"给人看的"动作一律省掉：
        # 字段间停留归零、缓存命中的输入条件重放直接跳过（每条能省好几秒）。
        if self.headless:
            self.input_step_delay = 0.0
            self.cached_input_hold = 0.0
            self.show_cached_inputs = False
        self.field_timeout_ms = max(
            1000, int(float(config.get("field_timeout_seconds", 5)) * 1000)
        )
        self.dropdown_timeout_seconds = max(
            0.5, float(config.get("dropdown_timeout_seconds", 5))
        )
        # 点击查询后的沉降时间：只留很短一段，真正的"结果何时就绪"由
        # _wait_for_table_refresh 的指纹轮询判断。原先这里写死 sleep(2)，
        # 每次查询都白等 2 秒（模糊重查会连查 3~4 轮）。
        self.query_click_settle_seconds = max(
            0.0, float(config.get("query_click_settle_seconds", 0.3))
        )
        # 结果表刷新等待上限。原先复用 dropdown_timeout_seconds(5 秒)，
        # 遇到"本次结果与上次完全相同"（两次都命中暂无数据）时会死等满 5 秒。
        self.table_refresh_timeout_seconds = max(
            0.5, float(config.get("table_refresh_timeout_seconds", 3.0))
        )
        configured_aliases = config.get("agent_aliases") or {}
        self.agent_aliases = {
            normalize_text(str(k)): str(v).strip()
            for k, v in {**DEFAULT_AGENT_ALIASES, **configured_aliases}.items()
            if str(k).strip() and str(v).strip()
        }
        self._logged_agent_aliases = set()
        # 结果表列名白名单：用于识别“抓错表格”，避免把解析失败写成 0 条漏单。
        self.result_columns = {
            str(c).strip()
            for c in (config.get("result_table_columns") or DEFAULT_RESULT_COLUMNS)
            if str(c).strip()
        }
        # 结果表解析状态：ok / failed；failed 时该条不写缓存也不判漏单。
        self._last_table_parse_status = "ok"
        self._last_table_parse_note = ""
        # 解析失败时只抓一次现场快照（HTML + 截图 + 结构摘要），供离线定位表结构。
        self._result_page_dumped = False
        self.debug_dump_dir = (
            config.get("debug_dump_dir")
            or os.environ.get("MAIL_AUDIT_DEBUG_DIR")
            or os.path.join("output", "debug")
        )
        self.force_live_query = bool(config.get("force_live_query", False))
        self._session_start_ts = None
        self._last_login_check = 0.0
        self.logger = logger
        self._playwright = None
        self._browser = None
        self._context = None
        self._page = None
        self._logged_in = False
        # 是否为输验证码临时降级成了可视化浏览器（登录成功后据此决定要不要切回无头）
        self._headed_fallback_used = False
        self._headless_restore_failed = False
        # GUI 的“停止”来自另一个线程；回调只读取 thread-safe Event 状态，
        # RPA 不持有或调用任何 Qt 对象。
        self._cancel_requested = cancel_requested
        self.cancelled = False
        self._cancel_notice_logged = False
        self._last_query_status = ""
        # 本条查询的可审计来源。_last_query_status 继续保持兼容（实时查询/缓存复用），
        # 细分结果另存，避免把“模糊命中”误当成查询失败。
        self._last_lookup_mode = "未执行"
        self._last_fuzzy_term = ""
        self._last_date_window = None
        self._last_date_filter_status = ""
        self.last_processed_rows: List[Dict] = []
        # 加载缓存
        self._query_cache: Dict[str, dict] = self._load_cache()

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    def _is_cancel_requested(self) -> bool:
        """安全读取 GUI 工作线程传来的停止状态。"""
        if self._cancel_requested is None:
            return False
        try:
            requested = bool(self._cancel_requested())
        except Exception as exc:
            self._log(f"读取停止状态失败，继续本次查询: {exc}", "warning")
            return False
        if requested:
            self.cancelled = True
            if not self._cancel_notice_logged:
                self._log("收到停止请求，正在取消当前工单查询…", "warning")
                self._cancel_notice_logged = True
        return requested

    async def _await_or_cancel(self, awaitable, cancelled_value):
        """等待 Playwright 操作，同时每 0.1 秒检查一次用户停止请求。"""
        if self._is_cancel_requested():
            return cancelled_value

        task = asyncio.ensure_future(awaitable)
        try:
            while True:
                done, _ = await asyncio.wait({task}, timeout=0.1)
                if task in done:
                    return task.result()
                if self._is_cancel_requested():
                    task.cancel()
                    try:
                        await task
                    except asyncio.CancelledError:
                        pass
                    return cancelled_value
        except BaseException:
            if not task.done():
                task.cancel()
            raise

    async def _ensure_browser(self, headless: Optional[bool] = None):
        """启动浏览器。headless=None 时按配置；显式传值用于有头/无头切换。"""
        if self._page is not None:
            return
        # 便携版把 Playwright 浏览器放在程序目录；源码运行仍沿用系统默认目录。
        from utils.runtime_paths import APP_ROOT
        bundled_browsers = os.path.join(str(APP_ROOT), "ms-playwright")
        if os.path.isdir(bundled_browsers):
            os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", bundled_browsers)
        from playwright.async_api import async_playwright
        mode = self.headless if headless is None else bool(headless)
        self._playwright = await async_playwright().start()
        # 优先使用系统 Edge；Edge 的无头模式仍使用真实浏览器内核，且不需要
        # Playwright 自带的 Chromium 文件。若明确配置了 chromium，则保持原逻辑。
        launch_kwargs = {
            "headless": mode,
            "args": ["--disable-blink-features=AutomationControlled"],
        }
        if self.browser_channel:
            launch_kwargs["channel"] = self.browser_channel
        try:
            self._browser = await self._playwright.chromium.launch(**launch_kwargs)
        except Exception as exc:
            channel = launch_kwargs.pop("channel", "")
            # 带完整浏览器的旧包/源码环境仍保留回退能力；Edge-only 轻量包
            # 没有 bundled Chromium 时，会给出明确提示而不是静默失败。
            if channel and channel != "chromium":
                self._log(
                    f"系统浏览器 {channel} 启动失败({exc})，尝试 Playwright 内置浏览器",
                    "warning",
                )
                try:
                    self._browser = await self._playwright.chromium.launch(**launch_kwargs)
                except Exception as fallback_exc:
                    raise RuntimeError(
                        "无法启动工单浏览器：请确认操作人员电脑已安装 Microsoft Edge，"
                        "或在 config.yaml 的 workorder.browser_channel 中配置可用浏览器。"
                    ) from fallback_exc
            else:
                raise

        context_kwargs = {
            "viewport": {"width": 1440, "height": 900},
            "locale": "zh-CN",
            "timezone_id": "Asia/Shanghai",
        }
        # 无头内核的 UA 里带 "HeadlessChrome"，这是平台风控最容易一眼抓住的
        # 特征（实测默认就是 HeadlessChrome/151）。显式改回普通 Chrome 的写法，
        # 版本号取自真实内核，不写死，浏览器升级后自动跟上。
        if mode:
            version = getattr(self._browser, "version", "") or ""
            if version:
                context_kwargs["user_agent"] = (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    f"Chrome/{version} Safari/537.36"
                )
        # 复用上次保存的登录态：无头下没有人工输验证码的机会，全靠这份 cookie。
        state_loaded = False
        if os.path.exists(self.login_state_path):
            context_kwargs["storage_state"] = self.login_state_path
            state_loaded = True
        self._context = await self._browser.new_context(**context_kwargs)
        self._page = await self._context.new_page()
        self._page.set_default_timeout(self.timeout * 1000)
        self._log(
            f"浏览器已启动: headless={mode}"
            + (f"，已加载登录态 {self.login_state_path}" if state_loaded else "，无可用登录态")
        )

    async def _restart_browser(self, headless: bool) -> None:
        """关闭当前浏览器并按指定模式重启（有头 ↔ 无头切换用）。"""
        if self._browser:
            try:
                await self._browser.close()
            except Exception:
                pass
        if self._playwright:
            try:
                await self._playwright.stop()
            except Exception:
                pass
        self._page = None
        self._context = None
        self._browser = None
        self._playwright = None
        await self._ensure_browser(headless=headless)

    async def _save_login_state(self) -> bool:
        """把当前登录态（cookie + localStorage）落盘，供下次无头运行复用。"""
        if self._context is None:
            return False
        try:
            directory = os.path.dirname(self.login_state_path)
            if directory:
                os.makedirs(directory, exist_ok=True)
            await self._context.storage_state(path=self.login_state_path)
            self._log(f"登录态已保存: {self.login_state_path}")
            return True
        except Exception as exc:
            self._log(f"登录态保存失败: {exc}", "warning")
            return False

    async def _is_login_form_page(self) -> bool:
        """页面是否还是登录表单（以密码框可见为准，比看 URL 可靠）。"""
        for selector in self.selectors["login_password"]:
            try:
                element = await self._page.query_selector(selector)
                if element and await element.is_visible():
                    return True
            except Exception:
                continue
        return False

    async def _restore_headless(self) -> bool:
        """登录态刚保存，验证它能否在无头下复用；能就切回无头，别让窗口一直挂着。"""
        await self._restart_browser(headless=True)
        ok = False
        try:
            await self._page.goto(self.url, wait_until="networkidle")
            await self._page.wait_for_load_state("networkidle")
            ok = not await self._is_login_form_page()
        except Exception as exc:
            self._log(f"无头模式登录态校验异常: {exc}", "warning")
        if ok:
            self._log("已切回无头模式，浏览器窗口已关闭")
            return True
        # 登录态没法在无头下复用（例如平台把凭据放在 sessionStorage 里，
        # storage_state 存不到），本次就保持可视化，不再反复闪窗口。
        self._headless_restore_failed = True
        self._log("登录态无法在无头下复用，本次保持可视化模式", "warning")
        await self._restart_browser(headless=False)
        try:
            await self._page.goto(self.url, wait_until="networkidle")
            await self._page.wait_for_load_state("networkidle")
        except Exception:
            pass
        return False

    async def _maybe_restore_headless(self) -> None:
        """本次是为输验证码临时降级成有头的，登录态存好后就该切回无头。"""
        if (
            self._headed_fallback_used
            and self.headless
            and not self._headless_restore_failed
        ):
            await self._restore_headless()

    async def _after_login_ready(self) -> None:
        """登录成功后的收尾：留证截图 + 日志预览。"""
        try:
            await self._page.screenshot(path="output/workorder_after_login.png", full_page=False)
            page_text = await self._page.evaluate(
                "() => document.body.innerText.substring(0, 500)"
            )
            self._log(f"页面文本预览: {page_text[:200]}")
        except Exception:
            pass

    async def _ensure_login_session(self) -> bool:
        """在长时间空闲后查询前重新确认登录态，过期时走正常人工登录流程。"""
        if not self._logged_in:
            return await self.login()
        if self._page is None:
            self._logged_in = False
            return await self.login()
        now = time.monotonic()
        if now - self._last_login_check < self.login_refresh_interval_seconds:
            return True
        try:
            # 只加载工单入口，不修改任何工单；之后 search_batch 会再次进入注册工单页。
            await self._page.goto(self.url, wait_until="domcontentloaded", timeout=self.timeout * 1000)
            if await self._is_login_form_page():
                self._logged_in = False
                return await self.login()
            self._last_login_check = now
            await self._save_login_state()
            self._log("长时间空闲后的登录态探测通过，已刷新本地登录态")
            return True
        except Exception as exc:
            self._log(f"登录态探测异常，将重新建立会话：{exc}", "warning")
            self._logged_in = False
            return await self.login()

    async def login(self) -> bool:
        if self._logged_in:
            return True
        await self._ensure_browser()
        try:
            self._log(f"导航到工单系统: {self.url}")
            await self._page.goto(self.url, wait_until="networkidle")
            await self._page.wait_for_load_state("networkidle")

            # 保存的登录态还有效时，页面直接落在工作台，不会出现登录表单。
            # 登录页有图形验证码，无头下不可能人工输入，全靠这一条路径。
            if not await self._is_login_form_page():
                self._logged_in = True
                self._last_login_check = time.monotonic()
                self._log("已复用保存的登录态，无需重新输验证码")
                await self._save_login_state()  # 顺带刷新 cookie 有效期
                await self._after_login_ready()
                return True

            # 走到这里就必须人工输验证码了。
            if self.headless:
                if not self.headless_fallback:
                    self._log(
                        "无头模式下登录态已失效，且 headless_fallback_to_headed=false，"
                        "本次跳过工单查询（需重新登录时请临时改回可视化模式）",
                        "error",
                    )
                    return False
                self._log("无头模式下登录态已失效，临时弹出浏览器，请人工输入验证码")
                await self._restart_browser(headless=False)
                self._headed_fallback_used = True
                await self._page.goto(self.url, wait_until="networkidle")
                await self._page.wait_for_load_state("networkidle")
                if not await self._is_login_form_page():
                    self._logged_in = True
                    self._last_login_check = time.monotonic()
                    self._log("切换后登录态已生效，免验证码登录成功")
                    await self._save_login_state()
                    await self._after_login_ready()
                    await self._maybe_restore_headless()
                    return True

            login_url = self._page.url
            self._log(f"当前登录页: {login_url}")
            await self._page.screenshot(path="output/workorder_login.png", full_page=False)

            # 填账号
            account_filled = False
            for sel in self.selectors["login_account"]:
                try:
                    el = await self._page.wait_for_selector(sel, timeout=3000)
                    if el:
                        await el.fill("")
                        await el.type(self.username)
                        account_filled = True
                        self._log(f"账号已填写 (selector={sel})")
                        break
                except Exception:
                    continue

            if not account_filled:
                self._log("未找到账号输入框", "error")
                return False

            # 填密码
            password_filled = False
            for sel in self.selectors["login_password"]:
                try:
                    el = await self._page.wait_for_selector(sel, timeout=3000)
                    if el:
                        await el.fill("")
                        await el.type(self.password)
                        password_filled = True
                        self._log(f"密码已填写 (selector={sel})")
                        break
                except Exception:
                    continue

            if not password_filled:
                self._log("未找到密码输入框", "error")
                return False

            self._log("=" * 50)
            self._log("账号密码已自动填写")
            self._log(">>> 请在浏览器中输入验证码并点击登录按钮 <<<")
            self._log("=" * 50)
            self._log(f"等待登录完成... (最长等待 {self.login_wait_seconds} 秒)")

            try:
                await self._page.wait_for_url(
                    lambda url: "/login" not in url,
                    timeout=self.login_wait_seconds * 1000,
                )
                self._log("检测到页面跳转，登录成功")
                await self._page.wait_for_load_state("networkidle")
                await asyncio.sleep(2)
                self._logged_in = True
                self._last_login_check = time.monotonic()
                await self._after_login_ready()
            except Exception:
                current = self._page.url
                if "/login" not in current and "workbench" in current:
                    self._log("已在工单系统页面，登录成功")
                    self._logged_in = True
                    self._last_login_check = time.monotonic()
                else:
                    self._log(f"登录超时，当前URL: {current}", "error")
                    return False

            # 登录态落盘：下次无头运行直接复用，不必再人工输验证码。
            await self._save_login_state()
            await self._maybe_restore_headless()
            return True
        except Exception as e:
            self._log(f"工单系统登录失败: {e}", "error")
            return False

    async def _find_visible_navigation_item(self, text: str, selector_key: str):
        """返回可见的菜单项及其定位来源，避免命中页面中不相关的同名文字。"""
        candidates = [("精确文本", self._page.get_by_text(text, exact=True))]
        candidates.extend(
            (f"selector={selector}", self._page.locator(selector))
            for selector in self.selectors.get(selector_key, [])
        )
        for source, locator in candidates:
            try:
                if await locator.count() < 1:
                    continue
                target = locator.first
                if await target.is_visible():
                    return target, source
            except Exception:
                continue
        return None, ""

    async def _click_navigation_item(self, text: str, selector_key: str) -> bool:
        """只点击精确菜单项；找不到或不可见时由上层走安全兜底。"""
        target, source = await self._find_visible_navigation_item(text, selector_key)
        if target is None:
            return False
        try:
            await target.click(timeout=min(int(self.timeout * 1000), 5000))
            self._log(f"工单导航已点击“{text}” ({source})")
            return True
        except Exception as exc:
            self._log(f"工单导航点击“{text}”失败: {exc}", "warning")
            return False

    async def _find_visible_registration_order_under_parent(self):
        """只在“工单管理”菜单项的子树中寻找可见的“注册工单”。"""
        parent, _ = await self._find_visible_navigation_item(
            "工单管理", "workorder_management"
        )
        if parent is None:
            return None, ""
        try:
            # 当前平台的“工单管理”文字节点位于 li 内，子菜单也是该 li 的后代。
            menu_scope = parent.locator("xpath=..")
            child = menu_scope.get_by_text("注册工单", exact=True)
            if await child.count() > 0 and await child.first.is_visible():
                return child.first, "工单管理子菜单"
        except Exception:
            pass
        return None, ""

    async def _click_registration_order_under_parent(self) -> bool:
        """点击已展开“工单管理”下的注册工单，绝不点击页面其他同名元素。"""
        target, source = await self._find_visible_registration_order_under_parent()
        if target is None:
            return False
        try:
            await target.click(timeout=min(int(self.timeout * 1000), 5000))
            self._log(f"工单导航已点击“注册工单” ({source})")
            return True
        except Exception as exc:
            self._log(f"工单导航点击“注册工单”失败: {exc}", "warning")
            return False

    async def _is_registration_order_page(self) -> bool:
        """以固定路由和查询控件共同确认已进入注册工单页。"""
        if self.registration_order_path not in (self._page.url or ""):
            return False
        for selector in self.selectors["company_input"]:
            try:
                element = await self._page.query_selector(selector)
                if element and await element.is_visible():
                    return True
            except Exception:
                continue
        return False

    async def _wait_for_registration_order_page(self) -> bool:
        """等待 SPA 完成菜单跳转；不等待 networkidle，以免被长连接卡住。"""
        deadline = time.monotonic() + max(float(self.timeout), 3.0)
        while time.monotonic() < deadline:
            if await self._is_registration_order_page():
                return True
            await asyncio.sleep(0.25)
        return await self._is_registration_order_page()

    def _registration_order_url(self) -> str:
        return urljoin(self.url, self.registration_order_path)

    async def _navigate_to_order_list(self) -> bool:
        """按“工单管理 → 注册工单”导航，并确认查询页才允许继续。"""
        if await self._is_registration_order_page():
            self._log("已处于注册工单查询页，跳过菜单导航")
            return True

        register_item, _ = await self._find_visible_registration_order_under_parent()
        if register_item is None:
            parent_clicked = await self._click_navigation_item(
                "工单管理", "workorder_management"
            )
            if parent_clicked:
                await asyncio.sleep(0.3)
            else:
                self._log("未找到“工单管理”菜单，无法执行标准两级导航", "warning")

        if await self._click_registration_order_under_parent():
            if await self._wait_for_registration_order_page():
                self._log("已通过“工单管理 → 注册工单”进入查询页")
                await self._page.screenshot(
                    path="output/workorder_reg_page.png", full_page=False
                )
                return True
            self._log("已点击“注册工单”，但未出现预期查询页", "warning")
        else:
            self._log("未找到可点击的“注册工单”菜单", "warning")

        # 仅当已失败的两级菜单无法确认时，使用已经验证过的固定页面地址兜底。
        fallback_url = self._registration_order_url()
        try:
            self._log(f"尝试固定注册工单地址兜底: {fallback_url}", "warning")
            await self._page.goto(fallback_url, wait_until="domcontentloaded")
            if await self._wait_for_registration_order_page():
                self._log("已通过固定地址进入注册工单查询页")
                await self._page.screenshot(
                    path="output/workorder_reg_page.png", full_page=False
                )
                return True
        except Exception as exc:
            self._log(f"固定注册工单地址访问失败: {exc}", "warning")

        self._log("无法确认注册工单查询页，本批不会继续执行查询", "error")
        await self._page.screenshot(path="output/workorder_nav_fail.png", full_page=False)
        return False

    def _canonical_dropdown_value(self, field_name: str, value: str) -> str:
        """返回用于网页下拉框的规范值；原始值仍保留在输出行中。"""
        text = str(value or "").strip()
        if field_name != "所属代理" or not text:
            return text
        # 先收敛“向善 / 向善”这类同名多邮箱拼接值，再做别名映射。
        text = self._dedupe_joined_names(text)
        canonical = self.agent_aliases.get(normalize_text(text), "")
        if canonical and normalize_text(canonical) != normalize_text(text):
            key = (text, canonical)
            if key not in self._logged_agent_aliases:
                self._logged_agent_aliases.add(key)
                self._log(
                    f"所属代理名称规范化: {text} → {canonical}（仅用于网页查询）",
                    "warning",
                )
            return canonical
        return text

    async def _find_best_dropdown_option(self, value_norm: str):
        """在下拉选项已展开时做安全的代理名称近似匹配。

        仅对短代理名称启用一字符级容错，且要求至少两个字符位置一致；
        国家和服务项目仍然只使用精确/包含匹配，避免误选业务项目。
        """
        if not value_norm or len(value_norm) > 8:
            return None
        best = None
        best_score = 0.0
        for sel in self.selectors["dropdown_option"]:
            try:
                options = await self._page.query_selector_all(sel)
            except Exception:
                continue
            for opt in options:
                try:
                    if not await opt.is_visible():
                        continue
                    text = (await opt.inner_text()).strip()
                    opt_norm = normalize_text(text)
                    if not opt_norm:
                        continue
                    ratio = difflib.SequenceMatcher(None, value_norm, opt_norm).ratio()
                    positional = sum(
                        1 for a, b in zip(value_norm, opt_norm) if a == b
                    )
                    # 代理名称一般只有 2~8 个字符；短名称允许一处形近字，
                    # 但必须保留至少两个同位置字符，降低误选风险。
                    short_typo = (
                        len(value_norm) <= 4
                        and len(opt_norm) == len(value_norm)
                        and positional >= max(2, len(value_norm) - 1)
                        and ratio >= 0.60
                    )
                    if ratio >= 0.82 or short_typo:
                        if ratio > best_score:
                            best = opt
                            best_score = ratio
                except Exception:
                    continue
        if best is not None:
            try:
                text = (await best.inner_text()).strip()
            except Exception:
                text = ""
            self._log(
                f"所属代理近似匹配: 输入={value_norm}, 页面选项={text}, 相似度={best_score:.2f}",
                "warning",
            )
        return best

    async def _select_dropdown(
        self,
        field_name: str,
        value: str,
        selector_key: Optional[str] = None,
    ) -> bool:
        """
        通用下拉框选择: 找到标签旁边的下拉框，点击展开，选对应选项
        field_name: 字段中文名（如"所属代理""国家""服务项目""状态"）
        value: 要选的值
        """
        if not value:
            self._log(f"{field_name} 值为空，跳过选择")
            return True

        value_norm = normalize_text(value)

        # 截图所示页面使用 input placeholder 作为四个查询控件的稳定定位。
        # 找到精确控件后不再退回“遍历所有下拉框”，否则可能误选注册类型。
        direct_key = selector_key or {
            "所属代理": "agent_input",
            "国家": "country_input",
            "服务项目": "service_item_input",
        }.get(field_name)
        if direct_key:
            for sel in self.selectors.get(direct_key, []):
                try:
                    control = await self._page.wait_for_selector(
                        sel, state="visible", timeout=self.field_timeout_ms
                    )
                    if not control:
                        continue
                    tag = await control.evaluate("el => el.tagName.toLowerCase()")
                    # Ant Design/Element UI 的可见占位文字在容器内，真正
                    # 接收键盘输入的 input 是容器后代且通常没有 placeholder。
                    # 先取内部搜索框，仍以容器作为点击/展开目标。
                    input_control = control
                    if tag not in ("input", "select", "textarea"):
                        try:
                            nested = await control.query_selector(
                                'input[role="combobox"], '
                                'input.ant-select-selection-search-input, '
                                'input.el-input__inner, input'
                            )
                            if nested and await nested.is_visible():
                                input_control = nested
                        except Exception:
                            input_control = control
                    if tag == "select":
                        await control.select_option(label=value)
                        self._log(f"{field_name} 选择: {value} (精确控件)")
                        return True

                    # 这是可搜索下拉框时，先点击，再把值输入控件，触发平台的
                    # 异步选项加载。非搜索下拉框则先等待已展开的静态选项。
                    await control.click()
                    option = await self._wait_for_dropdown_option(
                        value_norm, timeout=min(1.0, self.dropdown_timeout_seconds)
                    )
                    if option is None:
                        readonly = await input_control.get_attribute("readonly")
                        disabled = await input_control.get_attribute("disabled")
                        if readonly is None and disabled is None:
                            try:
                                await input_control.fill("")
                                await input_control.type(value, delay=30)
                            except Exception:
                                # 有些组件把键盘事件挂在外层容器，输入框本身
                                # 不接受 fill/type，此时退回键盘输入。
                                try:
                                    await self._page.keyboard.press("Control+A")
                                    await self._page.keyboard.type(value, delay=30)
                                except Exception:
                                    pass
                        option = await self._wait_for_dropdown_option(
                            value_norm, timeout=self.dropdown_timeout_seconds
                        )
                    if option is None and field_name == "所属代理":
                        # 搜索词可能没有直接命中（例如代理名称存在一字差异）。
                        # 清空搜索词后再从完整可见选项中做一次安全近似匹配。
                        try:
                            readonly = await input_control.get_attribute("readonly")
                            disabled = await input_control.get_attribute("disabled")
                            if readonly is None and disabled is None:
                                await input_control.fill("")
                                await asyncio.sleep(0.2)
                        except Exception:
                            pass
                        option = await self._find_best_dropdown_option(value_norm)
                    if option:
                        await option.click()
                        await asyncio.sleep(0.2)
                        self._log(f"{field_name} 选择: {value} (精确控件)")
                        return True
                    await self._page.keyboard.press("Escape")
                    available = await self._visible_dropdown_texts(limit=8)
                    suffix = f"；当前可见选项={available}" if available else ""
                    self._log(
                        f"{field_name} 未找到匹配选项: {value}{suffix}",
                        "warning",
                    )
                    return False
                except Exception as exc:
                    self._log(f"{field_name} 控件操作失败 ({sel}): {exc}", "warning")
                    continue

            # 页面存在明确的字段控件时，不再遍历全页面所有下拉框；那会把
            # “注册类型”等无关控件误当成目标，并造成数分钟的重复等待。
            await self._log_visible_query_controls(field_name)
            self._log(f"{field_name} 未找到可用的精确下拉控件", "warning")
            return False

        # 策略1: 找到字段标签，然后找旁边的 select 或下拉组件
        label_selectors = [
            f'label:has-text("{field_name}")',
            f'span:has-text("{field_name}")',
            f'div:has-text("{field_name}") >> nth=0',
        ]

        for label_sel in label_selectors:
            try:
                label = await self._page.query_selector(label_sel)
                if not label:
                    continue

                # 找标签附近的 select 元素
                select = await label.evaluate_handle(
                    "el => el.parentElement?.querySelector('select, .el-select, .ant-select, [class*=\"select\"], [class*=\"dropdown\"]')"
                )
                if select:
                    select_el = select.as_element()
                    if select_el:
                        # 如果是原生 select
                        tag = await select_el.evaluate("el => el.tagName.toLowerCase()")
                        if tag == "select":
                            await select_el.select_option(label=value)
                            self._log(f"{field_name} 选择: {value} (原生select)")
                            return True
                        # 如果是自定义下拉组件，点击展开
                        await select_el.click()
                        # 找选项
                        option = await self._wait_for_dropdown_option(
                            value_norm, timeout=self.dropdown_timeout_seconds
                        )
                        if option:
                            await option.click()
                            await asyncio.sleep(0.3)
                            self._log(f"{field_name} 选择: {value} (自定义下拉)")
                            return True
            except Exception:
                continue

        # 策略2: 直接找页面上所有下拉框，按顺序匹配
        all_selects = await self._page.query_selector_all(
            '.el-select, .ant-select, select, [class*="select"]:not([class*="selected"]):not(input)'
        )
        for sel_el in all_selects:
            try:
                tag = await sel_el.evaluate("el => el.tagName.toLowerCase()")
                if tag == "select":
                    options = await sel_el.query_selector_all("option")
                    for opt in options:
                        opt_text = await opt.inner_text()
                        if value_norm in normalize_text(opt_text) or normalize_text(opt_text) in value_norm:
                            await sel_el.select_option(label=opt_text.strip())
                            self._log(f"{field_name} 选择: {opt_text.strip()} (原生select遍历)")
                            return True
                else:
                    # 自定义下拉，点击展开
                    await sel_el.click()
                    option = await self._wait_for_dropdown_option(
                        value_norm, timeout=self.dropdown_timeout_seconds
                    )
                    if option:
                        await option.click()
                        await asyncio.sleep(0.3)
                        self._log(f"{field_name} 选择: {value} (自定义下拉遍历)")
                        return True
                    # 关闭下拉
                    await self._page.keyboard.press("Escape")
                    await asyncio.sleep(0.2)
            except Exception:
                continue

        self._log(f"{field_name} 未找到匹配选项: {value}", "warning")
        return False

    async def _log_visible_query_controls(self, failed_field: str = ""):
        """控件定位失败时只输出结构信息，不输出客户/账号等输入值。"""
        try:
            controls = await self._page.evaluate(
                """() => Array.from(document.querySelectorAll(
                    'input, textarea, select, .ant-select, .el-select, [role="combobox"]'
                )).filter(el => {
                    const r = el.getBoundingClientRect();
                    const s = getComputedStyle(el);
                    return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
                }).slice(0, 30).map(el => ({
                    tag: (el.tagName || '').toLowerCase(),
                    placeholder: el.getAttribute('placeholder') || '',
                    role: el.getAttribute('role') || '',
                    cls: typeof el.className === 'string' ? el.className.slice(0, 100) : '',
                    text: (el.innerText || el.textContent || '').trim().replace(/\\s+/g, ' ').slice(0, 50)
                }))"""
            )
            self._log(
                f"{failed_field} 控件诊断(仅结构): {json.dumps(controls, ensure_ascii=False)}",
                "warning",
            )
        except Exception as exc:
            self._log(f"{failed_field} 控件诊断失败: {exc}", "warning")

    async def _wait_for_dropdown_option(self, value_norm: str, timeout: float):
        """等待可见下拉选项，兼容平台异步加载选项的情况。"""
        deadline = time.monotonic() + max(0.1, timeout)
        while time.monotonic() < deadline:
            option = await self._find_dropdown_option(value_norm)
            if option is not None:
                return option
            await asyncio.sleep(0.2)
        return await self._find_dropdown_option(value_norm)

    async def _visible_dropdown_texts(self, limit: int = 8) -> List[str]:
        """返回当前可见的少量选项文字，帮助定位下拉数据/映射问题。"""
        texts = []
        for sel in self.selectors["dropdown_option"]:
            try:
                options = await self._page.query_selector_all(sel)
            except Exception:
                continue
            for opt in options:
                try:
                    if not await opt.is_visible():
                        continue
                    text = (await opt.inner_text()).strip()
                    if text and text not in texts:
                        texts.append(text)
                    if len(texts) >= limit:
                        return texts
                except Exception:
                    continue
        return texts

    async def _find_dropdown_option(self, value_norm: str):
        """在下拉框展开后，找到匹配的选项"""
        option_selectors = self.selectors["dropdown_option"]
        for sel in option_selectors:
            options = await self._page.query_selector_all(sel)
            for opt in options:
                try:
                    if not await opt.is_visible():
                        continue
                    opt_text = await opt.inner_text()
                    opt_norm = normalize_text(opt_text.strip())
                    if value_norm in opt_norm or opt_norm in value_norm:
                        return opt
                except Exception:
                    continue
        return None

    async def _fill_company_name(self, company: str) -> bool:
        """填入公司名称（文本输入框）"""
        # 找到公司名称输入框
        company_selectors = self.selectors["company_input"]

        for sel in company_selectors:
            try:
                el = await self._page.wait_for_selector(
                    sel, state="visible", timeout=self.field_timeout_ms
                )
                if el:
                    await el.fill("", timeout=self.field_timeout_ms)
                    await el.type(company, delay=15, timeout=self.field_timeout_ms)
                    self._log(f"公司名称已填入: {company} (selector={sel})")
                    return True
            except Exception as exc:
                self._log(f"公司输入控件不可用 ({sel}): {exc}", "warning")
                continue

        # 兜底: 找所有 text input
        all_inputs = await self._page.query_selector_all('input[type="text"], input:not([type])')
        for inp in all_inputs:
            try:
                placeholder = await inp.get_attribute("placeholder") or ""
                if any(kw in placeholder for kw in ["公司", "客户", "名称", "请输入"]):
                    await inp.fill("", timeout=self.field_timeout_ms)
                    await inp.type(company, delay=15, timeout=self.field_timeout_ms)
                    self._log(f"公司名称已填入: {company} (兜底selector)")
                    return True
            except Exception as exc:
                self._log(f"公司输入兜底控件不可用: {exc}", "warning")
                continue

        self._log(f"未找到公司名称输入框", "error")
        return False

    @staticmethod
    def _add_calendar_months(value, months: int):
        """给日期加日历月，月末按目标月份最后一天钳制。"""
        if isinstance(value, datetime):
            value = value.date()
        if not isinstance(value, date_type):
            return None
        month_index = value.year * 12 + (value.month - 1) + int(months)
        year, month_zero = divmod(month_index, 12)
        month = month_zero + 1
        day = min(value.day, calendar.monthrange(year, month)[1])
        return date_type(year, month, day)

    def _email_date_window(self, mail_date):
        """返回邮件发件日期前后一个日历月的闭区间日期。"""
        if self.date_window_months <= 0:
            return None
        if isinstance(mail_date, datetime):
            anchor = mail_date.date()
        elif isinstance(mail_date, date_type):
            anchor = mail_date
        else:
            parsed = self._parse_date_value(mail_date)
            anchor = parsed.date() if parsed else None
        if anchor is None:
            return None
        return (
            self._add_calendar_months(anchor, -self.date_window_months),
            self._add_calendar_months(anchor, self.date_window_months),
        )

    @staticmethod
    def _format_date_input(value) -> str:
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d")
        if isinstance(value, date_type):
            return value.strftime("%Y-%m-%d")
        return str(value or "")[:10]

    async def _fill_date_input(self, field_name: str, value, selector_key: str) -> bool:
        """填写页面日期输入框；找不到控件时交给本地日期筛选兜底。"""
        date_text = self._format_date_input(value)
        if not date_text:
            return False
        timeout_ms = min(self.field_timeout_ms, 1800)
        for sel in self.selectors.get(selector_key, []):
            try:
                control = await self._page.wait_for_selector(
                    sel, state="visible", timeout=timeout_ms
                )
                if not control:
                    continue
                target = control
                tag = await control.evaluate("el => el.tagName.toLowerCase()")
                if tag not in ("input", "textarea"):
                    nested = await control.query_selector(
                        'input[placeholder], input[aria-label], input'
                    )
                    if nested and await nested.is_visible():
                        target = nested
                await target.click()
                try:
                    await target.fill(date_text, timeout=timeout_ms)
                except Exception:
                    await self._page.keyboard.press("Control+A")
                    await self._page.keyboard.type(date_text, delay=30)
                try:
                    await target.press("Tab")
                except Exception:
                    pass
                self._log(f"{field_name}已填入: {date_text} (selector={sel})")
                return True
            except Exception as exc:
                self._log(f"{field_name}控件操作失败 ({sel}): {exc}", "warning")
        return False

    async def _fill_order_date_window(self, date_window) -> bool:
        """填写下单开始/结束日期；任一控件缺失时仍保留本地筛选。"""
        if not date_window:
            self._last_date_filter_status = "邮件日期为空，未设置网页日期筛选"
            return False
        date_from, date_to = date_window
        from_ok = await self._fill_date_input(
            "下单开始日期", date_from, "order_date_from_input"
        )
        to_ok = await self._fill_date_input(
            "下单结束日期", date_to, "order_date_to_input"
        )
        if from_ok and to_ok:
            self._last_date_filter_status = "网页已设置邮件日期前后1个月"
        elif from_ok or to_ok:
            self._last_date_filter_status = "网页日期控件部分设置，已使用本地日期筛选"
        else:
            self._last_date_filter_status = "网页日期控件未找到，已使用本地日期筛选"
        if not (from_ok and to_ok):
            self._log(self._last_date_filter_status, "warning")
        return from_ok and to_ok

    async def _clear_inputs(self):
        """只清空注册工单查询页的目标字段，不触碰页面其他控件。"""
        selectors = []
        selectors.extend(self.selectors.get("company_input", []))
        selectors.extend(self.selectors.get("agent_input", []))
        selectors.extend(self.selectors.get("country_input", []))
        selectors.extend(self.selectors.get("service_item_input", []))
        selectors.extend(self.selectors.get("order_date_from_input", []))
        selectors.extend(self.selectors.get("order_date_to_input", []))

        seen = set()
        for sel in selectors:
            if sel in seen:
                continue
            seen.add(sel)
            try:
                inputs = await self._page.query_selector_all(sel)
            except Exception:
                continue
            for inp in inputs:
                try:
                    if not await inp.is_visible():
                        continue
                    tag = await inp.evaluate("el => el.tagName.toLowerCase()")
                    if tag not in ("input", "textarea", "select"):
                        nested = await inp.query_selector(
                            'input[role="combobox"], '
                            'input.ant-select-selection-search-input, '
                            'input.el-input__inner, input'
                        )
                        if nested and await nested.is_visible():
                            inp = nested
                            tag = "input"
                        else:
                            # 下拉容器没有可编辑的内部 input；选择新值时
                            # 会由 _select_dropdown 直接替换单选值。
                            continue
                    readonly = await inp.get_attribute("readonly")
                    disabled = await inp.get_attribute("disabled")
                    if readonly is None and disabled is None:
                        await inp.fill("", timeout=self.field_timeout_ms)
                except Exception as exc:
                    self._log(f"清空查询控件失败 ({sel}): {exc}", "warning")

        # 只点击带有目标查询输入框的下拉组件中的清空按钮。
        for container_sel in (".el-select", ".ant-select"):
            try:
                containers = await self._page.query_selector_all(container_sel)
            except Exception:
                continue
            for container in containers:
                try:
                    if not await container.is_visible():
                        continue
                    target_input = await container.query_selector(
                        'input[role="combobox"], '
                        'input.ant-select-selection-search-input, '
                        'input[placeholder="选择所属代理"], input[placeholder*="所属代理"], '
                        'input[placeholder="国家"], input[placeholder*="国家"], '
                        'input[placeholder="服务项目"], input[placeholder*="服务项目"]'
                    )
                    if not target_input:
                        continue
                    clear_btn = await container.query_selector(
                        ".el-select__clear, .ant-select-clear, [class*='clear']"
                    )
                    if clear_btn and await clear_btn.is_visible():
                        await clear_btn.click(timeout=self.field_timeout_ms)
                except Exception as exc:
                    self._log(f"清空下拉选择失败: {exc}", "warning")

    async def _click_query_button(self) -> bool:
        """点击查询按钮"""
        btn_selectors = self.selectors["query_button"]
        for sel in btn_selectors:
            try:
                btn = await self._page.wait_for_selector(sel, timeout=3000)
                if btn:
                    await btn.click(timeout=self.field_timeout_ms)
                    self._log(f"点击查询按钮: {sel}")
                    # 工单页是 SPA，常驻连接可能永远达不到 networkidle；
                    # 这里只把它当"能省则省"的加速窗口，超时（常态）不再刷
                    # warning，真正的就绪判据交给 _wait_for_table_refresh。
                    try:
                        await self._page.wait_for_load_state(
                            "networkidle", timeout=min(self.timeout * 1000, 1500)
                        )
                    except Exception:
                        pass
                    # 固定 sleep(2) 已改为可配置的短沉降：指纹轮询会检测到
                    # 表格刷新后立刻返回，渲染快时不必陪着睡满 2 秒。
                    if self.query_click_settle_seconds:
                        await asyncio.sleep(self.query_click_settle_seconds)
                    return True
            except Exception:
                continue
        self._log("未找到查询按钮", "error")
        return False

    async def _debug_screenshot(self, path: str) -> None:
        """过程截图（查询前/查询后）。

        无头下没人看这些图，截一次要花上百毫秒且逐条累积，因此直接跳过；
        有头时照旧保留，方便人工确认填写是否正确。真正需要排障时，
        解析失败路径仍会单独产出 output/debug/ 下的整页快照。
        """
        if self.headless:
            return
        try:
            await self._page.screenshot(path=path, full_page=False)
        except Exception:
            pass

    async def _table_fingerprint(self) -> str:
        """提取当前可见查询表格的轻量指纹，用于判断结果是否刷新。"""
        try:
            return await self._page.evaluate(
                """() => Array.from(document.querySelectorAll('table'))
                    .filter(t => {
                        const r = t.getBoundingClientRect();
                        const s = getComputedStyle(t);
                        return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
                    })
                    .map(t => (t.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 1200))
                    .filter(Boolean).join('||')"""
            ) or ""
        except Exception:
            return ""

    async def _wait_for_table_refresh(self, before_fingerprint: str):
        """查询后等待结果表就绪。

        两个判据，满足其一即返回：

        1. **指纹变化且连续两次采样一致** → 新结果已渲染稳定。
           要求"连续两次一致"是为了避开表格清空/半渲染那一瞬：只看"变了"
           就返回，可能把正在加载的空表读成 0 条，下游会误判成漏单。
        2. 等满 table_refresh_timeout_seconds 仍未变化 → 本次结果与上次
           逐字相同（典型场景：两次都命中"暂无数据"），指纹永远不会变。
           旧实现这里会一直等到 dropdown 的 5 秒上限，是模糊重查慢的主因。
        """
        deadline = time.monotonic() + self.table_refresh_timeout_seconds
        last = None
        while time.monotonic() < deadline:
            current = await self._table_fingerprint()
            if current and current != before_fingerprint:
                if current == last:
                    return
                last = current
            await asyncio.sleep(0.1)

    async def _fill_query_form(
        self,
        company: str,
        agent: str = "",
        country: str = "",
        service_item: str = "",
        status: str = "",
        mail_date=None,
        date_window=None,
    ) -> bool:
        """清空并逐项填写注册工单查询条件。

        该方法同时用于实时查询和缓存命中时的可视化重放。缓存命中时只重放
        输入，不点击查询按钮，查询结果仍以缓存为准，避免把“看输入过程”和
        “重新访问工单系统”混为一件事。
        """
        await self._clear_inputs()
        if self.input_step_delay:
            await asyncio.sleep(min(0.3, self.input_step_delay))

        # 1. 填公司名称
        if not await self._fill_company_name(company):
            return False
        if self.input_step_delay:
            await asyncio.sleep(self.input_step_delay)

        # 2. 选所属代理
        if agent:
            if not await self._select_dropdown("所属代理", agent, "agent_input"):
                return False
            if self.input_step_delay:
                await asyncio.sleep(self.input_step_delay)

        # 3. 选国家
        if country:
            if not await self._select_dropdown("国家", country, "country_input"):
                return False
            if self.input_step_delay:
                await asyncio.sleep(self.input_step_delay)

        # 4. 选服务项目
        if service_item:
            if not await self._select_dropdown(
                "服务项目", service_item, "service_item_input"
            ):
                return False
            if self.input_step_delay:
                await asyncio.sleep(self.input_step_delay)

        # 5. 选状态
        if status:
            if not await self._select_dropdown("状态", status):
                return False
            if self.input_step_delay:
                await asyncio.sleep(self.input_step_delay)

        # 页面日期条件使用邮件发件日期前后一个日历月；即使页面控件定位失败，
        # search_one 后面的本地 match_records 仍会再次按同一窗口筛选。
        if mail_date is not None:
            await self._fill_order_date_window(
                date_window if date_window is not None else self._email_date_window(mail_date)
            )
            if self.input_step_delay:
                await asyncio.sleep(self.input_step_delay)
        return True

    async def search_one(
        self,
        company: str,
        agent: str = "",
        country: str = "",
        service_item: str = "",
        status: str = "",
        mail_date=None,
    ) -> List[Dict]:
        """
        查询单条数据:
        1. 填公司名称
        2. 选所属代理（下拉框）
        3. 选国家（下拉框）
        4. 选服务项目（下拉框）
        5. 选状态（下拉框）
        6. 点查询
        7. 无结果 → 模糊匹配公司名重查
        """
        if not self._logged_in:
            success = await self.login()
            if not success:
                return []

        self._last_query_status = ""
        self._last_lookup_mode = "精确查询"
        self._last_fuzzy_term = ""
        company, customer_note = self._sanitize_customer_for_query(company)
        if customer_note and not company:
            self._last_query_status = "查询条件填写失败"
            self._log(f"公司名称异常，跳过本条查询: {customer_note}", "error")
            return []
        if customer_note:
            self._log(f"公司名称已清洗: {customer_note}", "warning")

        # 代理别名只影响网页输入和缓存键，阶段一原始代理字段保持不变。
        query_agent = self._canonical_dropdown_value("所属代理", agent)
        date_window = self._email_date_window(mail_date)
        self._last_date_window = date_window
        if date_window is not None:
            self._last_date_filter_status = "待设置网页日期筛选"
        elif mail_date:
            self._last_date_filter_status = "邮件日期无法解析，未设置网页日期筛选"
        else:
            self._last_date_filter_status = "邮件日期为空，未设置网页日期筛选"

        fill_kwargs = {
            "company": company,
            "agent": query_agent,
            "country": country,
            "service_item": service_item,
            "status": status,
        }
        # 兼容现有调用方/测试替换的 _fill_query_form；只有阶段二行确实有
        # 发件日期时才增加日期参数。
        if mail_date is not None:
            fill_kwargs["mail_date"] = mail_date
            fill_kwargs["date_window"] = date_window

        # 断点续查：命中缓存仍可在真实浏览器中重放输入，便于人工观察。
        project_norm = self._compose_project_for_query(query_agent, country, service_item)
        weee_category_scope = country == "德国" and service_item.upper() == "WEEE"
        cached = None if self.force_live_query else self.get_cached_result(
            query_agent, company, project_norm, date_window=date_window
        )
        # 旧缓存只有主工单列表，没有“品类明细”时不能直接给出 WEEE 专项结论；
        # 自动放弃这一次缓存并实时打开品类页签，后续新缓存会保留该明细。
        if weee_category_scope and cached is not None and not any(
            isinstance(order, dict) and order.get("品类明细") for order in cached
        ):
            self._log("德国 WEEE 工单缓存缺少品类明细，改为实时读取注册工单品类页签", "info")
            cached = None
        if cached is not None:
            # 旧缓存没有 lookup_mode；仍可复用已有结果，但在输出中明确标成缓存，
            # 避免操作人员误以为本次重新执行了模糊查找。
            if not self._last_lookup_mode or self._last_lookup_mode == "精确查询":
                self._last_lookup_mode = "缓存复用"
            if self.show_cached_inputs:
                try:
                    replayed = await self._fill_query_form(**fill_kwargs)
                except Exception as exc:
                    replayed = False
                    self._log(f"缓存结果展示输入时发生异常: {exc}", "warning")
                if replayed:
                    self._last_query_status = "缓存复用-已展示"
                    self._log(
                        f"工单缓存命中，已在浏览器展示查询条件（不重新点击查询）: "
                        f"公司={company}, 代理={agent}, 项目={project_norm} "
                        f"({len(cached)} 条)"
                    )
                    if self.cached_input_hold:
                        await asyncio.sleep(self.cached_input_hold)
                else:
                    self._last_query_status = "缓存复用-展示失败"
                    self._log(
                        f"工单缓存命中，但浏览器查询条件展示失败: "
                        f"公司={company}, 代理={agent}, 项目={project_norm}",
                        "warning",
                    )
            else:
                self._log(
                    f"工单缓存命中（未操作浏览器）: 公司={company}, 代理={agent}, "
                    f"项目={project_norm} ({len(cached)} 条)"
                )
                self._last_query_status = "缓存复用"
            return self._orders_from_jsonable(cached)

        workorders = []

        try:
            # 1-5. 清空并填写所有查询条件
            filled = await self._fill_query_form(**fill_kwargs)
            if not filled:
                self._last_query_status = "查询条件填写失败"
                return []

            # 6. 点查询；记录查询前表格指纹，防止异步刷新时读取上一条结果
            before_fingerprint = await self._table_fingerprint()
            await self._debug_screenshot("output/workorder_before_query.png")
            clicked = await self._click_query_button()
            if not clicked:
                self._last_query_status = "查询按钮失败"
                return []

            await self._wait_for_table_refresh(before_fingerprint)
            await self._debug_screenshot("output/workorder_after_query.png")

            # 7. 读取结果
            workorders = await self._extract_table_data()
            if self._last_table_parse_status == "failed":
                # 结构读不出来时必须显式失败：若照旧当成 0 条，下游会把“没读到”
                # 当成“没录单”，整批邮件被误判成漏单，并且会写进缓存扩散。
                self._last_query_status = "结果表解析失败"
                self._log(
                    f"结果表解析失败，本条不写缓存也不判定漏单: "
                    f"{self._last_table_parse_note}",
                    "error",
                )
                return []
            self._last_query_status = "实时查询"
            self._last_lookup_mode = "精确查询"
            self._log(f"精确查询结果: {len(workorders)} 条 (公司={company})")

            # 8. 无结果 → 模糊匹配公司名称（只改公司名，其他不动）
            if not workorders and len(company) > 2:
                self._log(f"精确查询无结果，尝试模糊匹配公司名称...")

                # 候选词只做安全变体：繁简转换、去空白、去类型说明括号、
                # 逐级去开头行政区划；公司后缀保留不动，地名类括号不删。
                fuzzy_terms = self._get_fuzzy_company_terms(company)
                if len(fuzzy_terms) > 1:
                    self._log(
                        f"模糊匹配候选词({len(fuzzy_terms) - 1}): "
                        + " → ".join(term for term in fuzzy_terms if term != company)
                    )
                else:
                    self._log(
                        "公司名去掉类型说明括号与行政区划后与原名相同，无可用的模糊候选词",
                        "warning")


                for term in fuzzy_terms:
                    if term == company:
                        continue
                    # 只清空公司名输入框，其他下拉不动
                    company_inputs = await self._page.query_selector_all(
                        'input[placeholder*="公司"], input[placeholder*="客户"], input[placeholder*="名称"]'
                    )
                    for ci in company_inputs:
                        try:
                            await ci.fill("")
                            await ci.type(term)
                        except Exception:
                            continue

                    fuzzy_before_fingerprint = await self._table_fingerprint()
                    await self._click_query_button()
                    await self._wait_for_table_refresh(fuzzy_before_fingerprint)
                    workorders = await self._extract_table_data()
                    if self._last_table_parse_status == "failed":
                        # 页面结构读不出来时，换个词再查一样读不出来，白等一轮。
                        # 直接停，由函数末尾的复查把状态标成"结果表解析失败"。
                        self._log("模糊匹配阶段结果表仍解析失败，停止换词重试", "error")
                        workorders = []
                        break
                    if workorders:
                        self._last_lookup_mode = "模糊匹配"
                        self._last_fuzzy_term = term
                        self._log(f"模糊匹配命中: '{term}' → {len(workorders)} 条")
                        break

                if not workorders:
                    self._last_lookup_mode = "模糊匹配-无结果"
                    self._log(f"模糊匹配也无结果: {company}", "warning")

            if weee_category_scope and workorders:
                category_records = await self._extract_weee_category_detail_data()
                for order in workorders:
                    if isinstance(order, dict):
                        order["品类明细"] = category_records
                        order["德国WEEE品类明细来源"] = "注册工单品类明细页签" if category_records else "未读取到品类明细"

        except Exception as e:
            self._last_query_status = "查询异常"
            self._log(f"工单查询失败: {e}", "error")

        # 模糊匹配阶段会再次读结果表，同样可能解析失败；写缓存前必须复查，
        # 否则故障会被当成“0 条”缓存下来并向后续跑批扩散。
        if self._last_table_parse_status == "failed":
            self._last_query_status = "结果表解析失败"

        # 只有完成了网页查询才写缓存；字段/按钮/页面异常不能被保存成“0条”。
        if self._last_query_status in {"实时查询"}:
            self.save_to_cache(
                query_agent, company, project_norm,
                self._orders_to_jsonable(workorders), date_window=date_window,
            )
        return workorders

    def _compose_project_for_query(self, agent: str, country: str, service_item: str) -> str:
        """country+service_item 拼成一个稳定的项目键供缓存去重"""
        return f"{country or ''}_{service_item or ''}"

    @classmethod
    def _strip_admin_prefix(cls, name: str) -> str:
        """反复剥掉公司名开头的行政区划。

        「深圳市宝安区宝安某某有限公司」→「宝安某某有限公司」
        「上海某某有限公司」（名字里没有"市"字）→ 原样返回，不乱动
        """
        result = (name or "").strip()
        while True:
            match = cls._ADMIN_PREFIX_RE.match(result)
            if not match:
                break
            rest = result[match.end():].strip()
            if len(rest) < cls._MIN_STRIPPED_LEN:
                break
            # 剥完只剩个公司后缀（如「某某夜市有限公司」被切成「有限公司」）
            # 说明这刀切在了主体名上，宁可放弃剥离。
            if rest.startswith(cls._COMPANY_SUFFIXES):
                break
            result = rest
        return result

    @classmethod
    def _admin_unit_spans(cls, name: str) -> List[Tuple[int, int]]:
        """切出名字开头**连续的**行政区划单元区间（左→右）。

        「深圳市宝安区宝安某某有限公司」→ [(0, 3), (3, 6)]
        「上海某某科技有限公司」（名字里没有"市"字）→ []
        只在开头连续匹配，一碰到非区划内容就停，不会去动主体名里的字样。
        """
        text = name or ""
        spans: List[Tuple[int, int]] = []
        pos = 0
        while pos < len(text) and len(spans) < cls._MAX_ADMIN_UNITS:
            match = cls._ADMIN_PREFIX_RE.match(text[pos:])
            if not match:
                break
            end = pos + match.end()
            spans.append((pos, end))
            pos = end
        return spans

    @classmethod
    def _has_subject_name(cls, name: str) -> bool:
        """名字抠掉「开头区划 + 结尾公司后缀」后，是否还剩下的主体名。

        这是"这一级区划能不能删"的判据：
          - 「xx区有限公司」→ 抠完什么都不剩，主体名被剃光了 → 不能这么删
          - 「xx示例4328e80a有限公司」→ 还剩「xx」，主体名健在 → 可以这么删
        少这道闸门，就会搜出把区名当主体名的废词（「xx区有限公司」）。
        """
        text = (name or "").strip()
        if not text:
            return False
        for suffix in cls._COMPANY_SUFFIXES:
            if text.endswith(suffix):
                text = text[: -len(suffix)]
                break
        while True:
            match = cls._ADMIN_PREFIX_RE.match(text)
            if not match:
                break
            text = text[match.end():]
        return bool(text.strip())

    @classmethod
    def _strip_brackets(cls, name: str) -> str:
        """只删「主体类型说明」类括号，地名/国别括号原样保留（白名单，严格侧）。

        「xx有限公司（个体工商户）」      →「xx有限公司」   ← 删
        「某某（自然人独资）有限公司」    →「某某有限公司」 ← 删
        「深圳某某（上海）贸易有限公司」  → 原样           ← 保留（括号是实体标识）
        「某某（中国）有限公司」          → 原样           ← 保留（删了会撞上「某某有限公司」）
        「xx有限公司（）」                →「xx有限公司」   ← 空括号照删（无信息量）

        白名单的取舍见 _BRACKET_TYPE_HINTS 的注释：漏删只是少查一轮（保守），
        误删会拼出真实存在的别家公司名（假"已录单"），后者不可接受。
        """
        def repl(match: "re.Match") -> str:
            inner = match.group(0)[1:-1].strip()
            if not inner or any(h in inner for h in cls._BRACKET_TYPE_HINTS):
                return ""
            return match.group(0)

        return cls._BRACKET_RE.sub(repl, name or "").strip()

    def _get_fuzzy_company_terms(self, company: str) -> List[str]:
        """生成公司名称的模糊匹配搜索词（命中即停、按优先级排序）。

        业务口径（用户 2026-09-14 明确）：
          - 括号**只删「主体类型说明」类**（个体工商户 / 自然人独资 / 分公司…）：
              「xx有限公司（个体工商户）」→「xx有限公司」
            括号里装的是地名/国别（实体标识）时**整块保留**：
              「深圳某某（上海）贸易有限公司」→ 原样（不做任何缩减）
              「某某（中国）有限公司」→ 原样（删了会撞上真实存在的「某某有限公司」）
          - 开头行政区划**逐级单独删除**，不是一次剃光：
              「xx示例f0c9e053有限公司」
                  → 删市级「xx示例ae34c927有限公司」
                  → 删区级「xx示例4328e80a有限公司」
                  → 全删「xx有限公司」
              「xx示例7f9e6398有限公司」（区级后面直接接后缀、没有主体名）
                  → 删市级会得到废词「xx区有限公司」，跳过；
                    只产出「xx示例4328e80a有限公司」
          - **示例05b0c9dc有限公司 / 有限责任公司 / 集团…）必须保留，不得删除**；
            也不做「取前 N 字」这类截短。后缀是主体名的边界，删掉后搜出来的
            是一大批无关公司（如「青岛积米」会命中所有同前缀企业）。

        候选顺序即查询顺序：去括号（仅类型说明括号）→ 逐级单删（先删最靠前的
        区划，命中率最高，见用户第一条口径）→ 全删区划 → 原始全名垫底
        （调用方会跳过它，留着只是让日志里的候选清单完整）。
        """
        original = (company or "").strip()
        if not original:
            return []

        terms: List[str] = []

        def add(term: str) -> None:
            term = (term or "").strip()
            if not term or term == original or term in terms or len(term) < 2:
                return
            # 剃完只剩公司后缀、或只剩一串区划，都说明这刀切进了主体名，弃用。
            if len(term) < self._MIN_STRIPPED_LEN or not self._has_subject_name(term):
                return
            terms.append(term)

        # 先在原始写法上生成安全候选，再补一份简体写法。平台搜索框通常能
        # 忽略大小写，但不一定能处理繁简体；简体转换只改变字形，不删除公司
        # 后缀、国家或项目实体，因此风险远低于“取前 N 字”的宽泛搜索。
        base_names = [original]
        simplified = to_simplified(original)
        if simplified and simplified != original:
            base_names.append(simplified)

        for base_name in base_names:
            # 先去掉主体类型说明括号：xx有限公司（个体工商户）→ xx有限公司
            # 顺序固定为"先去括号再动区划"：括号里若有"市/区"字样，先去掉才不会
            # 干扰区划正则的贪婪边界。
            no_bracket = self._strip_brackets(base_name) or base_name
            add(no_bracket)

            # 只去除空白作为格式变体；保留所有标点、国别括号和法律后缀，
            # 防止把两个不同主体拼成一个宽泛查询词。
            compact = re.sub(r"\s+", "", no_bracket)
            add(compact)

            # 逐级单独删除：xx示例f0c9e053有限公司 → xx示例ae34c927有限公司 / xx示例4328e80a有限公司
            for start, end in self._admin_unit_spans(no_bracket):
                add(no_bracket[:start] + no_bracket[end:])

            # 示例6fafa6bf有限公司 → xx有限公司）；只有一级区划时
            # 这一步和上面的单删结果重合，add() 会自动去重。
            add(self._strip_admin_prefix(no_bracket))

        # 原始全名放最后（调用方会跳过它，保留只是让日志里的候选清单完整）
        terms.append(original)
        return terms

    @staticmethod
    def _rows_to_records(headers: List, rows: List) -> List[Dict]:
        """把「表头 + 表体行」按列索引映射成工单字典。

        这里是防“key 塌缩”的关键：表头为空时必须退化成 col_N，绝不能让每个
        字段名都变成空字符串——一旦塌缩，下游 match_records 取不到客户/项目/
        代理，整批邮件会被判成漏单。表头重名时加 #n 后缀，避免互相覆盖。
        """
        names: List[str] = []
        used: Dict[str, int] = {}
        for idx, raw in enumerate(headers or []):
            name = str(raw or "").strip() or f"col_{idx}"
            count = used.get(name, 0) + 1
            used[name] = count
            names.append(name if count == 1 else f"{name}#{count}")

        # Ant Design 的结果表首列是选择框、末列是滚动条占位，两者都没有表头也没有
        # 文本。这类列不产出字段，否则每条记录都会多出一堆 col_N 噪音；但“无表头
        # 却有值”的列必须保留（宁可退化成 col_N，也不能丢数据）。
        width = max(
            [len(headers or [])] + [len(cells or []) for cells in (rows or [])] or [0]
        )
        drop = set()
        for idx in range(width):
            raw = headers[idx] if idx < len(headers or []) else ""
            if str(raw or "").strip():
                continue
            has_value = any(
                str((cells[idx] if idx < len(cells or []) else "") or "").strip()
                for cells in (rows or [])
            )
            if not has_value:
                drop.add(idx)

        records: List[Dict] = []
        for cells in rows or []:
            row_data: Dict[str, str] = {}
            for idx, cell in enumerate(cells or []):
                if idx in drop:
                    continue
                key = names[idx] if idx < len(names) else f"col_{idx}"
                row_data[key] = str(cell or "").strip()
            # 补约定字段名（保留原始列名，不覆盖已有值）
            for raw_name, canonical in RESULT_COLUMN_ALIASES.items():
                if raw_name in row_data and canonical not in row_data:
                    row_data[canonical] = row_data[raw_name]
            if any(row_data.values()):
                records.append(row_data)
        return records

    async def _dump_result_page_snapshot(self, note: str) -> None:
        """解析失败时抓一次现场：结构摘要 + 截图 + 整页 HTML，放在 output/debug/。

        只在第一次失败时抓，避免每条都写一遍。页面上“表头读不到”这类问题光看日志
        猜不出来，必须拿到真实 DOM 才能定位。
        """
        if self._result_page_dumped:
            return
        self._result_page_dumped = True
        try:
            os.makedirs(self.debug_dump_dir, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            summary_path = os.path.join(self.debug_dump_dir, f"table_structure_{stamp}.txt")
            try:
                summary = await self._page.evaluate(TABLE_DEBUG_JS)
                with open(summary_path, "w", encoding="utf-8") as f:
                    json.dump(summary, f, ensure_ascii=False, indent=2)
            except Exception as e:
                summary_path = f"(结构摘要保存失败: {e})"

            html_path = os.path.join(self.debug_dump_dir, f"result_page_{stamp}.html")
            try:
                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(await self._page.content())
            except Exception as e:
                html_path = f"(HTML 保存失败: {e})"

            png_path = os.path.join(self.debug_dump_dir, f"result_page_{stamp}.png")
            try:
                await self._page.screenshot(path=png_path, full_page=False)
            except Exception as e:
                png_path = f"(截图保存失败: {e})"

            self._log(
                f"已保存结果页现场快照（{note}）: 结构={summary_path} | HTML={html_path} | 截图={png_path}"
            )
        except Exception as e:
            self._log(f"保存结果页现场快照失败: {e}", "warning")

    async def _mark_table_parse_failed(self, note: str) -> List[Dict]:
        """统一处理解析失败：标状态、写日志、抓一次现场快照，返回空列表。

        调用方拿到空列表后必须当成“未比对”，绝不能当成“没录单”。
        """
        self._last_table_parse_status = "failed"
        self._last_table_parse_note = note
        self._log(f"{note}，本条标记为未比对", "error")
        await self._dump_result_page_snapshot(note)
        return []

    async def _extract_table_data(self) -> List[Dict]:
        """从页面读取查询结果表格。

        表头只认 thead / .el-table__header-wrapper / .ant-table-header 里的 th；
        读不到表头时显式失败（_last_table_parse_status=failed），由调用方转成
        “未比对”，而不是返回 0 条被下游当成“没录单”。
        """
        self._last_table_parse_status = "ok"
        self._last_table_parse_note = ""

        try:
            candidates = await self._page.evaluate(TABLE_EXTRACT_JS)
        except Exception as e:
            return await self._mark_table_parse_failed(f"结果表结构读取异常: {e}")

        candidates = [c for c in (candidates or []) if isinstance(c, dict)]
        if not candidates:
            return await self._mark_table_parse_failed("页面上没有可见的结果表格")

        # 择优顺序：① 读到表头 ② 读到行 ③ 行多 ④ 列全。
        # 表头必须优先于行数：否则“表体表有 1 行、表头表 0 行”时，行数多的那组会
        # 胜出，而它恰恰没有表头——正是让整批邮件被判漏单的那个坑。
        best = max(
            candidates,
            key=lambda c: (
                bool(c.get("headerFound")),
                bool(c.get("rows")),
                len(c.get("rows") or []),
                len(c.get("headers") or []),
            ),
        )
        headers = [str(h or "").strip() for h in (best.get("headers") or [])]
        rows = best.get("rows") or []
        empty_flag = bool(best.get("empty"))

        if not best.get("headerFound"):
            if not rows and empty_flag:
                # 页面明确显示“暂无数据”，这时才允许判定为真的 0 条。
                self._log("结果表: 0 行（页面显示暂无数据）")
                return []
            return await self._mark_table_parse_failed(
                f"结果表结构无法识别：找到 {len(candidates)} 个表格，"
                f"读到 {len(rows)} 行但表头为空"
            )

        if not rows and not empty_flag:
            # 读到表头却一行都没有、页面也没显示空状态：无法区分“查询确实为空”和
            # “表体没读到”。宁可报错让上层标未比对，也不能默认成 0 条漏单。
            return await self._mark_table_parse_failed(
                f"结果表读到表头({len(headers)} 列)但 0 行，页面也未显示空状态，"
                "无法确认是查询无结果还是表体未读到"
            )

        records = self._rows_to_records(headers, rows)

        # 列名完全对不上，说明抓错了页面上别的表格；此时宁可报错也不要产出
        # “0 匹配”的假漏单。列名清单可用 config.result_table_columns 覆盖。
        if records and not (set(headers) & self.result_columns):
            return await self._mark_table_parse_failed(
                f"结果表列名与预期不符（实际列名={headers[:12]}）"
            )

        self._log(
            f"结果表: {len(headers)} 列, {len(records)} 行, 表头={headers[:8]}"
            + ("（页面显示暂无数据）" if not records and empty_flag else "")
        )
        return records

    async def _extract_weee_category_detail_data(self) -> List[Dict]:
        """读取注册工单页面的“品类明细”页签。

        该页签不是所有查询结果都提供；读不到时返回空列表，专项比对会显示
        “待人工核对”，绝不把“页面没有品类页签”当成品类缺失。主工单列表的
        查询结果仍由 ``_extract_table_data`` 负责，两个解析状态互不污染。
        """
        if self._page is None:
            return []
        tab_locator = self._page.locator(
            ".work-order-register-container .tab-header .tab-btn, .tab-header .tab-btn"
        )
        try:
            count = await tab_locator.count()
        except Exception:
            return []
        if count < 2:
            self._log("德国 WEEE 专项：注册工单页面未发现“品类明细”页签，待人工核对", "warning")
            return []

        switched = False
        try:
            # 当前工单系统的页签顺序固定为“工单列表 / 品类明细”。优先用可见文本
            # 定位，中文页面乱码或定制主题时再退回第二个页签。
            category_tab = tab_locator.filter(has_text="品类明细")
            if await category_tab.count():
                await category_tab.first.click()
            else:
                await tab_locator.nth(1).click()
            switched = True
            await asyncio.sleep(min(0.8, max(0.2, self.query_click_settle_seconds)))
            candidates = await self._page.evaluate(TABLE_EXTRACT_JS)
            candidates = [c for c in (candidates or []) if isinstance(c, dict)]
            if not candidates:
                self._log("德国 WEEE 专项：品类明细页签没有可见表格，待人工核对", "warning")
                return []
            best = max(
                candidates,
                key=lambda c: (
                    bool(c.get("headerFound")), bool(c.get("rows")),
                    len(c.get("rows") or []), len(c.get("headers") or []),
                ),
            )
            headers = [str(h or "").strip() for h in (best.get("headers") or [])]
            rows = best.get("rows") or []
            if not best.get("headerFound") or not rows:
                self._log("德国 WEEE 专项：品类明细表没有可解析行，待人工核对", "warning")
                return []
            records = self._rows_to_records(headers, rows)
            self._log(f"德国 WEEE 专项：读取品类明细 {len(records)} 行，表头={headers[:8]}")
            return records
        except Exception as exc:
            self._log(f"德国 WEEE 专项：读取品类明细失败，待人工核对：{exc}", "warning")
            return []
        finally:
            if switched:
                try:
                    list_tab = tab_locator.filter(has_text="工单列表")
                    if await list_tab.count():
                        await list_tab.first.click()
                    else:
                        await tab_locator.nth(0).click()
                    await asyncio.sleep(0.15)
                except Exception:
                    pass

    def _extract_country_from_project(self, project: str) -> str:
        """从项目名称中提取国家"""
        countries = [
            "德国", "法国", "意大利", "西班牙", "荷兰", "波兰",
            "瑞典", "比利时", "爱尔兰", "葡萄牙", "奥地利",
            "英国", "匈牙利", "捷克", "丹麦", "芬兰", "挪威",
            "罗马尼亚", "保加利亚", "希腊", "克罗地亚",
        ]
        for c in countries:
            if c in project:
                return c
        return ""

    def _extract_service_item_from_project(self, project: str) -> str:
        """从项目名称中提取服务项目类型"""
        if "WEEE" in project.upper():
            return "WEEE"
        if "电池" in project:
            return "电池法"
        if "包装" in project:
            return "包装法"
        if "EPR" in project.upper():
            return "EPR"
        if "一次性塑料" in project:
            return "一次性塑料"
        if "BAT" in project.upper():
            return "BAT"
        return ""

    @staticmethod
    def _apply_weee_category_audit(row: Dict, workorders: List[Dict], query_status: str = "") -> None:
        """把德国 WEEE 品类逐项核对结果写回阶段二行。"""
        if str(row.get("德国WEEE专项") or "") != "是":
            return
        result = compare_weee_categories(row.get("德国WEEE品类明细") or [], workorders or [])
        if query_status and query_status not in {"实时查询", "缓存复用", "缓存复用-已展示"}:
            result["status"] = "pending"
            result["reason"] = f"注册工单查询未完成：{query_status}"
        row["德国WEEE品类核对"] = json.dumps(result, ensure_ascii=False, separators=(",", ":"))
        row["德国WEEE品类状态"] = {
            "matched": "已匹配",
            "missing": "未找到",
            "pending": "待人工核对",
        }.get(result.get("status"), "待人工核对")
        row["德国WEEE专项说明"] = result.get("reason", "")

    async def search_batch(self, mail_rows: List[Dict]) -> List[Dict]:
        """
        批量查询: 逐条数据进入注册工单页面，填条件，查询，读取结果
        每条数据只与本行查询结果比对（按行隔离，防跨行串扰）
        返回: 已写入 是否已录单/匹配状态 字段的行列表
        风控: 条目间间隔 query_interval_seconds；总耗时超 max_total_seconds 主动终止
        断点续查: 已查询 (代理+公司+项目) 三联键命中缓存直接复用，不再走 RPA
        """
        self.cancelled = False
        self._cancel_notice_logged = False
        self.last_processed_rows = []
        if self._is_cancel_requested():
            return []

        success = await self._await_or_cancel(self._ensure_login_session(), False)
        if self.cancelled:
            return []
        if not success:
            return []

        self._session_start_ts = time.monotonic()
        # 只有确认到注册工单查询页才允许写入查询条件，防止误操作其他页面。
        navigated = await self._await_or_cancel(self._navigate_to_order_list(), False)
        if self.cancelled or not navigated:
            return []

        all_results = []
        total = len(mail_rows)
        terminated = False
        termination_reason = ""
        consecutive_query_failures = 0
        ok_results = 0
        empty_results = 0

        for idx, row in enumerate(mail_rows, 1):
            if self._is_cancel_requested():
                terminated = True
                termination_reason = "用户停止"
                break

            # 总超时检查
            if self._check_total_timeout():
                self._log(
                    f"⛔ 已达到总超时 {self.max_total_seconds}s, 终止于 {idx-1}/{total}, "
                    f"剩余 {total - idx + 1} 条下次续跑",
                    "error",
                )
                terminated = True
                termination_reason = "总超时"
                break

            if not row.get("客户"):
                continue

            company = row.get("客户", "")
            agent = row.get("代理", "")
            project = row.get("项目", "")
            country = self._extract_country_from_project(project)
            service_item = self._extract_service_item_from_project(project)
            mail_date = row.get("date")
            date_window = self._email_date_window(mail_date)
            if date_window:
                row["邮件日期筛选范围"] = (
                    f"{date_window[0].isoformat()} ~ {date_window[1].isoformat()}"
                )
            else:
                row["邮件日期筛选范围"] = ""

            # 客户字段不是公司名（指代、句子、数量描述等）时直接跳过：拿去查只会
            # 白跑一次 RPA，而且“查不到”会被写成漏单，污染结论。
            reject_reason = self._customer_reject_reason(company)
            if reject_reason:
                row["是否已录单"] = "未比对"
                row["匹配状态"] = "客户名称无效-跳过查询"
                row["RPA查询状态"] = "客户名称无效-跳过查询"
                row["查询方式"] = "未查询"
                row["模糊查询词"] = ""
                row["日期筛选说明"] = reject_reason
                self._apply_weee_category_audit(row, [], "客户名称无效-跳过查询")
                row["查询时间戳"] = datetime.now().isoformat(timespec="seconds")
                self._log(
                    f"第 {idx}/{total} 条客户字段不是有效公司名，跳过查询: "
                    f"{company!r}（{reject_reason}）",
                    "warning",
                )
                self.last_processed_rows.append(row)
                all_results.append(row)
                continue

            self._log(f"工单查询 {idx}/{total}: 公司={company}, 代理={agent}, "
                       f"国家={country}, 项目={service_item}, "
                       f"邮件日期范围={row.get('邮件日期筛选范围') or '未设置'}")

            # 工作台“漏单自动重查”会在该行留下强制重查标记；这一行
            # 必须绕过旧缓存，否则重新进入循环仍会得到上次的漏单结论。
            force_before = self.force_live_query
            if row.get("_force_workorder_query"):
                self.force_live_query = True
                self._log("该明细来自漏单回流队列，跳过历史缓存执行实时查询", "info")
            try:
                orders = await self._await_or_cancel(
                    self.search_one(
                        company=company,
                        agent=agent,
                        country=country,
                        service_item=service_item,
                        mail_date=mail_date,
                    ),
                    None,
                )
            finally:
                self.force_live_query = force_before
            if self.cancelled or self._is_cancel_requested():
                terminated = True
                termination_reason = "用户停止"
                break

            # 按行隔离：本行只与本行查询结果比对，防止跨行串扰掩盖漏单
            query_status = self._last_query_status or "查询未执行"
            if query_status in {
                "查询条件填写失败",
                "查询按钮失败",
                "查询异常",
                "查询未执行",
                "缓存复用-展示失败",
                # 结果表读不出来时同样按失败处理：绝不能落到“漏单”，
                # 否则解析故障会被伪装成业务结论。
                "结果表解析失败",
            }:
                row["是否已录单"] = "未比对"
                row["匹配状态"] = query_status
                row["RPA查询状态"] = query_status
                row["查询方式"] = self._last_lookup_mode or "失败"
                row["模糊查询词"] = self._last_fuzzy_term
                row["日期筛选说明"] = self._last_date_filter_status
                reason_detail = (
                    self._last_table_parse_note
                    if query_status == "结果表解析失败"
                    else query_status
                )
                row["_query_error"] = (
                    f"第{idx}条查询条件处理失败：{reason_detail}；"
                    "已跳过本条，后续记录继续执行"
                )
                self._apply_weee_category_audit(row, [], query_status)
                row["查询时间戳"] = datetime.now().isoformat(timespec="seconds")
                self.last_processed_rows.append(row)
                all_results.append(row)
                consecutive_query_failures += 1
                if consecutive_query_failures >= self.max_consecutive_query_failures:
                    self._log(
                        f"连续 {consecutive_query_failures} 条查询失败，"
                        f"第 {idx}/{total} 条后停止本批，避免页面整体异常时空转",
                        "error",
                    )
                    terminated = True
                    termination_reason = "查询失败"
                    break
                self._log(
                    f"第 {idx}/{total} 条未完成（{query_status}），"
                    "已标记当前行并继续后续记录",
                    "warning",
                )
                if idx < total and self.query_interval > 0:
                    await self._await_or_cancel(
                        asyncio.sleep(self.query_interval), None
                    )
                    if self.cancelled:
                        terminated = True
                        termination_reason = "用户停止"
                        break
                continue

            consecutive_query_failures = 0
            row["RPA查询状态"] = query_status
            row["查询方式"] = self._last_lookup_mode or "未记录"
            row["模糊查询词"] = self._last_fuzzy_term
            row["日期筛选说明"] = self._last_date_filter_status
            self.match_records([row], orders)
            self._apply_weee_category_audit(row, orders, query_status)
            all_results.append(row)
            self.last_processed_rows.append(row)
            lookup_detail = row["查询方式"]
            if row["模糊查询词"]:
                lookup_detail += f", 词={row['模糊查询词']}"
            self._log(
                f"  → 查到 {len(orders)} 条工单 (查询状态={query_status}, "
                f"查询方式={lookup_detail})"
            )
            ok_results += 1
            if not orders:
                empty_results += 1

            # 条目间间隔（最后一条不 sleep）
            if idx < total and self.query_interval > 0:
                await self._await_or_cancel(asyncio.sleep(self.query_interval), None)
                if self.cancelled:
                    terminated = True
                    termination_reason = "用户停止"
                    break

        if self.cancelled:
            self._log(
                f"批次查询已由用户停止，已处理 {len(all_results)}/{total} 条；"
                "已完成条目的查询缓存已保留，重跑时会复用。",
                "warning",
            )
        elif terminated:
            elapsed = time.monotonic() - (self._session_start_ts or time.monotonic())
            if termination_reason == "总超时":
                self._log(
                    f"批次查询因总超时终止, 已处理 {len(all_results)}/{total} 条 "
                    f"(耗时 {elapsed:.0f}s)",
                    "warning",
                )
            elif termination_reason == "查询失败":
                self._log(
                    f"批次查询因查询控件/条件失败终止, 已处理 "
                    f"{len(all_results)}/{total} 条 (耗时 {elapsed:.0f}s)",
                    "warning",
                )
            else:
                self._log(
                    f"批次查询终止, 已处理 {len(all_results)}/{total} 条 "
                    f"(耗时 {elapsed:.0f}s)",
                    "warning",
                )

        # 空结果占位提示：查询条件没真正生效时页面同样显示“暂无数据”，
        # 而那会被当成“客户没录单”。全部为空时给一条显眼的提醒供人工核对。
        if ok_results >= 5 and empty_results == ok_results:
            self._log(
                f"⚠ 本批 {ok_results} 条查询结果全部为 0 条。若这些客户并非都未录单，"
                "请核对查询条件是否真的生效（公司名称/所属代理/服务项目是否填进了页面）",
                "warning",
            )

        return all_results

    def match_records(
        self,
        mail_rows: List[Dict],
        workorders: List[Dict],
    ) -> List[Dict]:
        """按“查询返回条数”判定是否漏单。

        查询条件本身已经是“公司 + 代理 + 项目（+ 国家）”的精确条件，平台返回的
        条数就是最可靠的答案，因此：
          - 返回 1 条 → 直接认定已录单，不再用本地模糊比对否决；
          - 返回 ≥2 条 → 无法确定是哪一条，交人工排查；
          - 返回 0 条 → 上游已按公司名做模糊重查，仍为 0 才判漏单。
        本地四维比对（代理 / 客户 / 项目 / 日期）保留为**提示**，只写进匹配状态与
        日期筛选说明，绝不推翻结论——否则本地字段写法差异会制造假漏单。
        """
        results = []
        query_timestamp = datetime.now()

        for row in mail_rows:
            mail_date = self._parse_date_value(row.get("date"))
            # 工单日期 = **邮件发来的日期**（阶段一从邮件读到的发件日期），
            # 与平台返回的「下单日期」是两回事：客户先发邮件（工单日期），
            # 我们随后才在工单系统里下单（下单日期）。所以这一列永远等于邮件日期，
            # 平台那列只写进 `下单日期`，不要互相覆盖。
            row["工单日期"] = mail_date
            date_window = self._email_date_window(mail_date)
            if date_window and not row.get("邮件日期筛选范围"):
                row["邮件日期筛选范围"] = (
                    f"{date_window[0].isoformat()} ~ {date_window[1].isoformat()}"
                )
            mail_agent = row.get("代理", "")
            mail_customer = row.get("客户", "")
            mail_project = row.get("项目", "")
            if not str(mail_agent or "").strip():
                # 代理为空只是该维度不可校验，结论需要人工确认，不能算作漏单。
                row["代理校验"] = "代理为空-未按代理筛选，结果需人工确认"

            # 逐条标注本地比对结果，仅用于提示，不参与判定
            checked = []
            for wo in (workorders or []):
                wo_date = self._parse_wo_date(wo)
                date_delta_days = None
                date_ok = True
                if mail_date and wo_date:
                    # 优先使用邮件发件日期前后一个日历月，而不是固定 60 天。
                    # 这里使用 date 比较，避免邮件时区与网页无时区时间相减报错。
                    date_delta_days = abs((mail_date.date() - wo_date.date()).days)
                    if date_window:
                        date_ok = date_window[0] <= wo_date.date() <= date_window[1]
                    else:
                        date_ok = date_delta_days <= self.date_tolerance

                wo_agent = (wo.get("代理", "") or wo.get("agent", "")
                            or wo.get("所属代理", "") or "")
                wo_customer = (wo.get("客户", "") or wo.get("customer", "")
                               or wo.get("公司", "") or wo.get("公司名称", "") or "")
                wo_project = (wo.get("项目", "") or wo.get("project", "")
                              or wo.get("服务项目", "") or "")

                checked.append({
                    # 平台返回的「下单日期」：仅内部用于日期窗口比对，
                    # 写出去的是下面的 `下单日期`（页面原始文本）。
                    "平台下单日期": wo_date,
                    "下单日期": self._wo_order_date_raw(wo),
                    "工单代理": wo_agent,
                    "工单客户": wo_customer,
                    "工单项目": wo_project,
                    "日期异常": bool(mail_date and wo_date and not date_ok),
                    "日期差天数": date_delta_days,
                    "字段一致": bool(
                        self._match_agent(mail_agent, wo_agent)
                        and fuzzy_match_pair(
                            mail_customer, wo_customer, self.customer_threshold
                        )
                        and self._match_project(mail_project, wo_project)
                    ),
                })

            total = len(checked)
            if total == 1:
                # 查询条件已生效且只返回 1 条 → 直接认定已录单。
                only = checked[0]
                row["是否已录单"] = "是"
                row["下单日期"] = only["下单日期"]
                row["工单记录"] = [only]
                if not only["字段一致"]:
                    row["匹配状态"] = "已录单-字段不一致需复核"
                    row["日期筛选说明"] = (
                        "查询返回 1 条，按规则认定已录单；但本地比对不一致"
                        f"（代理: 名单={mail_agent or '空'} / 工单={only['工单代理'] or '空'}；"
                        f"客户: 名单={mail_customer} / 工单={only['工单客户']}；"
                        f"项目: 名单={mail_project} / 工单={only['工单项目']}），建议人工确认"
                    )
                elif only["日期异常"]:
                    row["匹配状态"] = "已录单-日期异常需复核"
                    row["日期筛选说明"] = (
                        "查询返回 1 条，按规则认定已录单；"
                        f"但该工单的下单日期与邮件日期相差 {only['日期差天数']} 天，"
                        f"超出邮件日期前后{self.date_window_months}个月窗口，建议人工确认"
                    )
                else:
                    row["匹配状态"] = "精确匹配"
            elif total >= 2:
                # 返回多条说明条件不够唯一，按约定交人工排查（不再自动收敛成 1 条）。
                row["是否已录单"] = "待复核"
                row["匹配状态"] = "多匹配-人工排查"
                dated = [c for c in checked if c.get("日期差天数") is not None]
                dated.sort(key=lambda c: c["日期差天数"])
                # 下单日期取「日期最近」的那条候选，保证与状态说明里的天数对得上。
                nearest_one = dated[0] if dated else checked[0]
                row["下单日期"] = nearest_one["下单日期"]
                row["工单记录"] = checked
                note = f"查询返回 {total} 条，无法确定唯一工单，交人工排查"
                if dated:
                    nearest = [
                        c for c in dated if c["日期差天数"] == dated[0]["日期差天数"]
                    ]
                    note += (
                        f"；日期最近的一条相差 {dated[0]['日期差天数']} 天"
                        + ("（并列，需人工确认）" if len(nearest) > 1 else "（供参考）")
                    )
                else:
                    note += "；候选工单均无下单日期，无法按日期缩小范围"
                if not any(c["字段一致"] for c in checked):
                    note += (
                        "；注意：候选工单与名单的代理/客户/项目全部对不上，"
                        "请核对查询条件是否真的生效"
                    )
                row["日期筛选说明"] = note
            else:
                row["是否已录单"] = "否"
                row["下单日期"] = ""
                row["匹配状态"] = "漏单"
                row["工单记录"] = []

            row["查询时间戳"] = query_timestamp

            results.append(row)

        return results

    @staticmethod
    def _parse_date_value(value) -> Optional[datetime]:
        if isinstance(value, datetime):
            return value
        if isinstance(value, date_type):
            return datetime(value.year, value.month, value.day)
        if not value:
            return None
        date_str = str(value).strip()
        if not date_str:
            return None
        formats = [
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%Y/%m/%d",
            "%Y年%m月%d日", "%d/%m/%Y",
        ]
        for fmt in formats:
            try:
                return datetime.strptime(date_str[:19] if len(date_str) > 19 else date_str, fmt)
            except ValueError:
                continue
        return None

    def _parse_wo_date(self, wo: Dict) -> Optional[datetime]:
        """解析**平台返回的「下单日期」**（页面上那一列）。

        仅供内部做“下单日期 vs 邮件日期”的窗口比对，**不要**把这个值写回
        `工单日期` 字段——`工单日期` 的定义是“邮件发来的日期”（发件日期）。
        """
        date_value = (
            wo.get("日期") or wo.get("下单日期") or wo.get("下单时间")
            or wo.get("工单日期") or wo.get("date")
            or wo.get("创建时间") or wo.get("注册时间") or ""
        )
        return self._parse_date_value(date_value)

    def _wo_order_date_raw(self, wo: Dict) -> str:
        """取平台返回的**原始**下单日期文本，不解析、不格式化。

        与 `_parse_wo_date` 同源，但保留页面原样（如 `2026-08-20`），
        供最终核对结果直接展示，方便人工和工单系统上的那列逐字对照。
        """
        value = (
            wo.get("下单日期") or wo.get("日期") or wo.get("下单时间")
            or wo.get("工单日期") or wo.get("date")
            or wo.get("创建时间") or wo.get("注册时间") or ""
        )
        return str(value or "").strip()

    def _match_agent(self, mail_agent: str, wo_agent: str) -> bool:
        """四维比对里的代理维度。

        任一侧为空都不能直接判否：阶段一没抓到代理时，就算工单表里确实有这条
        记录也永远匹配不上，整批会被误判成“漏单”。代理缺失只是这一个维度无法
        校验，是否录单应交给客户 + 项目 + 日期继续判断，并在行内标注供人工确认。
        """
        if not mail_agent or not wo_agent:
            return True
        mail_agent = self._canonical_dropdown_value("所属代理", mail_agent)
        wo_agent = self._canonical_dropdown_value("所属代理", wo_agent)
        if fuzzy_match_pair(mail_agent, wo_agent, 80):
            return True
        if mail_agent in wo_agent or wo_agent in mail_agent:
            return True
        return False

    def _match_project(self, mail_project: str, wo_project: str) -> bool:
        if not mail_project or not wo_project:
            return False
        if fuzzy_match_pair(mail_project, wo_project, 85):
            return True
        return False

    async def close(self):
        # 收尾前再刷一次登录态：跑完一批通常已过一段时间，写回最新的 cookie
        # 能让下次无头运行继续免验证码。
        if self._logged_in and self._context is not None:
            await self._save_login_state()
        try:
            if self._browser:
                await self._browser.close()
        except Exception:
            pass
        try:
            if self._playwright:
                await self._playwright.stop()
        except Exception:
            pass
        self._page = None
        self._context = None
        self._browser = None
        self._playwright = None
        self._logged_in = False
        self._session_start_ts = None

    # ---------------- 断点续查 / 请求间隔 ----------------

    @staticmethod
    def _cache_key(agent: str, company: str, project: str,
                   date_window=None) -> str:
        """规范化键：agent|company|project|日期窗口。

        日期窗口存在时必须纳入缓存键，否则同一家公司不同月份的邮件会
        复用另一封邮件的工单结果，正是多结果/缓存命中过多的主要风险。
        旧调用不带日期时仍保持三联键兼容。
        """
        a = normalize_text(agent or "").strip()
        c = normalize_text(company or "").strip()
        p = normalize_text(project or "").strip()
        key = f"{a}|{c}|{p}"
        if date_window:
            start, end = date_window
            key += f"|{start.isoformat()}|{end.isoformat()}"
        return key

    @staticmethod
    def _dedupe_joined_names(text: str) -> str:
        """把“向善 / 向善”这类重复拼接值收敛成“向善”。

        代理表里同一个代理往往登记了多个邮箱，阶段一多匹配时会按“ / ”拼起来，
        直接拿去选下拉或参与比对必然失败，所以在进入查询/比对前先按规范化名去重。
        """
        text = str(text or "").strip()
        if not text or " / " not in text:
            return text
        seen = set()
        parts = []
        for part in (p.strip() for p in text.split(" / ")):
            if not part:
                continue
            key = normalize_text(part)
            if key in seen:
                continue
            seen.add(key)
            parts.append(part)
        return " / ".join(parts) if parts else text

    # 明显不是公司名的片段：邮件正文里的指代、请求动作或口语片段被误当成客户
    # 字段时，不能拿去查工单系统。
    _CUSTOMER_REJECT_TOKENS = (
        "这家公司", "那家公司", "该公司", "本公司", "贵公司", "我方公司", "客户公司",
        "公司名称", "公司信息", "公司资料", "修改公司", "公司需要", "补充", "待签字",
        "授权书", "麻烦", "请帮忙", "请联系", "以下公司", "上述公司", "这个公司",
    )

    @staticmethod
    def _customer_reject_reason(value) -> str:
        """判断客户字段是否“根本不是公司名”，返回原因（空串表示可用）。"""
        text = "" if value is None else str(value).strip()
        if not text:
            return "客户为空"
        if re.fullmatch(r"\d+\s*家公司", text):
            return "只是数量描述"
        if re.search(r"[。！？；，、]", text):
            return "是句子而非公司名"
        for token in WorkOrderChecker._CUSTOMER_REJECT_TOKENS:
            if token in text:
                return f"含非公司名片段“{token}”"
        latin_company = re.search(
            r"(?:LTD|LIMITED|GMBH|S\.?A\.?S|B\.?V|INC|COMPANY|CORP)",
            text, flags=re.I,
        )
        if len(re.findall(r"[\u4e00-\u9fa5]", text)) < 2 and not latin_company:
            return "有效汉字不足且无公司后缀"
        # 只剩“公司”这类后缀、没有任何名称信息时同样不能拿去查询。
        if re.sub(r"[\s（）()]+", "", text) in {
            "公司", "公司名", "有限公司", "有限责任公司", "股份有限公司",
            "集团有限公司", "集团", "工厂",
        }:
            return "只有公司后缀没有名称"
        return ""

    @staticmethod
    def _sanitize_customer_for_query(value) -> tuple[str, str]:
        """拒绝把整段标题/表格行填入公司查询框，能唯一修复时仅保留公司名。"""
        raw = "" if value is None else str(value)
        text = raw.strip()
        if not text:
            return "", "客户为空"
        if not re.search(r"[\t\r\n]", raw):
            reason = WorkOrderChecker._customer_reject_reason(text)
            return ("", reason) if reason else (text, "")

        candidates = []
        for segment in re.split(r"[\t\r\n]+", raw):
            segment = segment.strip()
            if not segment:
                continue
            candidates.extend(re.findall(
                r"[\u4e00-\u9fa5A-Za-z0-9·]{2,30}?"
                r"(?:有限责任公司|股份有限公司|集团有限公司|有限公司|责任公司|公司)",
                segment,
            ))
            candidates.extend(re.findall(
                r"[A-Za-z][A-Za-z0-9 .,&'’]{2,60}?"
                r"(?:LIMITED|LTD\.?|GMBH|S\.A\.S|SAS|B\.V\.?)\b",
                segment, flags=re.I,
            ))
        candidates = [
            c.strip(" .,-")
            for c in candidates
            if len(c.strip(" .,-")) >= 4
            and not WorkOrderChecker._customer_reject_reason(c.strip(" .,-"))
        ]
        if candidates:
            return max(candidates, key=len), "客户字段含制表符，已提取其中的合法公司名"
        return "", "客户字段含制表符且未找到合法公司名"

    @staticmethod
    def _is_collapsed_record(record) -> bool:
        """判断一条缓存工单记录是否是“字段名全塌缩”的脏数据。

        旧版解析把表体第一行当成表头，于是每行只剩一个空字符串 key（值往往还是
        操作列的“编辑/审核/详情/更多”）。这类记录不含任何可比对字段，必须丢弃，
        否则修好解析后仍会从缓存里复用错误结论。
        """
        if not isinstance(record, dict) or not record:
            return True
        return not any(str(k).strip() for k in record)

    def _load_cache(self) -> Dict[str, dict]:
        if not self.cache_path or not os.path.exists(self.cache_path):
            return {}
        try:
            with open(self.cache_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            self._log(f"查询缓存加载失败，忽略: {e}", "warning")
            return {}
        if not isinstance(data, dict):
            return {}

        cleaned: Dict[str, dict] = {}
        dropped = 0
        for key, entry in data.items():
            if not isinstance(entry, dict):
                dropped += 1
                continue
            orders = entry.get("orders")
            if isinstance(orders, list):
                kept = [o for o in orders if not self._is_collapsed_record(o)]
                if len(kept) != len(orders):
                    dropped += 1
                    if not kept:
                        continue
                    entry = {**entry, "orders": kept}
            cleaned[key] = entry
        if dropped:
            self._log(
                f"查询缓存丢弃 {dropped} 条被表头塌缩污染的记录"
                f"（旧版解析缺陷遗留，重查后会写入正确结果）",
                "warning",
            )
        self._log(f"已加载查询缓存: {len(cleaned)} 条 ({self.cache_path})")
        return cleaned

    def _save_cache(self):
        if not self.cache_path:
            return
        try:
            os.makedirs(os.path.dirname(self.cache_path) or ".", exist_ok=True)
            tmp = self.cache_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self._query_cache, f, ensure_ascii=False, indent=2)
            os.replace(tmp, self.cache_path)
        except Exception as e:
            self._log(f"查询缓存写入失败: {e}", "warning")

    def _is_expired(self, entry: dict) -> bool:
        if self.cache_ttl_days <= 0:
            return False
        ts = entry.get("timestamp")
        if not ts:
            return True
        try:
            age = datetime.now() - datetime.fromisoformat(ts)
            return age.days >= self.cache_ttl_days
        except Exception:
            return True

    def get_cached_result(
        self, agent: str, company: str, project: str, date_window=None
    ) -> Optional[List[Dict]]:
        """命中且未过期返回缓存 orders，否则 None"""
        key = self._cache_key(agent, company, project, date_window=date_window)
        entry = self._query_cache.get(key)
        if not entry:
            return None
        if self._is_expired(entry):
            self._query_cache.pop(key, None)
            return None
        # 新增安全模糊候选后，历史版本的“0 条”缓存可能只是精确查询没有命中，
        # 从未执行过模糊重查。只淘汰这类旧空结果；旧的非空结果仍可正常复用，
        # 不会因为功能升级而把整批缓存全部打穿。
        if not entry.get("fuzzy_version") and not entry.get("orders"):
            self._query_cache.pop(key, None)
            self._save_cache()
            self._log("发现旧版空查询缓存，已失效并允许本次执行安全模糊重查", "info")
            return None
        self._last_lookup_mode = str(entry.get("lookup_mode") or "缓存复用")
        self._last_fuzzy_term = str(entry.get("fuzzy_term") or "")
        return entry.get("orders", [])

    def save_to_cache(
        self, agent: str, company: str, project: str, orders: List[Dict], date_window=None
    ):
        key = self._cache_key(agent, company, project, date_window=date_window)
        window_value = None
        if date_window:
            start, end = date_window
            window_value = [start.isoformat(), end.isoformat()]
        self._query_cache[key] = {
            "orders": orders,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "date_window": window_value,
            # 版本化元数据只影响审计，不改变旧缓存的读取结构；旧缓存仍可复用，
            # 但其来源会显示为“缓存复用”，不会伪装成本次精确/模糊查询。
            "lookup_mode": str(getattr(self, "_last_lookup_mode", "实时查询") or "实时查询"),
            "fuzzy_term": str(getattr(self, "_last_fuzzy_term", "") or ""),
            "fuzzy_version": 2,
        }
        self._save_cache()

    def _check_total_timeout(self) -> bool:
        """总超时检查：True=已超时需终止"""
        if self.max_total_seconds <= 0 or self._session_start_ts is None:
            return False
        return (time.monotonic() - self._session_start_ts) >= self.max_total_seconds

    @staticmethod
    def _orders_to_jsonable(orders: List[Dict]) -> List[Dict]:
        """datetime 转 ISO 字符串以便 JSON 化"""
        out = []
        for wo in orders:
            row = {}
            for k, v in wo.items():
                if isinstance(v, datetime):
                    row[k] = v.isoformat()
                else:
                    row[k] = v
            out.append(row)
        return out

    @staticmethod
    def _orders_from_jsonable(orders: List[Dict]) -> List[Dict]:
        """ISO 字符串还原为 datetime"""
        out = []
        for wo in orders:
            row = {}
            for k, v in wo.items():
                if isinstance(v, str) and len(v) >= 10 and v[4] == "-" and v[7] == "-":
                    try:
                        row[k] = datetime.fromisoformat(v)
                        continue
                    except ValueError:
                        pass
                row[k] = v
            out.append(row)
        return out

    # ---------------- 前置校验 ----------------

    @staticmethod
    def preprocess_rows(
        mail_rows: List[Dict],
        skip_threshold: str = "需人工确认",
    ) -> tuple:
        """
        阶段二前置预处理：
          1. 跳过「代理为空 且 置信度=需人工确认」的条目（不执行RPA查询，单独返回告警）
          2. 按 (代理, 公司, 项目) 去重 — 同组合才查一次，节省 RPA 时间
        返回: (rows_to_query, skipped_rows)

        注: 阶段二独立入口是从阶段一 xlsx 反读的, 那批 dict 的键是「表头名」
        (客户公司名称 / 标准化项目名称), 而一站式流程传的是内部键(客户 / 项目)。
        这里对两种键名都兼容, 否则会全部取到空值 → 全部行被误判为重复而只剩 1 条。
        """
        def _pick(row: dict, *names: str) -> str:
            for n in names:
                v = row.get(n)
                if v is not None and str(v).strip() != "":
                    return str(v).strip()
            return ""

        rows_to_query: List[Dict] = []
        skipped_rows: List[Dict] = []
        seen = set()

        for r in mail_rows:
            agent = _pick(r, "代理") or _pick(r, "代理(可能空)")
            confidence = _pick(r, "置信度")
            customer_raw = _pick(r, "客户", "客户公司名称")
            customer, customer_note = WorkOrderChecker._sanitize_customer_for_query(customer_raw)
            if customer_note:
                if customer:
                    r["客户"] = customer
                    r["_stage2_note"] = customer_note
                else:
                    r["_skip_reason"] = customer_note
                    skipped_rows.append(r)
                    continue

            # 跳过：代理空 + 置信度=需人工确认
            if not agent and confidence == skip_threshold:
                # 保留原 dict 身份；ExcelWriter 需要把它与 all_rows 对应起来。
                r["_skip_reason"] = "代理空且置信度=需人工确认, 阶段二不查RPA, 请人工补全后手动导入"
                skipped_rows.append(r)
                continue

            # 去重：(代理, 公司, 项目)。不同代理的同一客户/项目必须独立核对。
            project = _pick(r, "项目", "标准化项目名称")
            if not customer and not project:
                # 两项皆空 → 无法判定是否重复, 不得合并(否则整批会塌缩成 1 条)
                rows_to_query.append(r)
                continue
            key = (normalize_text(agent), normalize_text(customer), normalize_text(project))
            if key in seen:
                r["_skip_reason"] = "代理-公司-项目重复, 与前条合并查询"
                skipped_rows.append(r)
                continue
            seen.add(key)
            rows_to_query.append(r)

        return rows_to_query, skipped_rows
