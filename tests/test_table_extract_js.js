/**
 * 工单结果表结构读取脚本（modules/workorder_checker.py::TABLE_EXTRACT_JS）的离线验证。
 *
 * 用最小 DOM 桩还原两种真实结构，验证：
 *   表头只来自 thead / .el-table__header-wrapper / .ant-table-header，
 *   表体只来自 .el-table__body-wrapper / .ant-table-body，
 *   绝不会把表体第一行当成表头（这正是历史上整批被判“漏单”的根因），
 *   也不会把 ant 的计量行/“暂无数据”占位行当成数据行。
 *
 * 由 test_offline.py::test_workorder_table_extract_js 调用（需要 node，缺省则跳过）。
 * 直接运行: node tests/test_table_extract_js.js
 */
const fs = require('fs');
const path = require('path');

// ---------- 最小 DOM ----------
class El {
  constructor(tag, classes, attrs) {
    this.tag = tag.toLowerCase();
    this.classes = classes || [];
    this.attrs = attrs || {};
    this.children = [];
    this.parent = null;
    this.ownText = '';
  }
  add(child) { child.parent = this; this.children.push(child); return child; }
  get classList() { return { contains: (c) => this.classes.includes(c) }; }
  get className() { return this.classes.join(' '); }
  getAttribute(n) { return this.attrs[n] === undefined ? null : String(this.attrs[n]); }
  get innerText() { return this.textContent; }
  get textContent() {
    return (this.ownText || '') + this.children.map(c => c.textContent).join('');
  }
  getBoundingClientRect() {
    return this.attrs.__hidden ? { width: 0, height: 0 } : { width: 100, height: 20 };
  }
  querySelectorAll(sel) { return queryAll(this, sel); }
  querySelector(sel) { return queryAll(this, sel)[0] || null; }
  closest(sel) {
    const parts = sel.split(',').map(s => s.trim());
    let node = this;
    while (node) {
      if (parts.some(p => matchesCompound(node, p))) return node;
      node = node.parent;
    }
    return null;
  }
}

function descendants(root) {
  const out = [];
  const walk = (n) => { for (const c of n.children) { out.push(c); walk(c); } };
  walk(root);
  return out;
}

function matchesCompound(el, compound) {
  const tokens = compound.match(/^([a-zA-Z]+)?((?:\.[\w-]+)*)$/);
  if (!tokens) return false;
  const [, tag, classPart] = tokens;
  if (tag && el.tag !== tag.toLowerCase()) return false;
  if (classPart) {
    for (const cls of classPart.split('.').filter(Boolean)) {
      if (!el.classes.includes(cls)) return false;
    }
  }
  return true;
}

function matchesSteps(el, steps) {
  if (!matchesCompound(el, steps[steps.length - 1])) return false;
  let i = steps.length - 2;
  let node = el.parent;
  while (i >= 0 && node) {
    if (matchesCompound(node, steps[i])) i--;
    node = node.parent;
  }
  return i < 0;
}

function queryAll(root, sel) {
  const pool = descendants(root);
  const out = [];
  for (const part of sel.split(',').map(s => s.trim()).filter(Boolean)) {
    // 子选择器 `A > B` 在桩里按后代处理即可（真实浏览器里也是这个结果集）
    const steps = part.replace(/\s*>\s*/g, ' ').split(/\s+/).filter(Boolean);
    for (const el of pool) {
      if (matchesSteps(el, steps) && !out.includes(el)) out.push(el);
    }
  }
  return out;
}

function setGlobalDom(root) {
  global.document = { querySelectorAll: (sel) => queryAll(root, sel) };
  global.getComputedStyle = (el) => ({
    display: el.attrs.__display || 'block',
    visibility: 'visible',
  });
}

// ---------- 构造 el-table ----------
const tr = (cells) => {
  const row = new El('tr');
  for (const [text, colspan] of cells) {
    const cell = new El('td', [], colspan ? { colspan } : {});
    cell.ownText = text;
    row.add(cell);
  }
  return row;
};

function buildElTable({ headers, rows, empty, withFixed = false }) {
  const body = new El('body');
  const root = body.add(new El('div', ['el-table']));

  const headerWrapper = root.add(new El('div', ['el-table__header-wrapper']));
  const headerTable = headerWrapper.add(new El('table', ['el-table__header']));
  const thead = headerTable.add(new El('thead'));
  const headRow = thead.add(new El('tr'));
  for (const [text, colspan] of headers) {
    const th = new El('th', [], colspan ? { colspan } : {});
    th.ownText = text;
    headRow.add(th);
  }

  const bodyWrapper = root.add(new El('div', ['el-table__body-wrapper']));
  const bodyTable = bodyWrapper.add(new El('table', ['el-table__body']));
  const tbody = bodyTable.add(new El('tbody'));
  for (const r of rows) tbody.add(tr(r));
  if (empty) bodyWrapper.add(new El('div', ['el-table__empty-block']));

  if (withFixed) {
    const fixed = root.add(new El('div', ['el-table__fixed-right']));
    const fh = fixed.add(new El('div', ['el-table__header-wrapper']));
    const fht = fh.add(new El('table', ['el-table__header']));
    const fthead = fht.add(new El('thead'));
    const fhr = fthead.add(new El('tr'));
    for (const [text] of headers.slice(0, 2)) {
      const th = new El('th');
      th.ownText = text;
      fhr.add(th);
    }
    const fb = fixed.add(new El('div', ['el-table__body-wrapper']));
    const fbt = fb.add(new El('table', ['el-table__body']));
    const ftbody = fbt.add(new El('tbody'));
    for (const r of rows) ftbody.add(tr(r.slice(0, 2)));
  }
  return body;
}

// ---------- 构造 Ant Design 表格（工单平台真实使用的框架） ----------
// 结构：.ant-table-container > (.ant-table-header > table) + (.ant-table-body > table)
// 表体表没有 thead；空结果时会渲染 ant-table-measure-row 与 ant-table-placeholder。
function buildAntTable({ headers, rows, empty, withCheckbox = true, withPicker = false }) {
  const body = new El('body');
  if (withPicker) {
    // 页面上真实存在的两个日期选择器日历表：不可见，必须被过滤掉
    for (let k = 0; k < 2; k++) {
      const picker = body.add(new El('div', ['ant-picker-body']));
      const pt = picker.add(new El('table', ['ant-picker-content'], { __hidden: true }));
      const pth = pt.add(new El('thead'));
      pth.add(new El('tr')).add(Object.assign(new El('th'), { ownText: '一' }));
    }
  }
  const wrapper = body.add(new El('div', ['ant-table-wrapper']));
  const antTable = wrapper.add(new El('div', ['ant-table']));
  const container = antTable.add(new El('div', ['ant-table-container']));

  const headerBox = container.add(new El('div', ['ant-table-header']));
  const headerTable = headerBox.add(new El('table'));
  const thead = headerTable.add(new El('thead', ['ant-table-thead']));
  const headRow = thead.add(new El('tr'));
  if (withCheckbox) {
    const th = headRow.add(new El('th', ['ant-table-cell']));
    th.add(new El('label')).add(new El('input', ['ant-checkbox-input']));
  }
  for (const text of headers) {
    const th = new El('th', ['ant-table-cell']);
    th.ownText = text;
    headRow.add(th);
  }

  const nCol = headers.length + (withCheckbox ? 1 : 0);
  const bodyBox = container.add(new El('div', ['ant-table-body']));
  const bodyTable = bodyBox.add(new El('table'));
  const tbody = bodyTable.add(new El('tbody', ['ant-table-tbody']));

  const measure = tbody.add(new El('tr', ['ant-table-measure-row']));
  for (let i = 0; i < nCol; i++) measure.add(new El('td', [], { style: 'padding:0' }));

  if (empty) {
    const ph = tbody.add(new El('tr', ['ant-table-placeholder']));
    const td = ph.add(new El('td', ['ant-table-cell'], { colspan: nCol }));
    const box = td.add(new El('div', ['ant-empty']));
    const p = box.add(new El('p', ['ant-empty-description']));
    p.ownText = '暂无数据';
  } else {
    for (const r of rows) {
      const row = tbody.add(new El('tr', ['ant-table-row']));
      if (withCheckbox) row.add(new El('td', ['ant-table-cell'])).add(new El('input'));
      for (const c of r) {
        const td = row.add(new El('td', ['ant-table-cell']));
        td.ownText = c;
      }
    }
  }
  return body;
}

// ---------- 从模块源码里取出被测脚本 ----------
const modulePath = path.join(__dirname, '..', 'modules', 'workorder_checker.py');
const src = fs.readFileSync(modulePath, 'utf8');
const matched = src.match(/TABLE_EXTRACT_JS = r"""([\s\S]*?)"""/);
if (!matched) {
  console.log('[FAIL] 未能在 workorder_checker.py 中找到 TABLE_EXTRACT_JS');
  process.exit(1);
}
const extract = eval('(' + matched[1].trim() + ')');

let pass = 0, fail = 0;
const check = (name, cond, detail = '') => {
  if (cond) { pass++; console.log(`  [PASS] ${name}`); }
  else { fail++; console.log(`  [FAIL] ${name} ${detail}`); }
};

// A. 标准 el-table：表头来自 header-wrapper，行来自 body-wrapper
setGlobalDom(buildElTable({
  headers: [['工单编号'], ['客户'], ['所属代理'], ['服务项目'], ['下单日期']],
  rows: [
    [['WO-1'], ['深圳市甲科技有限公司'], ['乐天'], ['德国WEEE'], ['2026-09-03 10:00:00']],
    [['WO-2'], ['深圳市乙贸易有限公司'], ['乐天'], ['德国WEEE'], ['2026-09-04 10:00:00']],
  ],
}));
let out = extract();
check('表头取自 header-wrapper/thead', out.length === 1 && out[0].headerFound === true, JSON.stringify(out));
check('表头列名正确',
  JSON.stringify(out[0].headers) ===
  JSON.stringify(['工单编号', '客户', '所属代理', '服务项目', '下单日期']),
  JSON.stringify(out[0].headers));
check('表体行全部读到',
  out[0].rows.length === 2 && out[0].rows[0][1] === '深圳市甲科技有限公司',
  JSON.stringify(out[0].rows));

// B. 固定列会多渲染一组表：同一 .el-table 只出一个候选，且取列更全的那组
setGlobalDom(buildElTable({
  headers: [['工单编号'], ['客户'], ['所属代理']],
  rows: [[['WO-1'], ['甲'], ['乐天']]],
  withFixed: true,
}));
out = extract();
check('固定列只产出一个候选且列为完整列',
  out.length === 1 && out[0].headers.length === 3 && out[0].rows.length === 1,
  JSON.stringify(out.map(o => o.headers)));

// C. 空结果：表头仍在，行数为 0，并带 empty 标记
setGlobalDom(buildElTable({ headers: [['工单编号'], ['客户']], rows: [], empty: true }));
out = extract();
check('空结果仍能读到表头',
  out.length === 1 && out[0].headerFound === true && out[0].rows.length === 0, JSON.stringify(out));
check('空结果带 empty 标记', out[0].empty === true);

// D. 旧实现的坑：表体表没有 thead 时，绝不能用第一行 td 当表头
setGlobalDom(buildElTable({
  headers: [['工单编号'], ['客户'], ['所属代理']],
  rows: [[['WO-1'], ['甲'], ['乐天']]],
}));
out = extract();
check('表头不含表体数据（不把首行当表头）',
  out[0].headers[0] === '工单编号' && out[0].rows.length === 1, JSON.stringify(out[0]));

// E. colSpan 保持列索引对齐
setGlobalDom(buildElTable({
  headers: [['合计', 2], ['服务项目']],
  rows: [[['1200'], ['德国WEEE']]],
}));
out = extract();
check('colSpan 表头补占位保持对齐',
  out[0].headers.length === 3 && out[0].headers[0] === '合计' && out[0].headers[2] === '服务项目',
  JSON.stringify(out[0].headers));

// F. 全空占位行被丢弃
setGlobalDom(buildElTable({
  headers: [['工单编号'], ['客户']],
  rows: [[[''], ['']], [['WO-1'], ['甲']]],
}));
out = extract();
check('全空行被丢弃', out[0].rows.length === 1 && out[0].rows[0][0] === 'WO-1',
  JSON.stringify(out[0].rows));

// G. 页面完全没有表格
setGlobalDom(new El('body'));
check('无表格时返回空数组', extract().length === 0);

// ============ Ant Design（工单平台真实使用的框架） ============
const ANT_HEADERS = ['状态', '下单日期', '注册类型', '主工单编号/客户号', '证书号',
  '公司名称', '所属代理/直客', '国家', '服务项目', '操作'];

// H. 有数据：表头必须来自 .ant-table-header，行来自 .ant-table-body
setGlobalDom(buildAntTable({
  headers: ANT_HEADERS,
  rows: [['待审核', '2026-09-03 10:00:00', '新注册', 'WO-1', 'C-1',
    '深圳市甲科技有限公司', '示例代理甲', '意大利', '包装法', '详情']],
}));
out = extract();
check('ant: 表头取自 ant-table-header',
  out.length === 1 && out[0].headerFound === true
  && JSON.stringify(out[0].headers) === JSON.stringify([''].concat(ANT_HEADERS)),
  JSON.stringify(out.map(o => o.headers)));
check('ant: 行取自 ant-table-body',
  out[0].rows.length === 1 && out[0].rows[0][6] === '深圳市甲科技有限公司',
  JSON.stringify(out[0].rows));

// I. 空结果（真实日志里就是这个状态）：计量行 + “暂无数据”行都不算数据行
setGlobalDom(buildAntTable({ headers: ANT_HEADERS, rows: [], empty: true }));
out = extract();
check('ant: 空结果仍是 1 个候选且有表头',
  out.length === 1 && out[0].headerFound === true, JSON.stringify(out.length));
check('ant: “暂无数据”占位行不被当成数据行',
  out[0].rows.length === 0, JSON.stringify(out[0].rows));
check('ant: 空结果带 empty 标记', out[0].empty === true);

// J. 表头表与表体表必须配对成一个候选，否则行数多的表体表会胜出（历史根因）
setGlobalDom(buildAntTable({
  headers: ANT_HEADERS,
  rows: [['待审核', '2026-09-03', '新注册', 'WO-1', 'C-1', '甲', '示例代理甲', '意大利', '包装法', '详情']],
}));
out = extract();
check('ant: 表头表/表体表不会各自成为候选',
  out.length === 1 && out[0].headerFound === true && out[0].rows.length === 1,
  JSON.stringify(out.map(o => [o.headers.length, o.rows.length, o.headerFound])));

// K. 页面上同时存在不可见的日期选择器表格时必须被过滤
setGlobalDom(buildAntTable({ headers: ANT_HEADERS, rows: [], empty: true, withPicker: true }));
out = extract();
check('ant: 不可见的日期选择器表格被过滤',
  out.length === 1 && out[0].rows.length === 0, JSON.stringify(out.length));

// L. 没有选择框列时同样工作
setGlobalDom(buildAntTable({
  headers: ANT_HEADERS, rows: [], empty: true, withCheckbox: false,
}));
out = extract();
check('ant: 无选择框列时表头列数不变',
  out.length === 1 && out[0].headers.length === ANT_HEADERS.length,
  JSON.stringify(out[0] && out[0].headers.length));

console.log(`结果: ${pass} 通过, ${fail} 失败`);
process.exit(fail ? 1 : 0);
