// 运行 tests/*_browser.js（Playwright 浏览器回归）。
// 用法: NODE_PATH=<managed>/node_modules node _run_browser_tests.js [文件...]
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const ROOT = __dirname;
const DEFAULT_FILES = [
  'tests/workbench_german_desk_browser.js',
  'tests/workbench_ui03_browser.js',
  'tests/workbench_evidence_tabs_browser.js',
  'tests/workbench_calendar_browser.js',
];

(async () => {
  const files = process.argv.slice(2).length ? process.argv.slice(2) : DEFAULT_FILES;
  const browser = await chromium.launch({ headless: true });
  let failed = 0;
  for (const rel of files) {
    const abs = path.isAbsolute(rel) ? rel : path.join(ROOT, rel);
    const name = path.basename(abs);
    const context = await browser.newContext({ viewport: { width: 1600, height: 1000 } });
    const page = await context.newPage();
    try {
      // 每个用例文件导出一个 async page => {...}
      const fn = eval(fs.readFileSync(abs, 'utf8'));
      const result = await fn(page);
      console.log(`[PASS] ${name}  断言 ${result.passed} 项  ${JSON.stringify(result.checks ? result.checks.length : '')}`);
      if (result.checks) result.checks.forEach((c) => console.log(`         · ${c}`));
    } catch (error) {
      failed += 1;
      console.log(`[FAIL] ${name}`);
      console.log(`         ${String(error && error.message).split('\n')[0]}`);
    } finally {
      await context.close();
    }
  }
  await browser.close();
  console.log(failed ? `\n❌ ${failed}/${files.length} 个浏览器用例失败` : `\n✅ ${files.length}/${files.length} 个浏览器用例全部通过`);
  process.exit(failed ? 1 : 0);
})();
