import asyncio
from types import SimpleNamespace

from modules.excel_writer import ExcelWriter
from modules.workorder_checker import WorkOrderChecker


def make_checker():
    checker = WorkOrderChecker({
        "url": "http://localhost",
        "username": "",
        "password": "",
        "query_cache_path": None,
        "input_step_delay_seconds": 0,
        "query_click_settle_seconds": 0,
        "query_interval_seconds": 0,
    })
    checker._logged_in = True
    return checker


def test_exact_click_failure_does_not_read_previous_table():
    checker = make_checker()
    reads = []

    async def fill(**_kwargs):
        return True

    async def fingerprint():
        return "old table"

    async def click():
        return False

    async def extract():
        reads.append(True)
        return [{"工单编号": "STALE"}]

    checker._fill_query_form = fill
    checker._table_fingerprint = fingerprint
    checker._debug_screenshot = lambda *_args, **_kwargs: asyncio.sleep(0)
    checker._click_query_button = click
    checker._extract_table_data = extract
    result = asyncio.run(checker.search_one("示例科技有限公司", country="波兰", service_item="包装法"))

    assert result == []
    assert reads == []
    assert checker._last_query_status == "查询按钮失败"


def test_refresh_timeout_does_not_read_previous_table():
    checker = make_checker()
    reads = []

    async def fill(**_kwargs):
        return True

    async def fingerprint():
        return "same old table"

    async def click():
        return True

    async def no_refresh(_before):
        return False

    async def extract():
        reads.append(True)
        return [{"工单编号": "STALE"}]

    checker._fill_query_form = fill
    checker._table_fingerprint = fingerprint
    checker._debug_screenshot = lambda *_args, **_kwargs: asyncio.sleep(0)
    checker._click_query_button = click
    checker._wait_for_table_refresh = no_refresh
    checker._extract_table_data = extract
    result = asyncio.run(checker.search_one("示例科技有限公司", country="波兰", service_item="包装法"))

    assert result == []
    assert reads == []
    assert checker._last_query_status == "查询结果刷新失败"


def test_fuzzy_click_failure_does_not_read_exact_query_table_again():
    checker = make_checker()
    click_results = iter([True, False])
    reads = []

    class FakeInput:
        value = ""

        async def is_visible(self):
            return True

        async def fill(self, value):
            self.value = value

        async def input_value(self):
            return self.value

    fake_input = FakeInput()

    async def query_selector_all(_selector):
        return [fake_input]

    async def fill(**_kwargs):
        return True

    async def fingerprint():
        return "same fingerprint"

    async def click():
        return next(click_results)

    async def refreshed(_before):
        return True

    async def extract():
        reads.append(True)
        checker._last_table_parse_status = "ok"
        return []

    checker._page = SimpleNamespace(query_selector_all=query_selector_all)
    checker._fill_query_form = fill
    checker._table_fingerprint = fingerprint
    checker._debug_screenshot = lambda *_args, **_kwargs: asyncio.sleep(0)
    checker._click_query_button = click
    checker._wait_for_table_refresh = refreshed
    checker._extract_table_data = extract
    checker._get_fuzzy_company_terms = lambda _company: ["示例科技有限公司", "示例科技"]

    result = asyncio.run(checker.search_one("示例科技有限公司", country="波兰", service_item="包装法"))

    assert result == []
    assert len(reads) == 1  # Only the exact query was read; fuzzy failure must not reuse it.
    assert checker._last_query_status == "查询按钮失败"


def test_search_batch_forces_live_query_for_each_valid_row_even_if_disabled_in_config():
    checker = make_checker()
    checker.force_live_query = False
    checker._cancel_requested = None
    checker.cancelled = False
    checker._cancel_notice_logged = False
    checker.last_processed_rows = []
    checker.max_consecutive_query_failures = 3
    checker.max_total_seconds = 0
    observed_live_flags = []

    async def passthrough(awaitable, _cancelled_value):
        return await awaitable

    async def ready():
        return True

    async def search_one(**_kwargs):
        observed_live_flags.append(checker.force_live_query)
        checker._query_click_count += int(checker.force_live_query)
        checker._last_query_status = "实时查询"
        checker._last_lookup_mode = "精确查询"
        checker._last_fuzzy_term = ""
        checker._last_date_filter_status = "日期已设置"
        return [{"工单编号": "WO-1"}]

    def match(rows, _orders):
        rows[0]["是否已录单"] = "是"
        rows[0]["匹配状态"] = "精确匹配"

    checker._await_or_cancel = passthrough
    checker._ensure_login_session = ready
    checker._navigate_to_order_list = ready
    checker.search_one = search_one
    checker.match_records = match
    checker._log = lambda *_args, **_kwargs: None
    rows = [
        {"客户": "示例科技有限公司", "代理": "代理甲", "项目": "波兰包装法"},
        {"客户": "样例贸易有限公司", "代理": "代理甲", "项目": "波兰包装法"},
    ]

    result = asyncio.run(checker.search_batch(rows))

    assert len(result) == 2
    assert observed_live_flags == [True, True]
    assert checker._query_click_count == 2
    assert checker.force_live_query is False


def test_german_workorder_export_requires_completed_special_audit():
    base = {
        "项目": "德国WEEE",
        "德国WEEE专项": "是",
        "是否已录单": "是",
        "匹配状态": "精确匹配",
        "RPA查询状态": "实时查询",
    }
    pending = {**base, "德国WEEE品类核对": '{"status":"pending","reason":"字段不清楚"}'}
    assert ExcelWriter._workorder_success_and_reason(pending, "queried") == (
        "否", "德国WEEE品类核验待复核：字段不清楚"
    )
    missing = {**base, "德国WEEE品类核对": '{"status":"missing","reason":"类别缺失"}'}
    assert ExcelWriter._workorder_success_and_reason(missing, "queried") == (
        "否", "德国WEEE品类核验缺项/不符：类别缺失"
    )
    matched = {**base, "德国WEEE品类核对": '{"status":"matched","reason":"已匹配",'
                                            '"extra_items":[{"brand":"Other","category":"第5类"}]}'}
    success, reason = ExcelWriter._workorder_success_and_reason(matched, "queried")
    assert success == "是"
    assert "额外品类" in reason and "Other" in reason


def test_other_project_export_does_not_require_german_category_audit():
    row = {"项目": "波兰包装法", "是否已录单": "是", "匹配状态": "精确匹配",
           "RPA查询状态": "实时查询"}
    assert ExcelWriter._workorder_success_and_reason(row, "queried") == ("是", "")


def test_non_german_battery_project_uses_workorder_list_only():
    row = {"项目": "法国电池法", "国家": "法国", "是否已录单": "是",
           "匹配状态": "精确匹配", "RPA查询状态": "实时查询"}
    assert ExcelWriter._workorder_success_and_reason(row, "queried") == ("是", "")
