import asyncio
import json
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


class FakeQueryButton:
    def __init__(self):
        self.clicks = 0

    async def click(self, **_kwargs):
        self.clicks += 1


class FakeResponseContext:
    def __init__(self, response=None, timeout_after_click=False):
        self.response = response
        self.timeout_after_click = timeout_after_click

    async def __aenter__(self):
        return self

    async def __aexit__(self, *_args):
        if self.timeout_after_click:
            raise TimeoutError("no query response")

    @property
    def value(self):
        async def result():
            return self.response
        return result()


class FakeQueryPage:
    def __init__(self, button, response=None, timeout_after_click=False):
        self.button = button
        self.response = response
        self.timeout_after_click = timeout_after_click

    async def wait_for_selector(self, _selector, **_kwargs):
        return self.button

    def expect_response(self, predicate, **_kwargs):
        if self.response is not None:
            assert predicate(self.response)
        return FakeResponseContext(self.response, self.timeout_after_click)

    async def wait_for_load_state(self, *_args, **_kwargs):
        return None


def test_query_button_click_uses_python_playwright_response_context():
    checker = make_checker()
    checker._expected_query_date_window = ("2026-07-26", "2026-09-26")
    button = FakeQueryButton()
    response = SimpleNamespace(
        status=200,
        url="https://example.invalid/orderRegister/queryPageList",
        request=SimpleNamespace(
            post_data="orderDateBegin=2026-07-26&orderDateEnd=2026-09-26"
        ),
    )
    checker._page = FakeQueryPage(button, response)
    checker._log = lambda *_args, **_kwargs: None

    assert asyncio.run(checker._click_query_button()) is True
    assert button.clicks == checker._query_click_count == 1
    assert checker._last_query_response_status == "ok"


def test_query_button_does_not_retry_when_response_times_out_after_click():
    checker = make_checker()
    button = FakeQueryButton()
    checker._page = FakeQueryPage(button, timeout_after_click=True)
    checker._log = lambda *_args, **_kwargs: None

    assert asyncio.run(checker._click_query_button()) is True
    assert button.clicks == checker._query_click_count == 1
    assert checker._last_query_response_status == "unconfirmed"


def configure_batch_runtime(checker, search_one):
    checker._cancel_requested = None
    checker.cancelled = False
    checker._cancel_notice_logged = False
    checker.last_processed_rows = []
    checker.max_consecutive_query_failures = 3
    checker.max_total_seconds = 0

    async def passthrough(awaitable, _cancelled_value):
        return await awaitable

    async def ready():
        return True

    checker._await_or_cancel = passthrough
    checker._ensure_login_session = ready
    checker._navigate_to_order_list = ready
    checker.search_one = search_one
    checker._log = lambda *_args, **_kwargs: None


class FakeRegistrationTabs:
    def __init__(self, active, events, target=None):
        self.active = active
        self.events = events
        self.target = target

    def filter(self, has_text):
        return FakeRegistrationTabs(self.active, self.events, has_text)

    async def count(self):
        return 1 if self.target else 2

    @property
    def first(self):
        return self

    def nth(self, index):
        return FakeRegistrationTabs(
            self.active, self.events, "工单列表" if index == 0 else "品类明细"
        )

    async def click(self):
        self.active["tab"] = self.target or "工单列表"
        self.events.append(("tab", self.active["tab"]))


class FakeRegistrationPage:
    def __init__(self, events, active=None):
        self.events = events
        self.active = active or {"tab": "工单列表"}

    def locator(self, _selector):
        return FakeRegistrationTabs(self.active, self.events)

    async def evaluate(self, _script):
        self.events.append(("read_table", self.active["tab"]))
        return [{
            "headers": ["品牌", "类别"],
            "rows": [["Brand-A", "第五类"]],
            "headerFound": True,
        }]


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
    result = asyncio.run(checker.search_one("示例科技有限公司", country="波兰", service_item="包装法",
                                             mail_date="2026-08-20"))

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
    result = asyncio.run(checker.search_one("示例科技有限公司", country="波兰", service_item="包装法",
                                             mail_date="2026-08-20"))

    assert result == []
    assert reads == []
    assert checker._last_query_status == "查询结果刷新失败"


def test_refresh_wait_only_confirms_changed_stable_result():
    checker = make_checker()
    checker.table_refresh_timeout_seconds = 0.05
    checker._table_loading_visible = lambda: asyncio.sleep(0, result=False)
    checker._table_fingerprint = lambda: asyncio.sleep(0, result="same old result")
    assert asyncio.run(checker._wait_for_table_refresh("same old result")) is False


def test_successful_matching_query_response_confirms_identical_stable_table():
    checker = make_checker()
    checker.table_refresh_timeout_seconds = 0.4
    checker._last_query_response_status = "ok"
    checker._last_query_response_note = "HTTP 200; date window matched"
    checker._table_loading_visible = lambda: asyncio.sleep(0, result=False)
    checker._table_fingerprint = lambda: asyncio.sleep(0, result="same result")
    assert asyncio.run(checker._wait_for_table_refresh("same result")) is True


def test_http_error_query_response_never_confirms_table_refresh():
    checker = make_checker()
    checker._last_query_response_status = "http_error"
    checker._last_query_response_note = "HTTP 500"
    checker._table_loading_visible = lambda: asyncio.sleep(0, result=False)
    checker._table_fingerprint = lambda: asyncio.sleep(0, result="new-looking result")
    assert asyncio.run(checker._wait_for_table_refresh("old result")) is False


def test_query_response_does_not_require_date_parameters():
    checker = make_checker()
    checker._expected_query_date_window = ("2026-07-26", "2026-09-26")

    class FakeRequest:
        post_data = None

    class FakeResponse:
        status = 200
        request = FakeRequest()
        url = (
            "https://work.example/api/work/orderRegister/queryPageList"
            "?orderDateBegin=2026-07-26&orderDateEnd=2026-09-26"
        )

    checker._record_query_response(FakeResponse())
    assert checker._last_query_response_status == "ok"

    FakeResponse.url = "https://work.example/api/work/orderRegister/queryPageList"
    checker._record_query_response(FakeResponse())
    assert checker._last_query_response_status == "ok"

    checker._expected_query_date_window = None
    checker._record_query_response(FakeResponse())
    assert checker._last_query_response_status == "ok"


def test_failed_date_entry_clears_previous_submitted_window():
    checker = make_checker()
    checker._expected_query_date_window = ("2026-07-26", "2026-09-26")

    async def fill_date(field_name, _value, _selector):
        return field_name == "下单开始日期"

    checker._fill_date_input = fill_date
    checker._log = lambda *_args, **_kwargs: None
    result = asyncio.run(checker._fill_order_date_window(
        ("2026-08-01", "2026-08-31")
    ))

    assert result is False
    assert checker._expected_query_date_window is None


def test_query_form_continues_when_date_controls_are_incomplete():
    checker = make_checker()

    async def no_op():
        return None

    async def filled(*_args, **_kwargs):
        return True

    async def date_not_filled(_window):
        checker._last_date_filter_status = "网页日期控件未找到，查询条件不完整"
        return False

    checker._clear_inputs = no_op
    checker._fill_company_name = filled
    checker._select_dropdown = filled
    checker._fill_order_date_window = date_not_filled
    window = checker._email_date_window("2026-08-20")
    result = asyncio.run(checker._fill_query_form(
        "示例科技有限公司", agent="代理甲", country="波兰", service_item="包装法",
        mail_date="2026-08-20", date_window=window,
    ))
    assert result is True
    assert "查询条件不完整" in checker._last_date_filter_status


def test_read_only_date_inputs_are_set_by_calendar_selection_not_text_fill():
    class FakeDateInput:
        def __init__(self, page):
            self.value = ""
            self.page = page
            self.click_count = 0

        async def evaluate(self, expression):
            if "tagName" in expression:
                return "input"
            if "readOnly" in expression:
                return True
            return self.value

        async def click(self, **_kwargs):
            self.click_count += 1
            self.page.popup_open = True

        async def fill(self, value, **_kwargs):
            raise AssertionError("只读日期框不能直接 fill")

        async def press(self, _key):
            return None

    class FakePopup:
        def __init__(self, page):
            self.page = page

        async def count(self):
            return int(self.page.popup_open)

    class FakePage:
        def __init__(self):
            self.popup_open = False
            self.targets = {}

        async def wait_for_selector(self, selector, **_kwargs):
            return self.targets[selector]

        def locator(self, _selector):
            return FakePopup(self)

        class _Keyboard:
            async def press(self, _key):
                return None

        keyboard = _Keyboard()

    async def run_case(update_value):
        checker = make_checker()
        page = FakePage()
        start = FakeDateInput(page)
        end = FakeDateInput(page)
        page.targets = {"date-start": start, "date-end": end}
        checker._page = page
        checker.selectors = {
            "order_date_from_input": ["date-start"],
            "order_date_to_input": ["date-end"],
        }
        checker._log = lambda *_args, **_kwargs: None
        selected = []

        async def select_day(date_text):
            selected.append(date_text)
            if update_value:
                if len(selected) == 1:
                    start.value = date_text
                else:
                    end.value = date_text
                    page.popup_open = False
            return True

        checker._select_calendar_day = select_day
        start_ok = await checker._fill_date_input(
            "下单开始日期", "2026-07-26", "order_date_from_input"
        )
        end_ok = await checker._fill_date_input(
            "下单结束日期", "2026-09-26", "order_date_to_input"
        )
        return start_ok, end_ok, selected, start, end

    start_ok, end_ok, selected, start, end = asyncio.run(run_case(True))
    assert (start_ok, end_ok) == (True, True)
    assert selected == ["2026-07-26", "2026-09-26"]
    assert (start.value, end.value) == ("2026-07-26", "2026-09-26")
    assert start.click_count == 1
    assert end.click_count == 0

    start_ok, end_ok, *_ = asyncio.run(run_case(False))
    assert (start_ok, end_ok) == (False, False)


def test_category_tab_reenters_dates_queries_then_restores_dates_on_list_tab():
    checker = make_checker()
    events = []
    checker._page = FakeRegistrationPage(events)
    checker.query_click_settle_seconds = 0
    date_window = checker._email_date_window("2026-08-26")

    async def fill_dates(window):
        events.append(("dates", checker._page.active["tab"], window))
        checker._last_date_filter_status = "网页已设置邮件日期前后1个月"
        return True

    async def fingerprint():
        return "table-before-category-query"

    async def click_query():
        events.append(("query", checker._page.active["tab"]))
        return True

    async def wait_refresh(before):
        events.append(("refresh", checker._page.active["tab"], before))
        return True

    checker._fill_order_date_window = fill_dates
    checker._table_fingerprint = fingerprint
    checker._click_query_button = click_query
    checker._wait_for_table_refresh = wait_refresh
    checker._rows_to_records = lambda _headers, rows: [
        {"品牌": rows[0][0], "类别": rows[0][1]}
    ]

    records = asyncio.run(checker._extract_weee_category_detail_data(date_window))

    assert records == [{"品牌": "Brand-A", "类别": "第五类"}]
    assert checker._last_category_detail_status == "loaded"
    assert events == [
        ("tab", "品类明细"),
        ("dates", "品类明细", date_window),
        ("query", "品类明细"),
        ("refresh", "品类明细", "table-before-category-query"),
        ("read_table", "品类明细"),
        ("tab", "工单列表"),
        ("dates", "工单列表", date_window),
    ]


def test_category_tab_does_not_read_old_table_when_requery_refresh_fails():
    checker = make_checker()
    events = []
    checker._page = FakeRegistrationPage(events)
    checker.query_click_settle_seconds = 0
    date_window = checker._email_date_window("2026-08-26")

    async def fill_dates(window):
        events.append(("dates", checker._page.active["tab"], window))
        checker._last_date_filter_status = "网页已设置邮件日期前后1个月"
        return True

    async def fingerprint():
        return "old category table"

    async def click_query():
        events.append(("query", checker._page.active["tab"]))
        return True

    async def no_refresh(_before):
        events.append(("refresh_failed", checker._page.active["tab"]))
        return False

    checker._fill_order_date_window = fill_dates
    checker._table_fingerprint = fingerprint
    checker._click_query_button = click_query
    checker._wait_for_table_refresh = no_refresh
    checker._rows_to_records = lambda *_args: [
        {"品牌": "STALE", "类别": "第五类"}
    ]

    records = asyncio.run(checker._extract_weee_category_detail_data(date_window))

    assert records == []
    assert checker._last_category_detail_status == "failed"
    assert not any(event[0] == "read_table" for event in events)
    assert events[-2:] == [
        ("tab", "工单列表"),
        ("dates", "工单列表", date_window),
    ]


def test_search_one_returns_to_list_and_refills_date_before_query():
    checker = make_checker()
    events = []
    checker._page = FakeRegistrationPage(events, active={"tab": "品类明细"})
    checker._debug_screenshot = lambda *_args, **_kwargs: asyncio.sleep(0)

    async def fill(**kwargs):
        events.append(("fill", checker._page.active["tab"], kwargs.get("mail_date")))
        return True

    async def fingerprint():
        return "before-query"

    async def click_query():
        events.append(("query", checker._page.active["tab"]))
        return True

    async def wait_refresh(_before):
        return True

    async def extract():
        checker._last_table_parse_status = "ok"
        return [{"工单编号": "WO-1", "下单日期": "2026-08-26"}]

    checker._fill_query_form = fill
    checker._table_fingerprint = fingerprint
    checker._click_query_button = click_query
    checker._wait_for_table_refresh = wait_refresh
    checker._extract_table_data = extract

    result = asyncio.run(checker.search_one(
        "示例科技有限公司", country="波兰", service_item="包装法",
        mail_date="2026-08-26",
    ))

    assert len(result) == 1
    assert events[:3] == [
        ("tab", "工单列表"),
        ("fill", "工单列表", "2026-08-26"),
        ("query", "工单列表"),
    ]


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

    result = asyncio.run(checker.search_one("示例科技有限公司", country="波兰", service_item="包装法",
                                             mail_date="2026-08-20"))

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
        return [{"工单编号": "WO-1", "下单日期": "2026-08-20"}]

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
        {"客户": "示例科技有限公司", "代理": "代理甲", "项目": "波兰包装法", "date": "2026-08-20"},
        {"客户": "样例贸易有限公司", "代理": "代理甲", "项目": "波兰包装法", "date": "2026-08-20"},
        {"客户": "缺少日期有限公司", "代理": "代理甲", "项目": "波兰包装法"},
    ]

    result = asyncio.run(checker.search_batch(rows))

    assert len(result) == 3
    assert observed_live_flags == [True, True]
    assert checker._query_click_count == 2
    assert checker.force_live_query is False
    assert result[2]["RPA查询状态"] == "查询条件不完整-未执行"


def test_date_window_keeps_inclusive_boundaries_and_unknowns_but_excludes_outside():
    checker = make_checker()
    window = checker._email_date_window("2026-08-26")
    orders = [
        {"工单编号": "START", "下单日期": "2026-07-26"},
        {"工单编号": "END", "下单日期": "2026-09-26"},
        {"工单编号": "BEFORE", "下单日期": "2026-07-25"},
        {"工单编号": "AFTER", "下单日期": "2026-09-27"},
        {"工单编号": "EMPTY", "下单日期": ""},
        {"工单编号": "UNPARSEABLE", "下单日期": "not-a-date"},
    ]

    candidates, outside, unknown = checker._filter_workorders_by_date_window(orders, window)

    assert [order["工单编号"] for order in candidates] == [
        "START", "END", "EMPTY", "UNPARSEABLE"
    ]
    assert [order["工单编号"] for order in outside] == ["BEFORE", "AFTER"]
    assert [order["工单编号"] for order in unknown] == ["EMPTY", "UNPARSEABLE"]


def test_search_batch_excludes_out_of_window_orders_before_unique_match():
    checker = make_checker()

    async def search_one(**_kwargs):
        checker._last_query_status = "实时查询"
        checker._last_lookup_mode = "精确查询"
        checker._last_fuzzy_term = ""
        checker._last_date_filter_status = "网页日期已设置"
        return [
            {
                "工单编号": "IN-WINDOW",
                "下单日期": "2026-08-26",
                "代理": "代理甲",
                "公司名称": "示例科技有限公司",
                "项目": "波兰包装法",
            },
            {
                "工单编号": "OUT-OF-WINDOW",
                "下单日期": "2026-03-19",
                "代理": "代理甲",
                "公司名称": "示例科技有限公司",
                "项目": "波兰包装法",
            },
        ]

    configure_batch_runtime(checker, search_one)
    rows = [{
        "客户": "示例科技有限公司",
        "代理": "代理甲",
        "项目": "波兰包装法",
        "date": "2026-08-26",
    }]

    result = asyncio.run(checker.search_batch(rows))
    row = result[0]

    assert row["是否已录单"] == "是"
    assert row["匹配状态"] == "精确匹配"
    assert [order["下单日期"] for order in row["工单记录"]] == ["2026-08-26"]
    assert "窗口内 1 条" in row["日期筛选说明"]
    assert "窗口外剔除 1 条" in row["日期筛选说明"]


def test_search_batch_marks_unknown_order_date_for_review_and_blocks_german_success():
    checker = make_checker()

    async def search_one(**_kwargs):
        checker._last_query_status = "实时查询"
        checker._last_lookup_mode = "精确查询"
        checker._last_fuzzy_term = ""
        checker._last_date_filter_status = "网页日期已设置"
        return [{
            "工单编号": "UNKNOWN-DATE",
            "下单日期": "unknown",
            "代理": "代理甲",
            "公司名称": "示例科技有限公司",
            "项目": "德国WEEE",
            "品类明细": [{"品牌": "Brand-A", "类别": "第五类"}],
            "德国WEEE品类明细读取状态": "loaded",
        }]

    configure_batch_runtime(checker, search_one)
    rows = [{
        "客户": "示例科技有限公司",
        "代理": "代理甲",
        "项目": "德国WEEE",
        "date": "2026-08-26",
        "德国WEEE专项": "是",
        "德国WEEE品类明细": [{"brand": "Brand-A", "category": "第五类"}],
    }]

    row = asyncio.run(checker.search_batch(rows))[0]

    assert row["是否已录单"] == "待复核"
    assert row["匹配状态"] == "下单日期待复核"
    audit = json.loads(row["德国WEEE品类核对"])
    assert audit["status"] == "pending"
    success, reason = ExcelWriter._workorder_success_and_reason(row, "queried")
    assert success == "否"
    assert "日期缺失或无法解析" in reason


def test_german_detail_stays_pending_when_raw_multi_match_cannot_be_bound_after_date_filter():
    checker = make_checker()

    async def search_one(**_kwargs):
        checker._last_query_status = "实时查询"
        checker._last_lookup_mode = "精确查询"
        checker._last_fuzzy_term = ""
        checker._last_date_filter_status = "网页日期已设置"
        return [
            {
                "工单编号": "IN-WINDOW",
                "下单日期": "2026-08-26",
                "代理": "代理甲",
                "公司名称": "示例科技有限公司",
                "项目": "德国WEEE",
                "品类明细": [{"品牌": "Brand-A", "类别": "第五类"}],
                "德国WEEE品类明细读取状态": "loaded",
            },
            {
                "工单编号": "OUT-OF-WINDOW",
                "下单日期": "2026-03-19",
                "代理": "代理甲",
                "公司名称": "示例科技有限公司",
                "项目": "德国WEEE",
            },
        ]

    configure_batch_runtime(checker, search_one)
    rows = [{
        "客户": "示例科技有限公司",
        "代理": "代理甲",
        "项目": "德国WEEE",
        "date": "2026-08-26",
        "德国WEEE专项": "是",
        "德国WEEE品类明细": [{"brand": "Brand-A", "category": "第五类"}],
    }]

    row = asyncio.run(checker.search_batch(rows))[0]
    audit = json.loads(row["德国WEEE品类核对"])
    success, reason = ExcelWriter._workorder_success_and_reason(row, "queried")

    assert row["是否已录单"] == "待复核"
    assert row["匹配状态"] == "品类明细关联待复核"
    assert audit["status"] == "pending"
    assert "无法安全关联" in reason
    assert success == "否"


def test_search_batch_fails_when_all_live_candidates_are_outside_date_window():
    checker = make_checker()

    async def search_one(**_kwargs):
        checker._last_query_status = "实时查询"
        checker._last_lookup_mode = "精确查询"
        checker._last_fuzzy_term = ""
        checker._last_date_filter_status = "网页日期已设置"
        return [{
            "工单编号": "OLD",
            "下单日期": "2026-03-19",
            "代理": "代理甲",
            "公司名称": "示例科技有限公司",
            "项目": "德国WEEE",
        }]

    configure_batch_runtime(checker, search_one)
    rows = [{
        "客户": "示例科技有限公司",
        "代理": "代理甲",
        "项目": "德国WEEE",
        "date": "2026-08-26",
        "德国WEEE专项": "是",
        "德国WEEE品类明细": [{"brand": "Brand-A", "category": "第五类"}],
    }]

    row = asyncio.run(checker.search_batch(rows))[0]
    success, reason = ExcelWriter._workorder_success_and_reason(row, "queried")

    assert row["是否已录单"] == "否"
    assert row["匹配状态"] == "日期范围内无工单"
    assert success == "否"
    assert "下单日期均超出有效窗口" in reason


def test_german_workorder_export_requires_completed_special_audit():
    base = {
        "项目": "德国WEEE",
        "德国WEEE专项": "是",
        "是否已录单": "是",
        "匹配状态": "精确匹配",
        "RPA查询状态": "实时查询",
    }
    pending = {**base, "德国WEEE品类核对": '{"status":"pending","reason":"字段不清楚",'
                                               '"extra_items":[{"brand":"Other","category":"第5类"}]}'}
    pending_success, pending_reason = ExcelWriter._workorder_success_and_reason(pending, "queried")
    assert pending_success == "否"
    assert "字段不清楚" in pending_reason and "额外品类" in pending_reason
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
