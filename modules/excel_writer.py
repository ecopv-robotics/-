"""M6 Excel 输出模块 — 生成双 Excel 表格（漏单清单 + 过滤清单）"""
import os
from datetime import datetime
from typing import List, Dict

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# 颜色定义
COLOR_RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
COLOR_RED_FONT = Font(color="9C0006")
COLOR_YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
COLOR_YELLOW_FONT = Font(color="9C6500")
COLOR_BLUE_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
COLOR_BLUE_FONT = Font(color="1F4E79")
COLOR_ORANGE_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
COLOR_ORANGE_FONT = Font(color="833C00")
COLOR_GRAY_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
COLOR_GRAY_FONT = Font(color="595959")
COLOR_HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
COLOR_HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)

THIN_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

# 阶段一输出 xlsx 的「表头名 → 内部字段名」映射。
# 阶段二「独立运行」是从阶段一 xlsx 反读数据行的, 那批 dict 的键是中文表头;
# 若不还原成内部键, 下游按 row.get("客户") / row.get("项目") 取值为空,
# 会导致: 去重全部塌缩成 1 条、结果表客户/标准化项目列空白。
HEADER_ALIASES = {
    "客户公司名称": "客户",
    "标准化项目名称": "项目",
    "发件人邮箱": "sender_email",
    "发件日期": "date",
    "邮件主题": "subject",
    "收件人": "recipient",
}


def normalize_header_rows(headers: List[str], data_rows: List[list]) -> List[Dict]:
    """把「表头 + 数据行」反读成的二维表转换成内部字段。

    工作台旧版导出会同时写入标准字段（如“客户公司名称”）和内部字段
    （如“客户”）。两者映射到同一个内部键时，不能让后面的空内部字段
    覆盖前面的真实值；标准表头有值时也应优先于旧内部字段的陈旧值。
    """
    header_info = [
        (HEADER_ALIASES.get(h, h), h, h in HEADER_ALIASES)
        for h in headers
    ]
    out: List[Dict] = []
    for idx, r in enumerate(data_rows):
        # 保留 Excel 读出的 datetime；阶段二要按邮件日期与工单日期做容差比较。
        d: Dict = {}
        source_priority: Dict[str, int] = {}
        for (key, _source_header, is_standard_header), value in zip(header_info, r):
            value = "" if value is None else value
            if not str(value).strip():
                # 空值只在该键尚未出现时占位，不能覆盖已经读到的非空值。
                d.setdefault(key, value)
                continue
            # 显式标准表头优先于旧版内部键；同一优先级保留先出现的值。
            priority = 0 if is_standard_header else 1
            previous_priority = source_priority.get(key)
            if key not in d or not str(d.get(key) or "").strip() or (
                previous_priority is not None and priority < previous_priority
            ):
                d[key] = value
                source_priority[key] = priority
        d.setdefault("_row_index", idx)
        out.append(d)
    return out


class ExcelWriter:
    # 默认运行输出按业务结果分类；测试/外部旧调用可继续使用 categorized=False，
    # 保持原来把文件直接写在 output_dir 根目录的兼容行为。
    CATEGORY_DIRS = {
        "stage1": "stage1_email",
        "stage2": "stage2_workorder",
        "manual": "manual_review",
        "diagnostics": "diagnostics",
    }

    def __init__(self, output_dir: str = "output", logger=None, categorized: bool = False):
        self.output_dir = output_dir
        self.logger = logger
        self.categorized = bool(categorized)
        os.makedirs(output_dir, exist_ok=True)

    def _category_path(self, category: str, filename: str) -> str:
        """返回分类输出路径；关闭分类时退回旧根目录路径。"""
        if self.categorized:
            dirname = self.CATEGORY_DIRS.get(category, category)
            folder = os.path.join(self.output_dir, dirname)
            os.makedirs(folder, exist_ok=True)
            return os.path.join(folder, filename)
        return os.path.join(self.output_dir, filename)

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    def write_missing_report(self, rows: List[Dict]) -> str:
        """生成漏单清单 Excel"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = self._category_path("stage1", f"漏单清单_{timestamp}.xlsx")

        wb = Workbook()
        ws = wb.active
        ws.title = "漏单清单"

        headers = [
            "发件人邮箱", "发件日期", "主题", "正文(精简)",
            "代理", "代理匹配方式", "客户编号", "客户", "客户提取来源",
            "项目", "项目原始值", "需求",
            "是否已录单", "工单日期", "下单日期", "匹配状态",
            "查询时间戳", "备注", "邮件日期筛选范围", "日期筛选说明",
        ]

        # 写表头
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = COLOR_HEADER_FILL
            cell.font = COLOR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        # 写数据
        for row_idx, row in enumerate(rows, 2):
            values = [
                row.get("sender_email", ""),
                self._format_date(row.get("date")),
                row.get("subject", ""),
                row.get("body_text", "")[:200],
                row.get("代理", ""),
                row.get("代理匹配方式", ""),
                row.get("客户编号", ""),
                row.get("客户", ""),
                row.get("客户提取来源", ""),
                row.get("项目", ""),
                row.get("项目原始值", ""),
                row.get("需求", ""),
                row.get("是否已录单", ""),
                # 工单日期 = 邮件发来的日期（发件日期）；平台返回的那列在「下单日期」
                self._format_date(row.get("工单日期") or row.get("date")),
                row.get("下单日期", ""),
                row.get("匹配状态", ""),
                self._format_date(row.get("查询时间戳")),
                "",
                row.get("邮件日期筛选范围", ""),
                row.get("日期筛选说明", ""),
            ]

            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)

            # 颜色标记
            self._apply_row_color(ws, row_idx, len(headers), row)

        # 调整列宽
        col_widths = [25, 20, 40, 50, 15, 18, 16, 25, 18, 18, 25, 10, 12, 20, 14, 18, 20, 15, 25, 42]
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(idx)].width = width

        ws.freeze_panes = "A2"

        wb.save(filename)
        self._log(f"漏单清单已生成: {filename}")
        return filename

    def write_filtered_report(self, filtered_mails: List[Dict]) -> str:
        """生成过滤清单 Excel"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = self._category_path("stage1", f"过滤清单_{timestamp}.xlsx")

        wb = Workbook()
        ws = wb.active
        ws.title = "过滤清单"

        headers = ["发件人邮箱", "发件日期", "主题", "过滤原因", "处理时间戳"]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = COLOR_HEADER_FILL
            cell.font = COLOR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for row_idx, mail in enumerate(filtered_mails, 2):
            values = [
                mail.get("sender_email", ""),
                self._format_date(mail.get("date")),
                mail.get("subject", ""),
                mail.get("filter_reason", ""),
                now_str,
            ]
            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        col_widths = [25, 20, 50, 20, 20]
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(idx)].width = width

        ws.freeze_panes = "A2"
        wb.save(filename)
        self._log(f"过滤清单已生成: {filename}")
        return filename

    # --------------- 文件名规范：稳定名 + 时序副本 ---------------

    @staticmethod
    def _stable_and_timestamp(stable_name: str, ext: str = ".xlsx") -> tuple:
        """返回 (稳定名路径, 时序副本路径). 副本不重复时与稳定名相同写入。"""
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        # 时序副本: {stable}_{ts}.xlsx, 与稳定名同目录
        base, suf = os.path.splitext(stable_name)
        ts_path = f"{base}_{ts}{suf}"
        return stable_name, ts_path

    def _save_pair(self, wb, stable: str, ts: str, label: str) -> str:
        """先写时序副本，再复制成稳定名，返回本次真正可用的路径。

        稳定名（workorder_check_result.xlsx / to_workorder_list.xlsx 等）经常正被
        用户在 Excel/WPS 里开着，Windows 会锁住文件：直接 wb.save(stable) 抛
        PermissionError，会让整轮结果一起丢掉（2026-09-14 实测踩过）。
        时序副本是新的时间戳文件名、不会被占用，所以先写它保证结果一定落盘；
        稳定名写不进去只降级成警告，并把稳定名替换为副本路径返回给调用方。
        """
        import shutil
        wb.save(ts)
        try:
            shutil.copyfile(ts, stable)
        except Exception as e:
            self._log(
                f"{label}稳定名写入失败（可能正被 Excel/WPS 打开）: {stable} — {e}；"
                f"本次结果保留在时序副本: {ts}",
                "warning",
            )
            return ts
        return stable

    @staticmethod
    def _add_change_tracking_sheets(wb) -> None:
        """为阶段一文件预留人工变更台账；工作台导出时会写入实际内容。"""
        modified = wb.create_sheet("修改明细")
        modified.append([
            "邮件编号", "明细编号", "修改时间", "修改字段", "修改前", "修改后", "修改原因", "操作来源",
        ])
        added = wb.create_sheet("新增明细")
        added.append([
            "邮件编号", "明细编号", "新增时间", "代理", "客户公司", "国家", "服务项目", "具体业务", "新增原因", "操作来源",
        ])
        for sheet in (modified, added):
            for cell in sheet[1]:
                cell.fill = COLOR_HEADER_FILL
                cell.font = COLOR_HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = THIN_BORDER
            for column in sheet.columns:
                sheet.column_dimensions[get_column_letter(column[0].column)].width = max(16, min(36, len(str(column[0].value or "")) + 8))
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions

    def write_stage1_outputs(self, all_rows: List[Dict], filtered_mails: List[Dict]) -> tuple:
        """
        阶段一输出三份:
          1. to_workorder_list.xlsx    (稳定名, 时序副本) — 有把握查工单的:代理命中1条 / 客户非空
          2. to_review_list.xlsx       (稳定名, 时序副本) — 需人工补全:代理未命中/候选≥2
          3. filtered_mail_record.xlsx  (稳定名, 时序副本) — 被过滤邮件审计
        返回 (to_workorder_path, to_review_path)
        """
        to_work, to_review = self._partition_rows(all_rows)

        # 1) to_workorder_list.xlsx
        wo_stable, _ = self._write_to_workorder_list(to_work)
        # 2) to_review_list.xlsx
        rv_stable, _ = self._write_to_review_list(to_review)
        # 3) filtered_mail_record.xlsx (稳定名 + 时序副本)
        fr_stable, fr_ts = self._stable_and_timestamp(
            self._category_path("stage1", "filtered_mail_record.xlsx")
        )
        self._write_filtered_mail_record(filtered_mails, fr_stable, fr_ts)

        self._log(f"[stage1] 待查清单: {wo_stable}")
        self._log(f"[stage1] 漏单复查: {rv_stable}")
        self._log(f"[stage1] 过滤日志: {fr_stable}")
        return wo_stable, rv_stable

    def write_workorder_check_result(
        self, all_rows: List[Dict], to_query: List[Dict], skipped: List[Dict] = None
    ) -> tuple:
        """
        阶段二输出 workorder_check_result.xlsx — 包含原始邮件信息 + 工单查询结果 + 跳过条目标注
        返回 (可用产物路径, 时序副本路径)；稳定名被 Excel 占用时第一个元素是时序副本。
        """
        stable, ts = self._stable_and_timestamp(
            self._category_path("stage2", "workorder_check_result.xlsx")
        )
        wb = self._do_write_workorder_check(all_rows, to_query, skipped)
        primary = self._save_pair(wb, stable, ts, "工单核对结果")
        # 漏单并不是一次性的红色标记，而是后续人工处理的独立待办。
        # 与完整工单核对结果一起输出稳定名“漏单.xlsx”，便于工作台和人员
        # 以相同口径继续处理。即使本批没有漏单，也输出仅含表头的文件，
        # 避免旧批次残留的漏单被误认为当前结果。
        missing_path = self._write_missing_workorder_report(all_rows, to_query, skipped)
        self._log(f"[stage2] 时序副本: {ts}")
        self._log(f"[stage2] 工单核对结果: {primary}")
        self._log(f"[stage2] 漏单待处理表: {missing_path}")
        return primary, ts

    def write_all_outputs(self, all_rows: List[Dict], filtered_mails: List[Dict]) -> tuple:
        """一站式输出: 沿用原 v1.0 行为 — 漏单清单 + 过滤清单 (兼容旧调用者)"""
        missing = self.write_missing_report(all_rows)
        filtered = self.write_filtered_report(filtered_mails)
        return missing, filtered

    # --------------- 私有: 行分区与各文件写入 ---------------

    @staticmethod
    def _partition_rows(rows: List[Dict]) -> tuple:
        """
        只有代理、客户、项目、需求均完整且所有质量闸门通过，才进入待查清单。
        任何缺字段、低置信、语义/业务校验异常都进入漏单复查。
        """
        to_work, to_review = [], []
        for r in rows:
            agent = (r.get("代理") or "").strip()
            confidence = (r.get("置信度") or "").strip()
            match_way = (r.get("代理匹配方式") or "").strip()
            company = (r.get("客户") or "").strip()
            project = (r.get("项目") or "").strip()
            request = (r.get("需求") or "").strip()
            semantic_status = (r.get("语义校验状态") or "").strip()
            business_status = (r.get("业务规则校验状态") or "").strip()
            if (
                not agent or not company or not project or request not in {"注册", "新增", "撤单"}
                or "多匹配" in match_way or "待确认" in match_way
                or confidence in {"low", "需人工确认"}
                or semantic_status in {"uncertain", "invalid", "LLM调用失败/未完成"}
                or business_status == "需人工确认"
            ):
                to_review.append(r)
            else:
                to_work.append(r)
        return to_work, to_review

    @staticmethod
    def _extraction_method(row: Dict) -> str:
        methods = []
        if row.get("llm_used"):
            methods.append("LLM补充")
        if row.get("ocr_used"):
            methods.append("OCR兜底")
        return " + ".join(methods) if methods else "规则提取"

    @staticmethod
    def _review_hint(row: Dict) -> str:
        """把机器判断转成可执行的人工核对提示。"""
        hints = []
        if row.get("filter_status") == "uncertain" and row.get("filter_reason"):
            hints.append(str(row["filter_reason"]))
        if not row.get("代理") or "待确认" in str(row.get("代理匹配方式", "")):
            hints.append("核对代理")
        if not row.get("客户"):
            hints.append("核对客户")
        if not row.get("项目"):
            hints.append("核对项目")
        if row.get("需求") not in {"注册", "新增", "撤单"}:
            hints.append("核对需求")
        if row.get("置信度") == "需人工确认":
            hints.append("完成后再导入工单核对")
        semantic_status = str(row.get("语义校验状态", "") or "")
        if semantic_status in {"uncertain", "invalid", "LLM调用失败/未完成"}:
            reason = str(row.get("语义校验原因", "") or "").strip()
            hints.append("LLM语义校验" + (f": {reason}" if reason else "需人工确认"))
        if row.get("业务规则校验状态") == "需人工确认":
            hints.append(str(row.get("业务规则问题", "") or "确定性业务校验需人工确认"))
        if row.get("字段LLM状态") == "失败，已转人工复核":
            hints.append(str(row.get("字段LLM失败原因", "") or "字段LLM输出未采用"))
        return "；".join(dict.fromkeys(hints))

    def _write_to_workorder_list(self, rows: List[Dict]) -> tuple:
        """to_workorder_list.xlsx — 阶段二唯一输入。稳定名 + 时序副本。"""
        stable, ts = self._stable_and_timestamp(
            self._category_path("stage1", "to_workorder_list.xlsx")
        )

        wb = Workbook()
        ws = wb.active
        ws.title = "工单待查"

        headers = [
            "发件人邮箱", "收件人", "发件日期", "邮件主题", "邮件正文摘要(最多300字)", "附件名称",
            "附件明细来源", "附件表格记录数", "附件表格输出数", "附件表格数量校验",
            "代理", "代理匹配方式",
            "客户编号", "客户公司名称", "客户提取来源",
            "标准化项目名称", "项目原始值",
            "需求", "置信度", "筛选依据", "智能提取方式", "人工复核提示", "数据来源",
            "语义校验状态", "语义问题编号", "语义问题字段", "语义问题证据", "语义当前值", "语义建议值", "语义校验建议", "语义校验原因", "语义校验模型",
            "字段LLM状态", "字段LLM失败原因", "业务规则校验状态", "业务规则问题编号", "业务规则问题",
            "附件证据", "邮件正文原文", "附件文件索引",
            "德国WEEE专项", "德国WEEE品类明细", "德国WEEE品类状态", "德国WEEE品类核对", "德国WEEE专项说明",
        ]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = COLOR_HEADER_FILL
            cell.font = COLOR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        for row_idx, row in enumerate(rows, 2):
            values = [
                row.get("sender_email", ""),
                row.get("recipient", ""),
                self._format_date(row.get("date")),
                row.get("subject", ""),
                (row.get("body_text", "") or "")[:300],
                row.get("附件名称", ""),
                row.get("附件明细来源", ""),
                row.get("附件表格记录数", ""),
                row.get("附件表格输出数", ""),
                row.get("附件表格数量校验", ""),
                row.get("代理", ""),
                row.get("代理匹配方式", ""),
                row.get("客户编号", ""),
                row.get("客户", ""),
                row.get("客户提取来源", ""),
                row.get("项目", ""),
                row.get("项目原始值", ""),
                row.get("需求", ""),
                row.get("置信度", ""),
                row.get("filter_reason", ""),
                self._extraction_method(row),
                self._review_hint(row),
                row.get("数据来源", row.get("客户提取来源", "")),  # 兼容字段
                row.get("语义校验状态", ""),
                row.get("语义问题编号", ""),
                row.get("语义问题字段", ""),
                row.get("语义问题证据", ""),
                row.get("语义当前值", ""),
                row.get("语义建议值", ""),
                row.get("语义校验建议", ""),
                row.get("语义校验原因", ""),
                row.get("语义校验模型", ""),
                row.get("字段LLM状态", ""),
                row.get("字段LLM失败原因", ""),
                row.get("业务规则校验状态", ""),
                row.get("业务规则问题编号", ""),
                row.get("业务规则问题", ""),
                row.get("附件证据", ""),
                str(
                    row.get("body_original")
                    or row.get("body_raw")
                    or row.get("body_text", "")
                    or ""
                )[:30000],
                row.get("附件文件索引", ""),
                row.get("德国WEEE专项", ""),
                row.get("德国WEEE品类明细", ""),
                row.get("德国WEEE品类状态", ""),
                row.get("德国WEEE品类核对", ""),
                row.get("德国WEEE专项说明", ""),
            ]
            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            # 颜色标记: confidence=low -> 蓝; 多匹配 -> 橙
            if row.get("置信度") == "low" or "待确认" in (row.get("代理匹配方式") or ""):
                for col in [11, 14, 16]:
                    ws.cell(row=row_idx, column=col).fill = COLOR_BLUE_FILL
                    ws.cell(row=row_idx, column=col).font = COLOR_BLUE_FONT

        col_widths = [25, 28, 20, 40, 48, 35, 42, 14, 14, 20, 18, 22, 16, 28, 22, 18, 25, 10, 12, 36, 18, 36, 22, 30, 18, 20, 18, 45, 32, 32, 30, 36, 20, 20, 36, 20, 20, 55, 55, 16, 48, 20, 54, 34]
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(idx)].width = width
        ws.freeze_panes = "A2"
        self._add_change_tracking_sheets(wb)
        stable = self._save_pair(wb, stable, ts, "工单待查清单")
        self._log(f"工单待查清单: {stable} ({len(rows)} 行)")
        return stable, ts

    def _write_to_review_list(self, rows: List[Dict]) -> tuple:
        """to_review_list.xlsx — 漏单复查: 代理未命中/候选≥2 等需人工补全。稳定名 + 时序副本。"""
        stable, ts = self._stable_and_timestamp(
            self._category_path("stage1", "to_review_list.xlsx")
        )

        wb = Workbook()
        ws = wb.active
        ws.title = "漏单复查"

        headers = [
            "发件人邮箱", "收件人", "发件日期", "邮件主题", "邮件正文摘要(最多300字)", "附件名称",
            "附件明细来源", "附件表格记录数", "附件表格输出数", "附件表格数量校验",
            "代理(可能空)", "候选代理(匹配度降序)", "代理匹配方式",
            "客户编号", "客户公司名称", "标准化项目名称", "需求",
            "置信度", "筛选依据", "智能提取方式", "人工复核提示", "数据来源", "待人工补全",
            "语义校验状态", "语义问题编号", "语义问题字段", "语义问题证据", "语义当前值", "语义建议值", "语义校验建议", "语义校验原因", "语义校验模型",
            "字段LLM状态", "字段LLM失败原因", "业务规则校验状态", "业务规则问题编号", "业务规则问题",
            "附件证据", "邮件正文原文", "附件文件索引",
            "德国WEEE专项", "德国WEEE品类明细", "德国WEEE品类状态", "德国WEEE品类核对", "德国WEEE专项说明",
        ]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = COLOR_HEADER_FILL
            cell.font = COLOR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        for row_idx, row in enumerate(rows, 2):
            # 候选代理: 多匹配时写候选列表；空时写 "未匹配"
            agent_candidates = row.get("候选代理", row.get("match_candidates", ""))
            if not agent_candidates:
                candidates_str = ""
            else:
                if isinstance(agent_candidates, list):
                    candidates_str = " | ".join(
                        f"{(c.get('代理') or (c.get('data') or {}).get('代理',''))}({c.get('score','')})"
                        if isinstance(c, dict) else str(c)
                        for c in agent_candidates
                    )
                else:
                    candidates_str = str(agent_candidates)

            values = [
                row.get("sender_email", ""),
                row.get("recipient", ""),
                self._format_date(row.get("date")),
                row.get("subject", ""),
                (row.get("body_text", "") or "")[:300],
                row.get("附件名称", ""),
                row.get("附件明细来源", ""),
                row.get("附件表格记录数", ""),
                row.get("附件表格输出数", ""),
                row.get("附件表格数量校验", ""),
                row.get("代理", ""),
                candidates_str,
                row.get("代理匹配方式", ""),
                row.get("客户编号", ""),
                row.get("客户", ""),
                row.get("项目", ""),
                row.get("需求", ""),
                row.get("置信度", ""),
                row.get("filter_reason", ""),
                self._extraction_method(row),
                self._review_hint(row),
                row.get("数据来源", row.get("客户提取来源", "")),
                "请人工补全代理后导入 to_workorder_list.xlsx",
                row.get("语义校验状态", ""),
                row.get("语义问题编号", ""),
                row.get("语义问题字段", ""),
                row.get("语义问题证据", ""),
                row.get("语义当前值", ""),
                row.get("语义建议值", ""),
                row.get("语义校验建议", ""),
                row.get("语义校验原因", ""),
                row.get("语义校验模型", ""),
                row.get("字段LLM状态", ""),
                row.get("字段LLM失败原因", ""),
                row.get("业务规则校验状态", ""),
                row.get("业务规则问题编号", ""),
                row.get("业务规则问题", ""),
                row.get("附件证据", ""),
                str(
                    row.get("body_original")
                    or row.get("body_raw")
                    or row.get("body_text", "")
                    or ""
                )[:30000],
                row.get("附件文件索引", ""),
                row.get("德国WEEE专项", ""),
                row.get("德国WEEE品类明细", ""),
                row.get("德国WEEE品类状态", ""),
                row.get("德国WEEE品类核对", ""),
                row.get("德国WEEE专项说明", ""),
            ]
            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            # 整行橙色 — 需人工复查
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).fill = COLOR_ORANGE_FILL
                ws.cell(row=row_idx, column=col).font = COLOR_ORANGE_FONT

        col_widths = [25, 28, 20, 40, 48, 35, 42, 14, 14, 20, 18, 35, 22, 16, 28, 22, 10, 12, 36, 18, 36, 22, 30, 18, 20, 18, 45, 32, 32, 30, 36, 20, 20, 36, 20, 20, 42, 55, 55, 16, 48, 20, 54, 34]
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(idx)].width = width
        ws.freeze_panes = "A2"
        self._add_change_tracking_sheets(wb)
        stable = self._save_pair(wb, stable, ts, "漏单复查清单")
        self._log(f"漏单复查清单: {stable} ({len(rows)} 行)")
        return stable, ts

    def _write_filtered_mail_record(self, filtered_mails: List[Dict],
                                     stable_path: str, ts_path: str) -> str:
        """filtered_mail_record.xlsx — 过滤日志审计。稳定名 + 时序副本。"""
        wb = Workbook()
        ws = wb.active
        ws.title = "过滤日志"

        headers = [
            "发件人邮箱", "收件人", "发件日期", "主题", "正文摘要(最多300字)",
            "附件名称", "过滤原因", "意图LLM状态", "处理时间戳", "正文原文",
        ]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = COLOR_HEADER_FILL
            cell.font = COLOR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for row_idx, mail in enumerate(filtered_mails, 2):
            values = [
                mail.get("sender_email", ""),
                mail.get("recipient", ""),
                self._format_date(mail.get("date")),
                mail.get("subject", ""),
                (mail.get("body_text", "") or "")[:300],
                "；".join(
                    str(att.get("filename", "") or "").strip()
                    for att in (mail.get("attachments") or [])
                    if isinstance(att, dict) and att.get("filename")
                ),
                mail.get("filter_reason", ""),
                mail.get("intent_llm_status", ""),
                now_str,
                str(mail.get("body_text", "") or "")[:30000],
            ]
            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        col_widths = [25, 28, 20, 50, 48, 35, 36, 22, 20, 55, 16, 48, 20, 54, 34]
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(idx)].width = width
        ws.freeze_panes = "A2"
        stable_path = self._save_pair(wb, stable_path, ts_path, "过滤日志")
        self._log(f"过滤日志: {stable_path} ({len(filtered_mails)} 行)")
        return stable_path

    def _do_write_workorder_check(self, all_rows: List[Dict], to_query: List[Dict],
                                   skipped: List[Dict] = None):
        """workorder_check_result.xlsx — 阶段二完整输出含跳过说明。

        只负责构建工作簿，保存交给调用方（_save_pair），便于统一处理
        稳定名被 Excel 占用时的降级。
        """
        wb = Workbook()
        ws = wb.active
        ws.title = "工单核对"

        headers = [
            "发件人邮箱", "发件日期", "邮件主题", "正文(精简)",
            "代理", "客户", "标准化项目", "需求", "置信度",
            "是否已录单", "工单日期", "下单日期", "匹配状态",
            "RPA查询状态", "查询方式", "模糊查询词", "跳过后说明", "查询时间戳",
            "邮件日期筛选范围", "日期筛选说明",
            "德国WEEE专项", "德国WEEE品类明细", "德国WEEE品类状态", "德国WEEE品类核对", "德国WEEE专项说明",
        ]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = COLOR_HEADER_FILL
            cell.font = COLOR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        # 已查: 写入to_query的数据 + 查询结果
        queried = {id(r): r for r in to_query}
        skipped_by_id = {id(s): s for s in (skipped or [])}

        for row_idx, row in enumerate(all_rows, 2):
            if id(row) in queried:
                q = queried[id(row)]
                values = [
                    q.get("sender_email", ""),
                    self._format_date(q.get("date")),
                    q.get("subject", ""),
                    (q.get("body_text", "") or "")[:200],
                    q.get("代理", ""),
                    q.get("客户", ""),
                    q.get("项目", ""),
                    q.get("需求", ""),
                    q.get("置信度", ""),
                    q.get("是否已录单", ""),
                    # 工单日期 = 邮件发来的日期；平台那列写在「下单日期」
                    self._format_date(q.get("工单日期") or q.get("date")),
                    q.get("下单日期", ""),
                    q.get("匹配状态", ""),
                    q.get("RPA查询状态", "已查询"),
                    q.get("查询方式", ""),
                    q.get("模糊查询词", ""),
                    q.get("_query_error", ""),
                    self._format_date(q.get("查询时间戳")),
                    q.get("邮件日期筛选范围", ""),
                    q.get("日期筛选说明", ""),
                    q.get("德国WEEE专项", ""),
                    q.get("德国WEEE品类明细", ""),
                    q.get("德国WEEE品类状态", ""),
                    q.get("德国WEEE品类核对", ""),
                    q.get("德国WEEE专项说明", ""),
                ]
            elif id(row) in skipped_by_id:
                s = skipped_by_id[id(row)]
                values = [
                    s.get("sender_email", ""),
                    self._format_date(s.get("date")),
                    s.get("subject", ""),
                    (s.get("body_text", "") or "")[:200],
                    s.get("代理", ""),
                    s.get("客户", ""),
                    s.get("项目", ""),
                    s.get("需求", ""),
                    s.get("置信度", ""),
                    "未查询",
                    self._format_date(s.get("工单日期") or s.get("date")),
                    "",
                    "前置校验跳过",
                    "跳过",
                    s.get("查询方式", "未查询"),
                    s.get("模糊查询词", ""),
                    s.get("_skip_reason", ""),
                    "",
                    s.get("邮件日期筛选范围", ""),
                    s.get("日期筛选说明", ""),
                    s.get("德国WEEE专项", ""),
                    s.get("德国WEEE品类明细", ""),
                    s.get("德国WEEE品类状态", ""),
                    s.get("德国WEEE品类核对", ""),
                    s.get("德国WEEE专项说明", ""),
                ]
            else:
                # 阶段二中途停止/异常时，保留原始待查行并明确标记未执行，
                # 不能再把整批 to_query 误报成“已查询”。
                values = [
                    row.get("sender_email", ""),
                    self._format_date(row.get("date")),
                    row.get("subject", ""),
                    (row.get("body_text", "") or "")[:200],
                    row.get("代理", ""),
                    row.get("客户", ""),
                    row.get("项目", ""),
                    row.get("需求", ""),
                    row.get("置信度", ""),
                    "未比对",
                    self._format_date(row.get("工单日期") or row.get("date")),
                    "",
                    row.get("匹配状态", "未执行"),
                    "未执行",
                    row.get("查询方式", "未执行"),
                    row.get("模糊查询词", ""),
                    "本次批次尚未执行到此条",
                    self._format_date(row.get("查询时间戳")),
                    row.get("邮件日期筛选范围", ""),
                    row.get("日期筛选说明", ""),
                    row.get("德国WEEE专项", ""),
                    row.get("德国WEEE品类明细", ""),
                    row.get("德国WEEE品类状态", ""),
                    row.get("德国WEEE品类核对", ""),
                    row.get("德国WEEE专项说明", ""),
                ]
            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)

            # 颜色: 已查询但漏单 -> 红; 前置跳过 -> 灰
            # 注: 下标按 values 位置写死；新增“查询方式/模糊查询词”后，
            # 匹配状态=12、RPA查询状态=13、查询方式=14、模糊查询词=15。
            is_missing = values[9] == "否"
            is_skipped = values[13] == "跳过"
            is_pending = values[13] == "未执行"
            if is_missing:
                for col in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=col).fill = COLOR_RED_FILL
                    ws.cell(row=row_idx, column=col).font = COLOR_RED_FONT
            elif is_skipped:
                for col in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=col).fill = COLOR_GRAY_FILL
                    ws.cell(row=row_idx, column=col).font = COLOR_GRAY_FONT
            elif is_pending:
                for col in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=col).fill = COLOR_GRAY_FILL
                    ws.cell(row=row_idx, column=col).font = COLOR_GRAY_FONT
            elif "日期异常" in str(values[12]):
                for col in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=col).fill = COLOR_YELLOW_FILL
                    ws.cell(row=row_idx, column=col).font = COLOR_YELLOW_FONT

        col_widths = [25, 20, 40, 50, 18, 28, 22, 10, 12, 12, 20, 14, 18, 14, 18, 28, 30, 20, 25, 42, 16, 48, 20, 54, 34]
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(idx)].width = width
        ws.freeze_panes = "A2"
        # 保存交给调用方（_save_pair），保证时序副本先落盘、稳定名被占用时不丢结果
        return wb

    def _write_missing_workorder_report(
        self, all_rows: List[Dict], to_query: List[Dict], skipped: List[Dict] = None,
    ) -> str:
        """输出稳定的“漏单.xlsx”，只收录已经明确判为漏单的项目。

        查询失败、前置校验跳过和尚未执行的行仍属于“未完成”，但尚不能断言为
        漏单，因此不混入该表。这样漏单处理人员不会把技术失败误当业务漏单。
        """
        stable, ts = self._stable_and_timestamp(
            self._category_path("stage2", "漏单.xlsx")
        )
        wb = Workbook()
        ws = wb.active
        ws.title = "漏单邮件"
        headers = [
            "序号", "项目编号/订单号", "邮件标题", "发件人", "收件人", "邮件日期",
            "对应业务", "客户公司", "代理", "服务项目", "漏单原因", "处理状态",
            "处理人", "处理时间", "查询时间", "附件名称",
        ]
        ws.append(headers)
        for cell in ws[1]:
            cell.fill = COLOR_HEADER_FILL
            cell.font = COLOR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        # all_rows / to_query 在正常流程中引用同一批 dict；仍去重处理，以兼容
        # 外部调用者传入副本的情形。
        candidates: List[Dict] = []
        seen = set()
        for row in list(to_query or []) + list(all_rows or []):
            marker = id(row)
            if marker in seen:
                continue
            seen.add(marker)
            candidates.append(row)

        missing_index = 0
        for row in candidates:
            found = str(row.get("是否已录单", "") or "").strip()
            match_status = str(row.get("匹配状态", "") or "").strip()
            rpa_status = str(row.get("RPA查询状态", "") or "").strip()
            is_missing = "漏单" in match_status or (
                found == "否" and not any(token in rpa_status for token in ("失败", "未查询", "待"))
            )
            if not is_missing:
                continue
            missing_index += 1
            values = [
                missing_index,
                row.get("项目编号") or row.get("订单号") or "",
                row.get("subject", ""),
                row.get("sender_email", ""),
                row.get("recipient", ""),
                self._format_date(row.get("date")),
                row.get("需求", ""),
                row.get("客户", ""),
                row.get("代理", ""),
                row.get("项目", ""),
                row.get("_query_error") or match_status or "未查到匹配工单",
                "待处理",
                "",
                "",
                self._format_date(row.get("查询时间戳")),
                row.get("附件名称", ""),
            ]
            ws.append(values)
            for cell in ws[ws.max_row]:
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.fill = COLOR_RED_FILL
                cell.font = COLOR_RED_FONT

        widths = [8, 20, 42, 28, 28, 20, 14, 30, 18, 28, 34, 14, 16, 20, 20, 38]
        for index, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(index)].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        return self._save_pair(wb, stable, ts, "漏单待处理表")

    def _apply_row_color(self, ws, row_idx: int, col_count: int, row: Dict):
        """按规则给行上色"""
        status = row.get("匹配状态", "")
        is_missing = row.get("是否已录单") == "否"
        confidence = row.get("置信度", "")
        filter_status = row.get("filter_status", "")
        has_date_anomaly = any(
            wo.get("日期异常") for wo in row.get("工单记录", [])
        )

        if filter_status == "uncertain":
            for col in range(1, col_count + 1):
                ws.cell(row=row_idx, column=col).fill = COLOR_GRAY_FILL
                ws.cell(row=row_idx, column=col).font = COLOR_GRAY_FONT
            return

        if is_missing:
            for col in range(1, col_count + 1):
                ws.cell(row=row_idx, column=col).fill = COLOR_RED_FILL
                ws.cell(row=row_idx, column=col).font = COLOR_RED_FONT
            return

        if "多匹配" in status:
            # 列号按本表表头写死：插入/删除列时必须同步（下单日期插入后匹配状态=15）
            cell = ws.cell(row=row_idx, column=15)
            cell.fill = COLOR_ORANGE_FILL
            cell.font = COLOR_ORANGE_FONT
            return

        if has_date_anomaly:
            cell = ws.cell(row=row_idx, column=15)
            cell.fill = COLOR_YELLOW_FILL
            cell.font = COLOR_YELLOW_FONT
            return

        if confidence == "low" or "待确认" in row.get("代理匹配方式", ""):
            for col in [5, 7, 9]:
                ws.cell(row=row_idx, column=col).fill = COLOR_BLUE_FILL
                ws.cell(row=row_idx, column=col).font = COLOR_BLUE_FONT

    def _format_date(self, dt) -> str:
        if dt is None:
            return ""
        if isinstance(dt, str):
            return dt
        if isinstance(dt, datetime):
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        return str(dt)
