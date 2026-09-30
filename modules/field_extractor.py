"""M3 字段提取模块 — 表驱动匹配为主, LLM 补充提取为辅

流程:
  E1 规则提取: 需求(正则) → 代理(邮箱查表) → 客户(主题→正文→附件) → 项目(附件四扫描)
  E2 LLM 补充: 对规则提取后仍有空字段、或多公司关系待确认的邮件, 调 LLM 补充提取
     - 仅对缺失字段做补充, 不覆盖规则已提取的结果
     - LLM 提取的项目需与附件四做二次匹配标准化
     - LLM 失败 → 保持规则结果, 标记"待确认"
"""
import json
import html
import re
import os
from typing import List, Dict, Set, Optional

from utils.fuzzy_match import fuzzy_search, fuzzy_match_pair, normalize_text
from modules.project_normalizer import COUNTRIES
from modules.weee_category_audit import extract_weee_items


class FieldExtractor:
    def __init__(
        self,
        agent_email_map: Dict[str, dict],
        project_names: List[dict],
        logger=None,
        llm_client=None,
        ocr_fallback: bool = True,
    ):
        """
        agent_email_map: {邮箱地址(lower): {"代理": str, "代理简称": str, "收件人邮箱": str}}
        project_names: [{"项目编号": str, "项目名称": str, "国家": str, "业务类型": str}, ...]
        llm_client: 可选的 LLM 客户端, 用于补充提取
        ocr_fallback: 是否启用 OCR 图片兜底（常规源+LLM 之后仍缺字段才触发）
        """
        self.agent_email_map = agent_email_map
        self.project_names = project_names
        self.project_name_list = [p["项目名称"] for p in project_names if p.get("项目名称")]
        self.project_candidates = [(p["项目名称"], p) for p in project_names if p.get("项目名称")]
        # 邮箱并不是代理唯一标识。真实主题常以“巨齿鲨+客户+项目”或
        # “示例代理己——客户编号+公司”开头，因此预先把附件三中的全称、简称、
        # 别名建成可审计的主题匹配索引。
        self._agent_alias_index = self._build_agent_alias_index()
        self.logger = logger
        self.llm_client = llm_client
        self.ocr_fallback = ocr_fallback

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    def extract_fields(self, mail: dict) -> List[Dict]:
        """
        从单封邮件提取字段
        返回: [{代理, 客户, 项目原始值, 项目(标准化), 需求, 代理匹配方式,
                客户提取来源, 置信度, match_candidates, ...}]
        """
        subject = mail.get("subject", "") or ""
        body = mail.get("body_text", "") or ""
        attachments = mail.get("attachments", [])
        sender_email = mail.get("sender_email", "").lower()

        # E1a: 提取需求类型
        need = self._extract_need(subject + " " + body)

        # E1a+: 结构化「公司+项目」清单行 (代理群发多公司邮件, 主题/正文逐行声明)
        groups = self._structured_groups(subject, body)

        # 业务约定：正文是代理/客户实际填写的申请内容，主题可能只是复制标题
        # 或存在笔误。因此当正文也明确列出公司时，丢弃仅来自主题的同名候选，
        # 让后续展开优先使用正文中的客户名称。
        body_company_names = {
            normalize_text(name) for name in self._company_substrings(body)
        }
        if body_company_names and groups:
            body_groups = [
                group for group in groups
                if normalize_text(group.get("customer", "")) in body_company_names
            ]
            # 编号批量清单允许品牌名没有“公司/有限公司”后缀；这类组已经
            # 明确标记为“正文批量清单”，不能因正文候选只抓到其中的法定公司
            # 就把品牌客户删掉。
            has_numbered_bulk = any(
                str(group.get("source", "")).startswith("正文批量清单")
                for group in groups
            )
            if body_groups and not has_numbered_bulk:
                groups = body_groups

        # E1b: 提取代理
        agent_result = self._extract_agent(sender_email, subject, body, attachments)

        # E1c: 提取客户。明确的 EPR 申请表是客户提交的注册资料，
        # 优先级高于正文/主题；没有申请表时才回落到原有正文→主题→附件。
        epr_customer = self._extract_customer_from_epr_application(attachments)
        customer_from_epr_form = bool(epr_customer.get("customer") or epr_customer.get("company_candidates"))
        if epr_customer and (not groups or len(groups) == 1):
            if groups and len(groups) == 1:
                groups[0]["customer"] = epr_customer["customer"]
                groups[0]["source"] = epr_customer["source"]
                groups[0]["company_candidates"] = epr_customer.get("company_candidates", [])
            customer_result = epr_customer
        elif groups and groups[0].get("customer"):
            customer_result = {
                "customer": groups[0]["customer"],
                "source": groups[0]["source"],
            }
        else:
            customer_result = self._extract_customer(subject, body, attachments)
        customer_result = self._sanitize_customer_result(customer_result)

        # 公司名可能同时出现在主题、正文、附件文件名和附件表格中。单一来源
        # 抽到的值不应直接覆盖其它来源；先做一次跨来源印证，只有至少两个
        # 独立来源支持同一主体时才替换一个明显冲突/缺失的候选。多公司结构化
        # 邮件仍由后面的 groups 保留逐行关系，不在这里压成一个公司。
        text_company_candidates = {
            normalize_text(name)
            for name in (
                self._company_substrings(subject)
                + self._company_substrings(body)
            )
            if name
        }
        # 正文/主题已经明确出现多个主体时，不能把跨来源排序结果压成一个
        # customer_result；后面的结构化分组/LLM 负责保留公司—项目关系。
        if not customer_from_epr_form and len(groups) <= 1 and len(text_company_candidates) <= 1:
            consensus = self._cross_source_company_consensus(
                subject, body, attachments
            )
            current_customer = str(customer_result.get("customer", "") or "").strip()
            current_norm = normalize_text(current_customer)
            consensus_norm = normalize_text(consensus.get("customer", "")) if consensus else ""
            if consensus and consensus_norm and (
                not current_norm or current_norm != consensus_norm
            ):
                customer_result["customer"] = consensus["customer"]
                customer_result["source"] = consensus["source"]
                self._log(
                    f"公司名跨来源印证: {current_customer or '空'} → "
                    f"{consensus['customer']} ({consensus['source']})"
                )
        customer_result["customer_code"] = self._client_code_near_company(
            body, customer_result.get("customer", "")
        ) or self._client_code_near_company(subject, customer_result.get("customer", ""))

        # E1d: 提取项目
        project_result = self._extract_projects(subject, body, attachments, groups)

        # 结构化 Excel 的每一行是业务记录，不是普通的“附件文本”。正文可能只
        # 罗列其中一部分（例如正文 15 家、xlsx 18 家），此时必须以附件表的
        # 实际行数展开，不允许正文命中后把附件剩余三行静默丢掉。勾选式 EPR
        # 表仍以勾选结果决定项目；但若同一附件中存在客户明细表，客户数量及
        # 每行对应关系必须以该明细表为准，不能因有勾选表而丢失客户行。
        attachment_groups = self._attachment_structured_groups(
            attachments, project_result
        )
        if attachment_groups and not project_result.get("epr_form_groups"):
            project_result = self._prefer_attachment_record_groups(
                project_result, attachment_groups
            )

        # 有些代理邮件正文只写“附件为两家公司”，公司名分别写在多个压缩包文件名中。
        # 若只取第一个附件公司，会静默漏掉后续客户；在主题/正文没有结构化客户
        # 关系、且项目已明确时，把每个不同的附件公司展开成独立明细。
        if (not groups and project_result.get("epr_form") is None
                and not project_result.get("epr_form_groups")):
            attachment_companies = self._attachment_company_names(attachments)
            # 标题已经给出多个完整主体时，优先使用标题的完整名称。附件文件名
            # 经常为了简短省略“东莞市虎门”等行政前缀，不能用截短的文件名覆盖
            # 标题里的法定/个体工商户名称。
            subject_companies = self._company_substrings(subject)
            if len(subject_companies) >= 2:
                attachment_companies = subject_companies
            attachment_projects = [
                p.get("standard_name") for p in project_result.get("projects") or []
                if p.get("standard_name")
            ]
            if len(attachment_companies) >= 2 and attachment_projects:
                groups = [
                    {
                        "customer": company,
                        "source": "附件文件名结构化",
                        "projects": list(attachment_projects),
                        "needs_review": False,
                    }
                    for company in attachment_companies
                ]
                project_result["groups"] = groups
                self._log(
                    f"附件文件名识别多家公司: {attachment_companies}，"
                    f"项目={attachment_projects}"
                )

        # 规则/附件/LLM 的结构化分组都必须经过同一家公司主体闸门。旧路径
        # 只清洗了普通 customer_result，直接把 groups[i]["customer"] 写入输出，
        # 因而“一家公司”这类占位词可能绕过复检；无法证明为公司时保留空值和
        # 人工复核标记，绝不把占位词写入业务明细。
        if project_result.get("groups"):
            project_result["groups"] = self._sanitize_customer_groups(
                project_result.get("groups") or []
            )

        # E2: LLM 补充提取 — 仅对缺失字段
        llm_used = False
        llm_attempted = False
        llm_failure_reason = ""
        missing_agent = not agent_result.get("agent") or agent_result.get("match_method") == "待确认"
        complete_group_customers = bool(project_result.get("groups")) and all(
            group.get("customer") for group in project_result.get("groups") or []
        )
        missing_customer = (
            not customer_result.get("customer") or customer_result.get("source") == "待确认"
        ) and not complete_group_customers
        # 有 EPR 申请表勾选结果时, 项目以勾选为唯一权威 —— 不走 LLM / 规则兜底
        has_epr_form = project_result.get("epr_form") is not None
        missing_project = not project_result.get("projects") and not has_epr_form
        missing_need = not need
        coverage_missing = self._project_coverage_gap(subject, body, project_result)
        if coverage_missing:
            project_result["need_review"] = True
            project_result["coverage_missing"] = coverage_missing
            self._log(
                f"项目完整性校验发现漏项: {coverage_missing}；转字段复检",
                "warning",
            )

        # 多公司邮件是规则层最容易漏掉的场景：规则可能已经找到了一个客户、
        # 一组项目，但没有建立公司—项目关系。这时字段表面上并不为空，旧逻辑
        # 因而不会调用 LLM，最终第二家公司会被静默吞掉。仅对“至少两家公司且
        # 结构化分组缺失/待确认”的邮件启用记录级补充，避免每封普通邮件都增加
        # 调用成本；有勾选式申请表时仍以勾选结果为唯一权威。
        all_text = subject + "\n" + body
        entity_companies = self._company_substrings(all_text)
        structured_groups = project_result.get("groups") or groups or []
        grouped_companies = {
            normalize_text(g.get("customer", "")) for g in structured_groups
            if g.get("customer")
        }
        missing_entities = {
            normalize_text(name) for name in entity_companies
        } - grouped_companies
        multi_entity_hint = (
            not has_epr_form
            and not project_result.get("epr_form_groups")
            and len(entity_companies) >= 2
            and (
                not structured_groups
                or bool(missing_entities)
                or any(g.get("needs_review") for g in structured_groups)
            )
        )

        # 代理缺失通常只是代理邮箱表不完整，单独为它调用 LLM 成本高且不可靠。
        # 仅当客户、项目或业务动作确实缺失时才调；若本次已经要调，仍允许 LLM
        # 顺带补全代理，不覆盖规则结果。
        requires_llm = (
            missing_customer or missing_project or missing_need or multi_entity_hint
            or bool(coverage_missing)
        )
        if self.llm_client and self.llm_client.enabled and requires_llm:
            llm_attempted = True
            self._log(f"规则提取有缺失, 调 LLM 补充: {subject[:40]}")
            llm_result = self.llm_client.extract_fields_llm(subject, body, attachments, sender_email)
            if llm_result:
                llm_used = True
                # 仅补充缺失字段, 不覆盖已有结果
                if missing_agent and llm_result.get("代理"):
                    agent_result["agent"] = llm_result["代理"]
                    agent_result["match_method"] = "LLM补充"
                if missing_customer and llm_result.get("客户"):
                    llm_customer = self._accept_llm_customer(llm_result.get("客户"))
                    if llm_customer:
                        customer_result["customer"] = llm_customer
                        customer_result["source"] = "LLM补充"
                    else:
                        # 模型无法确认时保持空值；后续业务校验会把该行送人工，
                        # 绝不把职位、注册资本、说明句塞进客户列。
                        self._log("字段Agent客户候选未通过公司主体闸门，保持空值", "warning")
                if not customer_result.get("customer_code") and llm_result.get("客户编号"):
                    customer_result["customer_code"] = str(llm_result["客户编号"]).strip()
                if missing_need and llm_result.get("需求"):
                    need = llm_result["需求"]
                if (missing_project or coverage_missing) and llm_result.get("项目"):
                    llm_projects = llm_result["项目"]
                    llm_project_result = self._match_llm_projects(llm_projects)
                    if missing_project:
                        # 附件结构化表已经建立了“每行一个主体”的关系时，
                        # LLM 只是在补充共同的项目，不能用新的 project_result
                        # 把这些行分组丢掉。旧写法会把两家公司重新压成一条，
                        # 尤其影响“德国项目每日登记表”这类只有“种类”列、
                        # 项目从邮件上下文补全的附件。
                        attachment_groups_before_llm = list(
                            project_result.get("groups") or []
                        )
                        attachment_count_before_llm = project_result.get(
                            "attachment_record_count"
                        )
                        project_result = llm_project_result
                        if attachment_groups_before_llm:
                            project_result["groups"] = attachment_groups_before_llm
                            if attachment_count_before_llm:
                                project_result["attachment_record_count"] = (
                                    attachment_count_before_llm
                                )
                            llm_project_names = [
                                str(item.get("standard_name") or "").strip()
                                for item in llm_project_result.get("projects") or []
                                if str(item.get("standard_name") or "").strip()
                            ]
                            # 只有一个共同项目时才安全继承到每条附件记录；
                            # 多项目无法从“种类”列区分，必须保留人工核对，
                            # 不能制造公司×项目笛卡尔积。
                            if len(llm_project_names) == 1:
                                for group in attachment_groups_before_llm:
                                    if not group.get("projects"):
                                        group["projects"] = list(llm_project_names)
                            project_result["need_review"] = bool(
                                project_result.get("need_review")
                            ) or any(
                                group.get("needs_review")
                                for group in attachment_groups_before_llm
                            )
                    else:
                        # 复检只能补漏，不能覆盖规则已经确认的项目。
                        existing = list(project_result.get("projects") or [])
                        known = {p.get("standard_name") for p in existing}
                        for item in llm_project_result.get("projects") or []:
                            if item.get("standard_name") not in known:
                                existing.append(item)
                                known.add(item.get("standard_name"))
                        project_result["projects"] = existing
                        groups_for_merge = project_result.get("groups") or []
                        if len(groups_for_merge) == 1:
                            group = groups_for_merge[0]
                            group["projects"] = list(dict.fromkeys(
                                list(group.get("projects") or [])
                                + [p.get("standard_name") for p in existing if p.get("standard_name")]
                            ))

                # 新协议返回“记录”数组。只有多实体提示成立时才用它修复公司—项目
                # 关系；普通邮件仍完全沿用规则结果，防止 LLM 改写已确认字段。
                llm_groups = self._llm_record_groups(llm_result.get("记录"))
                if llm_groups and multi_entity_hint:
                    current_projects = {
                        p.get("standard_name") for p in project_result.get("projects") or []
                        if p.get("standard_name")
                    }
                    llm_projects = {
                        p for g in llm_groups for p in g.get("projects") or []
                    }
                    low_confidence = any(
                        g.get("llm_confidence") not in {"high", "medium"}
                        for g in llm_groups
                    )
                    # 覆盖同一个项目不代表覆盖全部申请人。批量邮件必须同时
                    # 保留主体，以及规则已经明确建立的公司—项目关系。
                    current_groups = project_result.get("groups") or structured_groups
                    complete_mapping = (
                        (not current_projects or current_projects.issubset(llm_projects))
                        and self._record_mapping_covers(
                            current_groups, llm_groups, entity_companies
                        )
                    )
                    if complete_mapping:
                        project_result = {
                            "projects": [
                                {"raw_value": p, "standard_name": p}
                                for g in llm_groups for p in g.get("projects") or []
                            ],
                            "epr_form": None,
                            "need_review": low_confidence,
                            "groups": llm_groups,
                        }
                        self._log(
                            f"LLM 多公司关联提取: {[(g['customer'], g['projects']) for g in llm_groups]}"
                        )
                    else:
                        project_result["need_review"] = True
                        self._log("LLM 多公司关联未覆盖全部主体或公司项目关系，保留规则结果并转人工", "warning")
            else:
                # 模型调用失败或输出没有通过 Pydantic 时，绝不使用部分/畸形结果；
                # 保留规则结果并阻断进入阶段二，交由人工查看原邮件。
                project_result["need_review"] = True
                workflow = getattr(self.llm_client, "agent_workflow", None)
                last = getattr(workflow, "last_result", None)
                llm_failure_reason = getattr(last, "reason", "") or "模型调用失败或输出结构不合格"
                self._log(f"字段 LLM 输出未采用，已转人工: {llm_failure_reason}", "warning")
            customer_result = self._sanitize_customer_result(customer_result)

        # LLM 仅补字段后仍必须重新核对项目覆盖；缺项不能因为模型调用成功
        # 就被误标为已确认。
        remaining_coverage_missing = self._project_coverage_gap(subject, body, project_result)
        if remaining_coverage_missing:
            project_result["need_review"] = True
            project_result["coverage_missing"] = remaining_coverage_missing

        # E3: OCR 图片兜底 — 常规源(标题/正文/文档附件/文件名)+LLM 全部读完仍缺字段才触发
        # 触发条件只看 客户/项目（OCR 能实际补的字段）; 代理缺失不单独触发
        # （代理主要靠邮箱表匹配, 表未命中时几乎每封都缺, 单独触发会使兜底退化为常态）
        # 代理在兜底过程中机会性提取（OCR 文本中的邮箱查代理表）
        ocr_used = False
        still_missing = (
            (not customer_result.get("customer") and not complete_group_customers)
            or (not project_result.get("projects") and not has_epr_form)
        )
        if self.ocr_fallback and still_missing:
            ocr_used = self._ocr_fallback(attachments, agent_result, customer_result, project_result)
        customer_result = self._sanitize_customer_result(customer_result)
        if not customer_result.get("customer_code"):
            customer_result["customer_code"] = self._client_code_near_company(
                body, customer_result.get("customer", "")
            ) or self._client_code_near_company(subject, customer_result.get("customer", ""))

        # 勾选不完整 / 未勾选 → 强制转人工(不因 OCR 或 LLM 而"猜"出项目)
        need_review = bool(project_result.get("need_review"))

        # 组装输出
        rows = []
        row_groups = project_result.get("groups") or []
        if row_groups:
            # 结构化清单: 每家公司 × 每个项目 一行
            for g in row_groups:
                project_codes = g.get("project_codes") if isinstance(g.get("project_codes"), dict) else {}
                g_customer = {
                    "customer": g.get("customer", ""),
                    "source": g.get("source", "待确认"),
                    "customer_code": g.get("customer_code", ""),
                    "attachment_record_source": g.get("attachment_record_source", ""),
                    "agent": g.get("agent", ""),
                    "agent_match_method": g.get("agent_match_method", ""),
                    "company_candidates": g.get("company_candidates", []),
                }
                group_projects = [name for name in g.get("projects") or [] if name]
                # 附件表已证明该行存在，但项目列无法规范化时也必须保留一条
                # 占位复核记录；不能因为项目为空把“第 16～18 行”吞掉。
                if not group_projects:
                    rows.append(self._build_row(
                        mail, subject, body, agent_result, g_customer,
                        None, need, True, llm_used, ocr_used,
                    ))
                    continue
                for proj_name in group_projects:
                    codes = project_codes.get(proj_name, [])
                    g_customer["customer_code"] = (
                        codes[-1] if isinstance(codes, list) and codes else g.get("customer_code", "")
                    )
                    proj = {"standard_name": proj_name, "raw_value": proj_name}
                    rows.append(self._build_row(
                        mail, subject, body, agent_result, g_customer,
                        proj, need, need_review, llm_used, ocr_used,
                    ))
        elif project_result["projects"]:
            for proj in project_result["projects"]:
                rows.append(self._build_row(
                    mail, subject, body, agent_result, customer_result,
                    proj, need, need_review, llm_used, ocr_used,
                ))
        else:
            # 没提取到项目，也要输出一行
            rows.append(self._build_row(
                mail, subject, body, agent_result, customer_result,
                None, need, need_review, llm_used, ocr_used,
            ))

        # 结构化行异常（例如所有行都没有可用项目）不能让上层因 rows[0]
        # 崩溃，也不能静默丢弃邮件。生成一条明确的人工复核记录。
        if not rows:
            self._log("结构化邮件未生成任何字段行，已降级为人工复核记录", "warning")
            rows.append(self._build_row(
                mail, subject, body, agent_result, customer_result,
                None, need, True, llm_used, ocr_used,
            ))

        # 把“附件表格预期记录数”直接写到每一条结果中，后续 M4 标准化和
        # 工作台都能复核，而不是只在日志中说“可能少了几条”。
        expected_count = int(project_result.get("attachment_record_count") or 0)
        if expected_count:
            actual_count = len(rows)
            count_ok = actual_count == expected_count
            count_message = (
                f"附件结构化表 {expected_count} 条，输出 {actual_count} 条"
                + ("，数量一致" if count_ok else "，数量不一致")
            )
            if not count_ok:
                self._log(count_message + "；转人工复核", "warning")
            for row in rows:
                row["附件表格记录数"] = expected_count
                row["附件表格输出数"] = actual_count
                row["附件表格数量校验"] = "通过" if count_ok else "需人工确认"
                if not count_ok:
                    row["置信度"] = "需人工确认"
                    existing = str(row.get("人工复核提示", "") or "").strip()
                    row["人工复核提示"] = "；".join(
                        value for value in (existing, count_message) if value
                    )

        for row in rows:
            if llm_attempted and llm_used:
                row["字段LLM状态"] = "结构校验通过"
                row["字段LLM失败原因"] = ""
            elif llm_attempted:
                row["字段LLM状态"] = "失败，已转人工复核"
                row["字段LLM失败原因"] = llm_failure_reason
                row["置信度"] = "需人工确认"
                existing = str(row.get("人工复核提示", "") or "").strip()
                hint = f"字段LLM未采用: {llm_failure_reason}"
                row["人工复核提示"] = "；".join(x for x in (existing, hint) if x)
            else:
                row["字段LLM状态"] = "未调用"
                row["字段LLM失败原因"] = ""

        return rows

    def _build_row(self, mail: dict, subject: str, body: str,
                   agent_result: dict, customer_result: dict,
                   proj: Optional[dict], need: str,
                   need_review: bool, llm_used: bool, ocr_used: bool) -> dict:
        """组装单行输出; proj=None 表示项目缺失的占位行"""
        row_agent = str(
            customer_result.get("agent") or agent_result.get("agent") or ""
        ).strip()
        row_agent_method = str(
            customer_result.get("agent_match_method")
            or agent_result.get("match_method")
            or "待确认"
        ).strip()
        if proj is not None:
            confidence = "需人工确认" if (need_review or mail.get("filter_status") == "uncertain") else (
                "low" if ocr_used else self._calc_confidence(agent_result, customer_result, proj)
            )
            project = proj["standard_name"]
            raw_value = proj.get("raw_value", "")
        else:
            confidence = "需人工确认" if (need_review or mail.get("filter_status") == "uncertain") else "low"
            project = ""
            raw_value = "EPR申请表勾选不完整(见需人工复查)" if need_review else ""
        # 德国 WEEE 专项：只读取邮件/附件中明确出现的品牌与品类，
        # 后续阶段二再把这些明细与注册工单的品类明细逐项比对。
        weee = extract_weee_items(
            subject=subject,
            body=mail.get("body_original") or body,
            attachments=mail.get("attachments") or [],
            project=project,
            llm_client=self.llm_client,
            company=customer_result.get("customer", ""),
        )
        weee_items = weee.get("items") or []
        return {
            "sender_email": mail.get("sender_email", ""),
            "sender_name": mail.get("sender_name", ""),
            "recipient": mail.get("recipient", ""),
            "date": mail.get("date"),
            "subject": subject,
            "body_text": body,
            # 抽取用 body_text 已去除签名/寒暄；证据区必须保留完整正文。
            "body_original": (
                mail.get("body_original")
                or mail.get("body_raw")
                or body
            ),
            "附件名称": "；".join(
                str(att.get("filename", "")).strip()
                for att in (mail.get("attachments") or [])
                if att.get("filename")
            ),
            # 把附件中的可回溯行/工作表快照一起带到阶段一产物。这里使用
            # 紧凑 JSON，而不是临时文件路径；邮件解析结束后临时 xlsx 会被
            # 删除，工作台仍能直接展示并高亮本次字段对应的表格行。
            "附件证据": self._attachment_evidence_json(
                mail.get("attachments") or [],
                customer_result.get("attachment_record_source", ""),
            ),
            # 只写安全文件名，不写绝对路径。工作台通过同一附件缓存目录
            # 提供打开/下载；旧邮件没有该列时仍可使用结构化预览。
            "附件文件索引": self._attachment_file_index_json(mail.get("attachments") or []),
            "代理": row_agent,
            "代理匹配方式": row_agent_method,
            "客户编号": customer_result.get("customer_code", ""),
            "客户": customer_result["customer"],
            "客户提取来源": customer_result["source"],
            "附件明细来源": customer_result.get("attachment_record_source", ""),
            "客户候选证据": json.dumps(
                customer_result.get("company_candidates") or [],
                ensure_ascii=False,
                separators=(",", ":"),
            ),
            "项目": project,
            "项目原始值": raw_value,
            "需求": need,
            "德国WEEE专项": "是" if weee.get("enabled") else "否",
            "德国WEEE品类明细": json.dumps(weee_items, ensure_ascii=False, separators=(",", ":")) if weee.get("enabled") else "",
            "德国WEEE品类状态": ("待工单品类核对" if weee.get("status") in {"ready", "pending"} else "不适用"),
            "德国WEEE品类核对": "",
            "德国WEEE专项说明": (
                f"字段提取：{weee.get('extraction_method', '规则提取')}；"
                f"复检：{(weee.get('llm_review') or {}).get('status', '未启用')}"
            ) if weee.get("enabled") else "",
            "德国WEEE字段提取方式": weee.get("extraction_method", "规则提取") if weee.get("enabled") else "",
            "德国WEEE字段复检": json.dumps(weee.get("llm_review") or {}, ensure_ascii=False, separators=(",", ":")) if weee.get("enabled") else "",
            "置信度": confidence,
            "match_candidates": agent_result.get("candidates", []),
            "llm_used": llm_used,
            "ocr_used": ocr_used,
            "uid": mail.get("uid"),
            "filter_status": mail.get("filter_status"),
            "filter_reason": mail.get("filter_reason", ""),
        }

    @staticmethod
    def _attachment_evidence_json(attachments: list, record_source: str = "") -> str:
        """生成可写入 Excel 的附件证据快照。

        结构化 xlsx 行按当前明细保留，工作台会按同一封邮件合并这些行；
        非结构化工作表保留有限预览。所有内容均截断，避免超出 Excel 单元格
        32767 字符上限，也避免把整份大型附件重复写入每一行。
        """
        payload = []
        source_text = str(record_source or "")
        for attachment in attachments or []:
            if not isinstance(attachment, dict):
                continue
            filename = str(attachment.get("filename") or "").strip()
            if not filename:
                continue
            records = []
            raw_records = attachment.get("structured_records") or []
            for record in raw_records:
                if not isinstance(record, dict):
                    continue
                # 附件表格行与当前输出行一一对应时只写这一行；普通明细
                # 没有定位信息则保留该附件的有限记录，供工作台汇总展示。
                record_filename = str(record.get("attachment_name") or filename).strip()
                locator = f"{record_filename} / {str(record.get('sheet_name') or '工作表').strip()} 第{record.get('row_number') or '?'}行"
                # EPR forms use member/sheet-level sources and can contain both
                # Chinese and English company rows. Preserve their actual row
                # evidence across the mail; a form source is not a row locator.
                if source_text and locator not in source_text and str(record.get("record_type") or "") not in {"weee_catalog", "epr_application"}:
                    continue
                records.append({
                    "attachment_name": record_filename,
                    "sheet_name": str(record.get("sheet_name") or "工作表"),
                    "row_number": record.get("row_number") or "?",
                    "column_number": record.get("column_number") or "",
                    "customer": str(record.get("customer") or "")[:240],
                    "cells": [str(cell)[:240] for cell in (record.get("cells") or [])[:24]],
                    "raw_text": str(record.get("raw_text") or "")[:1200],
                    "brand": str(record.get("brand") or "")[:240],
                    "category": str(record.get("category") or "")[:240],
                    "record_type": str(record.get("record_type") or "entity"),
                })
            sheets = []
            for sheet in attachment.get("sheets") or []:
                if not isinstance(sheet, dict):
                    continue
                preview_rows = []
                for row in (sheet.get("preview_rows") or [])[:80]:
                    if not isinstance(row, dict):
                        continue
                    preview_rows.append({
                        "row_number": row.get("row_number") or "?",
                        "cells": [str(cell)[:160] for cell in (row.get("cells") or [])[:24]],
                    })
                sheets.append({
                    "sheet_name": str(sheet.get("sheet_name") or "工作表"),
                    "member_path": str(sheet.get("member_path") or ""),
                    "rows": preview_rows,
                })
            text_preview = str(attachment.get("text_content") or "")[:1200]
            # 仅对 xlsx/结构化附件写表格行；普通附件仍提供文本快照。
            if records or sheets or text_preview:
                payload.append({
                    "filename": filename,
                    "records": records,
                    "sheets": sheets,
                    "text": text_preview,
                })
        if not payload:
            return ""
        try:
            from utils.evidence_snapshot import encode_evidence_snapshot
            encoded = encode_evidence_snapshot(payload)
        except (TypeError, ValueError):
            return ""
        return encoded

    @staticmethod
    def _attachment_file_index_json(attachments: list) -> str:
        """把已持久化附件映射为工作台可用的安全 token。

        文件本体保存在 cache/attachments，Excel 只保存 basename，避免把本机
        路径暴露到导出文件，也避免 Windows 路径在不同机器上失效。
        """
        items = []
        for attachment_index, attachment in enumerate(attachments or []):
            if not isinstance(attachment, dict):
                continue
            filename = str(attachment.get("filename") or "").strip()
            filepath = str(attachment.get("filepath") or "").strip()
            if not filename or not filepath:
                continue
            token = os.path.basename(filepath)
            if not token or token in {".", ".."}:
                continue
            # 相同文件名甚至相同内容也可能是邮件中的两个独立附件；序号只用于
            # 工作台稳定区分列表项，不改变原件 token 或阶段二业务字段。
            items.append({
                "filename": filename,
                "token": token,
                "attachment_index": attachment_index,
            })
        if not items:
            return ""
        try:
            return json.dumps(items, ensure_ascii=False, separators=(",", ":"))
        except (TypeError, ValueError):
            return ""

    def _attachment_structured_groups(self, attachments: list, project_result: dict) -> List[dict]:
        """把附件 Excel 中的物理数据行转成一行一个业务主体的分组。

        这里不做“18 × 正文 15 项”的笛卡尔积。表格行自带项目时用该项目；
        表格未写项目而邮件只有一个已确认项目时才安全继承；其余情况保留空项目
        并转人工复核，仍然保留附件中的每一行。
        """
        fallback_projects = list(dict.fromkeys(
            str(project.get("standard_name", "") or "").strip()
            for project in project_result.get("projects") or []
            if str(project.get("standard_name", "") or "").strip()
        ))
        groups: List[dict] = []
        seen_rows = set()
        for attachment in attachments or []:
            for record in attachment.get("structured_records") or []:
                attachment_name = str(
                    record.get("attachment_name") or attachment.get("filename") or "附件表格"
                ).strip()
                # EPR 申请表单主体字段属于“每份表单一个申请人”，由项目解析
                # 按表单路径绑定；不能作为普通清单数据行重复展开。
                if str(record.get("record_type") or "") == "epr_application":
                    continue
                # 品牌/品类清单是 WEEE 专项证据，不是一个客户明细；
                # 不把它展开成空公司工单行，避免破坏原有“附件行数=业务明细数”口径。
                if str(record.get("record_type") or "") == "weee_catalog" and not str(record.get("customer") or "").strip():
                    continue
                sheet_name = str(record.get("sheet_name") or "工作表").strip()
                row_number = record.get("row_number") or "?"
                row_key = (attachment_name, sheet_name, str(row_number))
                if row_key in seen_rows:
                    continue
                seen_rows.add(row_key)

                raw_customer = str(record.get("customer") or "").strip()
                cleaned = self._sanitize_customer_result({
                    "customer": raw_customer,
                    "source": "附件表格结构化",
                }) if raw_customer else {"customer": "", "source": "附件表格结构化"}
                # 表头已判定为客户列，因此即便是“商贸行/品牌名”这类无公司后缀
                # 的合法主体也不能被通用公司正则清空；但如果清洗器明确判定
                # 这是 POA、注册资本、签字时间等表单字段，不能再用 raw_customer
                # 回填，否则这些标签会重新污染客户字段。
                customer = cleaned.get("customer", "")
                raw_project_text = " ".join(
                    str(record.get(key) or "").strip()
                    for key in ("project", "country", "business", "request", "raw_text")
                ).strip()
                record_projects = self._projects_from_attachment_record(raw_project_text)
                if not record_projects and len(fallback_projects) == 1:
                    record_projects = list(fallback_projects)

                source = f"附件表格：{attachment_name} / {sheet_name} 第{row_number}行"
                attachment_agent = self._canonical_attachment_agent(record.get("agent", ""))
                groups.append({
                    "customer": customer,
                    "customer_code": str(record.get("customer_code") or "").strip(),
                    "agent": attachment_agent,
                    "agent_match_method": "附件表格代理列" if attachment_agent else "",
                    "source": "附件表格结构化",
                    "attachment_record_source": source,
                    "projects": record_projects,
                    "needs_review": not bool(customer and record_projects),
                })
        return groups

    def _projects_from_attachment_record(self, record_text: str) -> List[str]:
        """从一条 Excel 明细行取项目；只返回该行能直接支持的项目。"""
        if not record_text:
            return []
        projects = self._drop_generic_epr(
            self._extract_projects_by_rules(record_text)
        )
        found = [
            str(project.get("standard_name") or "").strip()
            for project in projects if str(project.get("standard_name") or "").strip()
        ]
        # 项目名称表的精确命中补充规则命中，兼容“标准化项目名称”这一列。
        for name in self.project_name_list:
            if name and name in record_text and name not in found:
                found.append(name)
        return list(dict.fromkeys(found))

    def _prefer_attachment_record_groups(self, project_result: dict, groups: List[dict]) -> dict:
        """以附件结构化行替换同邮件的正文展开，并保留可审计的数量基线。"""
        emitted = []
        for group in groups:
            emitted.extend(str(name) for name in group.get("projects") or [] if name)
        # 空项目占位行也要保留，避免被后续“无项目”判断当作附件不存在。
        result = dict(project_result)
        result["groups"] = groups
        result["projects"] = [
            {"raw_value": name, "standard_name": name}
            for name in dict.fromkeys(emitted)
        ] or list(project_result.get("projects") or [])
        result["attachment_record_count"] = len(groups)
        result["need_review"] = bool(result.get("need_review")) or any(
            group.get("needs_review") for group in groups
        )
        self._log(
            f"附件结构化表优先: {len(groups)} 条记录，"
            f"来源={[group.get('attachment_record_source') for group in groups[:3]]}"
            + (" …" if len(groups) > 3 else "")
        )
        return result

    def _ocr_fallback(self, attachments: list, agent_result: dict,
                      customer_result: dict, project_result: dict) -> bool:
        """
        E3: OCR 图片兜底提取
        - 只在常规源+LLM 之后仍缺失字段时被调用
        - 只补缺失字段, 不覆盖已有结果
        - 结果标记来源=OCR图片兜底, 不作为可信真值（后续人工复核）
        - 一张图识别出多个项目 → project_result 收多条 → 上层拆多行输出
        - 全部图片识别后仍无目标信息 → 返回 False（保持原空结果）
        """
        pending = [a for a in attachments if a.get("ocr_pending") and a.get("filepath")]
        if not pending:
            return False

        self._log(f"字段缺失, 启动OCR兜底: {len(pending)} 张图片")
        filled = False
        for att in pending:
            text = self._ocr_image(att["filepath"])
            att["ocr_pending"] = False
            if text:
                att["text_content"] = text
            if not text:
                continue

            # 代理: OCR 文本中的邮箱 → 查代理表
            if not agent_result.get("agent"):
                try:
                    from utils.attachment_parser import extract_emails_from_text
                    for em in extract_emails_from_text(text):
                        entry = self.agent_email_map.get(em.lower())
                        if entry and entry.get("代理"):
                            agent_result["agent"] = entry["代理"]
                            agent_result["match_method"] = "OCR图片兜底"
                            filled = True
                            break
                except Exception:
                    pass

            # 客户: 公司名模式抽取
            if not customer_result.get("customer"):
                cust = self._extract_company_from_attachment_text(text)
                if cust:
                    customer_result["customer"] = cust
                    customer_result["source"] = "OCR图片兜底"
                    filled = True

            # 项目: 附件四标准名 + 离线规则(国家+业务), 多项目收集后由上层拆行
            # 例外: 已有 EPR 申请表勾选结果时, 勾选是唯一权威, OCR 不得补/改项目
            if not project_result.get("projects") and not project_result.get("epr_form"):
                projs = self._extract_projects_by_rules(text)
                existing = {p["standard_name"] for p in projs}
                for name in self.project_name_list:
                    if name and name in text and name not in existing:
                        projs.append({"raw_value": name, "standard_name": name})
                        existing.add(name)
                if projs:
                    project_result["projects"] = projs
                    filled = True

        if not filled:
            self._log("OCR兜底未提取到目标字段", "warning")
        return filled

    def _ocr_image(self, filepath: str) -> str:
        """调用 PaddleOCR 识别单张图片, 失败返回空串（不影响主流程）"""
        try:
            from utils.attachment_parser import _parse_image
            return _parse_image(filepath) or ""
        except Exception as e:
            self._log(f"OCR兜底识别失败 {filepath}: {e}", "warning")
            return ""

    def _match_llm_projects(self, llm_projects: List[str]) -> dict:
        """将 LLM 返回的项目名与附件四做保守的相似度标准化。"""
        from rapidfuzz import fuzz

        matched = []
        for proj in llm_projects:
            proj_norm = normalize_text(proj)
            best_match = None
            best_score = 0
            for std_name in self.project_name_list:
                std_norm = normalize_text(std_name)
                if proj_norm == std_norm:
                    best_match = std_name
                    best_score = 100
                    break
                score = max(
                    fuzz.ratio(proj_norm, std_norm),
                    fuzz.partial_ratio(proj_norm, std_norm),
                )
                if proj_norm in std_norm or std_norm in proj_norm:
                    score = max(score, 95)
                if score > best_score:
                    best_match = std_name
                    best_score = score
            # 阈值不足时保留 LLM 原文并交由后续人工/normalizer 判断，不能因为
            # 名称更长就强行映射到貌似相近的附件四项目。
            if best_match and best_score >= 88:
                matched.append({"raw_value": proj, "standard_name": best_match})
            else:
                matched.append({"raw_value": proj, "standard_name": proj})

        # 去重
        seen = set()
        unique = []
        for p in matched:
            if p["standard_name"] not in seen:
                seen.add(p["standard_name"])
                unique.append(p)
        return {"projects": unique}

    def _llm_record_groups(self, records) -> List[dict]:
        """把 LLM 返回的多公司记录转成阶段一可展开的结构化分组。

        LLM 只能作为规则层的补充，故这里不接受空客户/空项目，也不把一个
        模糊的自然语言段落直接写入客户列。项目仍经过同一套标准化匹配，避免
        LLM 返回 ``德国电池``、``Battery DE`` 等非标准写法后造成重复项目。
        """
        if not isinstance(records, list):
            return []
        grouped = {}
        for record in records:
            if not isinstance(record, dict):
                continue
            customer = str(
                record.get("客户") or record.get("customer") or
                record.get("公司") or record.get("company") or ""
            ).strip()
            if not customer or len(customer) > 100:
                continue
            clean_customer = self._accept_llm_customer(customer)
            if not clean_customer:
                continue

            raw_projects = record.get("项目")
            if raw_projects is None:
                raw_projects = record.get("projects")
            if isinstance(raw_projects, str):
                raw_projects = [p.strip() for p in re.split(r"[+＋/／、，,;；&＆]", raw_projects) if p.strip()]
            if not isinstance(raw_projects, list):
                continue
            raw_projects = [str(p).strip() for p in raw_projects if str(p).strip()]
            if not raw_projects:
                continue
            matched = self._match_llm_projects(raw_projects).get("projects") or []
            projects = [p["standard_name"] for p in matched if p.get("standard_name")]
            if not projects:
                continue

            key = normalize_text(clean_customer)
            item = grouped.setdefault(key, {
                "customer": clean_customer,
                "source": "LLM多公司关联",
                "projects": [],
                "llm_confidence": str(record.get("confidence") or "medium").lower(),
            })
            confidence = str(record.get("confidence") or "medium").lower()
            rank = {"low": 0, "medium": 1, "high": 2}
            if rank.get(confidence, 0) < rank.get(item.get("llm_confidence", "low"), 0):
                item["llm_confidence"] = confidence
            for project in projects:
                if project not in item["projects"]:
                    item["projects"].append(project)
        return list(grouped.values())

    @staticmethod
    def _record_mapping_covers(current_groups, proposed_groups, expected_companies=()) -> bool:
        """模型替换必须覆盖全部主体及已经确定的逐主体项目关系。"""
        proposed = {}
        for group in proposed_groups or []:
            key = normalize_text(group.get("customer", ""))
            if key:
                proposed.setdefault(key, set()).update(group.get("projects") or [])
        expected = {normalize_text(name) for name in expected_companies if name}
        for group in current_groups or []:
            key = normalize_text(group.get("customer", ""))
            if not key:
                continue
            expected.add(key)
            # 待确认的交叉展开可由模型收窄，但主体本身仍不能消失。
            if not group.get("needs_review") and not set(group.get("projects") or []).issubset(
                proposed.get(key, set())
            ):
                return False
        return expected.issubset(proposed)

    def _extract_need(self, text: str) -> str:
        """提取需求类型: 注册/新增/撤单"""
        if "撤单" in text:
            return "撤单"
        if "新增" in text:
            return "新增"
        if "注册" in text:
            return "注册"
        return ""

    @staticmethod
    def _agent_name_from_entry(entry) -> str:
        if isinstance(entry, dict):
            return str(entry.get("代理", "") or entry.get("agent", "")).strip()
        return str(entry or "").strip()

    def _canonical_attachment_agent(self, value: str) -> str:
        """将附件注册表代理列映射到代理表正式名称。"""
        raw = str(value or "").strip()
        if not raw:
            return ""
        folded = normalize_text(raw).casefold()
        matches = []
        for alias, officials in self._agent_alias_index.items():
            if normalize_text(alias).casefold() == folded:
                matches.extend(officials)
        matches = list(dict.fromkeys(matches))
        # 附件列是明确的行级来源；代理表未登记时保留原值，交人工复核。
        return matches[0] if len(matches) == 1 else raw

    def _build_agent_alias_index(self) -> Dict[str, List[str]]:
        """建立 代理别名 → 正式代理名 的索引。

        代理表存在“代理”“代理简称”，个别手工表还会有“别名/代理别名”列。
        这些名称用于主题匹配；输出永远写正式代理名，避免把别名带入工单系统。
        """
        index: Dict[str, List[str]] = {}
        for entry in self.agent_email_map.values():
            official = self._agent_name_from_entry(entry)
            if not official:
                continue
            aliases = [official]
            if isinstance(entry, dict):
                for key in ("代理简称", "别名", "代理别名", "alias", "aliases"):
                    value = entry.get(key)
                    if isinstance(value, (list, tuple, set)):
                        aliases.extend(str(v).strip() for v in value)
                    elif value:
                        aliases.extend(
                            part.strip() for part in re.split(r"[、,，;；/／|]", str(value))
                        )
            for alias in aliases:
                alias = str(alias or "").strip()
                # 单字中文代理名误命中概率过高；英文单词也至少三位。
                if len(alias) < 2:
                    continue
                bucket = index.setdefault(alias, [])
                if official not in bucket:
                    bucket.append(official)
        return index

    def _extract_agent_from_subject(self, subject: str) -> Optional[dict]:
        """从主题中的代理全称/简称/别名精确识别代理。

        先匹配分隔符边界，再允许主题首段出现代理前缀（如“示例代理己——”）。
        相同别名指向多个正式代理时不猜测，返回多匹配人工确认。
        """
        if not subject or not self._agent_alias_index:
            return None
        candidates: List[tuple] = []
        head = re.sub(r"^\s*(?:re|fw|fwd)\s*[:：-]\s*", "", subject, flags=re.I)
        head = re.split(r"[+＋—–|]", head, maxsplit=1)[0]
        for alias, official_names in self._agent_alias_index.items():
            escaped = re.escape(alias)
            match = re.search(
                rf"(?:^|[\s+＋—–_\-:：/／|]){escaped}(?=$|[\s+＋—–_\-:：/／|,，;；(（])",
                subject,
                re.I,
            )
            position = match.start() if match else -1
            if position < 0:
                prefix_position = head.lower().find(alias.lower())
                if 0 <= prefix_position <= 8:
                    position = prefix_position
            # 代理名有时不是单独的分隔词，而是嵌在首段/代理描述中：
            # ``SY003-首阳跨境咨询-...``、``SY003-首阳-...``。
            # 仅在分隔后的短段开头匹配已知别名，避免从客户公司名中截取代理。
            if position < 0:
                for segment in re.finditer(
                    r"(?:^|[+＋—–|/:：\-])([^+＋—–|/:：\-]+)", subject
                ):
                    value = segment.group(1).strip()
                    if not value.casefold().startswith(alias.casefold()):
                        continue
                    suffix = value[len(alias):]
                    if not suffix or len(suffix) > 20:
                        continue
                    if re.match(r"^[\u4e00-\u9fffA-Za-z0-9]", suffix):
                        position = segment.start(1)
                        break
            if position >= 0:
                for official in official_names:
                    candidates.append((position, -len(alias), official, alias))
        if not candidates:
            return None
        candidates.sort()
        best_position, best_length = candidates[0][0], candidates[0][1]
        names = []
        for position, length, official, _ in candidates:
            if position != best_position or length != best_length:
                continue
            if official not in names:
                names.append(official)
        if len(names) == 1:
            return {
                "agent": names[0],
                "match_method": "主题代理别名精确匹配",
                "candidates": [],
            }
        return {
            "agent": " / ".join(names),
            "match_method": "主题代理别名多匹配-人工排查",
            "candidates": [{"text": "主题代理别名", "data": {"代理": name}, "score": 100}
                           for name in names],
        }

    def _extract_agent(
        self, sender_email: str, subject: str, body: str, attachments: list
    ) -> dict:
        """
        提取代理: 通过发件人邮箱查附件三
        逻辑: 精确匹配 → 模糊匹配 → 多匹配展示
        """
        def _get_agent_name(entry):
            return self._agent_name_from_entry(entry)

        # 精确匹配
        if sender_email in self.agent_email_map:
            entry = self.agent_email_map[sender_email]
            return {
                "agent": _get_agent_name(entry),
                "match_method": "邮箱精确匹配",
                "candidates": [],
            }

        # 发件邮箱未登记时，主题中的代理全称/简称比模糊邮箱更可靠。
        # 例如“巨齿鲨+客户公司+比利时包装法”。
        subject_match = self._extract_agent_from_subject(subject)
        if subject_match:
            return subject_match

        # 也尝试从正文中提取邮箱地址
        from utils.attachment_parser import extract_emails_from_text
        body_emails = extract_emails_from_text(body)
        for em in body_emails:
            em_lower = em.lower()
            if em_lower in self.agent_email_map:
                entry = self.agent_email_map[em_lower]
                return {
                    "agent": _get_agent_name(entry),
                    "match_method": "正文邮箱精确匹配",
                    "candidates": [],
                }

        # 模糊匹配 — 用发件人邮箱去模糊匹配代理邮箱表
        candidates = []
        for em, entry in self.agent_email_map.items():
            score = self._email_similarity(sender_email, em)
            if score >= 80:
                candidates.append({
                    "text": em,
                    "data": entry,
                    "score": score,
                })

        candidates.sort(key=lambda x: x["score"], reverse=True)

        if not candidates:
            self._log(f"代理未匹配: {sender_email}", "warning")
            return {
                "agent": "",
                "match_method": "待确认",
                "candidates": [],
            }

        if len(candidates) == 1 and candidates[0]["score"] >= 90:
            entry = candidates[0]["data"]
            return {
                "agent": _get_agent_name(entry),
                "match_method": "邮箱模糊匹配",
                "candidates": [],
            }

        # 多匹配 — 全部展示
        # 同一个代理常登记多个邮箱，直接 join 会拼出“向善 / 向善”这种重复值，
        # 会让下游下拉选择和字段比对必然失败；这里按规范化名去重后再拼接。
        self._log(f"代理多匹配({len(candidates)}条): {sender_email}", "warning")
        seen_agent_names = set()
        agent_names = []
        for c in candidates:
            name = _get_agent_name(c["data"])
            key = normalize_text(name)
            if not key or key in seen_agent_names:
                continue
            seen_agent_names.add(key)
            agent_names.append(name)
        return {
            "agent": " / ".join(agent_names),
            "match_method": "多匹配-人工排查",
            "candidates": candidates,
        }

    def _email_similarity(self, a: str, b: str) -> int:
        """邮箱地址相似度比较"""
        from rapidfuzz import fuzz
        a_norm = normalize_text(a)
        b_norm = normalize_text(b)
        return int(fuzz.ratio(a_norm, b_norm))

    @staticmethod
    def _attachment_company_text(attachment: dict) -> str:
        """把申请表附件的文本、工作表和结构化记录合并成可审计文本。"""
        chunks: List[str] = []
        if not isinstance(attachment, dict):
            return ""
        for key in ("text_content", "text"):
            value = attachment.get(key)
            if value:
                chunks.append(str(value))
        for sheet in attachment.get("sheets") or []:
            if not isinstance(sheet, dict):
                continue
            for row in sheet.get("rows") or []:
                if not isinstance(row, dict):
                    continue
                chunks.extend(str(cell) for cell in row.get("cells") or [] if cell is not None)
        for record in attachment.get("structured_records") or []:
            if isinstance(record, dict):
                chunks.extend(str(value) for value in record.values() if value is not None)
        return "\n".join(chunks)

    @classmethod
    def _looks_like_epr_application(cls, attachment: dict) -> bool:
        """只把明确的 EPR 申请/注册表当作公司权威来源。"""
        if not isinstance(attachment, dict):
            return False
        if any(isinstance(item, dict) for item in (attachment.get("epr_forms") or [])):
            return True
        hint = " ".join(
            str(attachment.get(key) or "")
            for key in ("filename", "text_content", "text")
        )
        return bool(re.search(
            r"(?:EPR\s*(?:application|registration)?\s*form|泛欧\s*EPR|"
            r"EPR申请表|EPR申报表|申请表|申报表|注册申请|registration\s+form)",
            hint,
            re.I,
        ))

    def _extract_customer_from_epr_application(self, attachments: list) -> dict:
        """先读取 EPR 表格明确的公司字段；没有时才回退到申请表文本正则。"""
        candidates: List[dict] = []
        for attachment in attachments or []:
            if not self._looks_like_epr_application(attachment):
                continue
            structured_candidates_before = len(candidates)
            for record in attachment.get("structured_records") or []:
                if not isinstance(record, dict):
                    continue
                if record.get("record_type") == "epr_application":
                    value = str(record.get("customer") or "").strip()
                    cleaned = self._sanitize_customer_result({
                        "customer": value, "source": "附件EPR申请表结构化字段",
                    }).get("customer", "")
                    if cleaned:
                        candidates.append({
                            "value": cleaned,
                            "field_type": str(record.get("company_field") or "company_name"),
                            "source": str(record.get("attachment_name") or attachment.get("filename") or "EPR申请表"),
                            "evidence": str(record.get("raw_text") or record.get("field_label") or "")[:300],
                        })
                    continue
                for key in (
                    "customer", "company", "客户", "客户公司", "客户公司名称",
                    "公司", "公司名称", "申请单位", "applicant", "company name",
                ):
                    value = str(record.get(key) or "").strip()
                    if not value:
                        continue
                    candidate = self._company_substring(value) or self._extract_company_from_attachment_text(value)
                    if candidate:
                        candidates.append({
                            "value": candidate,
                            "field_type": "company_name",
                            "source": str(record.get("attachment_name") or attachment.get("filename") or "EPR申请表"),
                            "evidence": str(record.get("raw_text") or value)[:300],
                        })
            # 格式化字段优先。只有该表没有任何有效标签-值结果时，才扫描
            # 申请表文本，避免把“Company name (in Chinese)”这类表头误作主体。
            if len(candidates) == structured_candidates_before:
                text = self._attachment_company_text(attachment)
                candidate = self._extract_company_from_attachment_text(text)
                if candidate:
                    candidates.append({
                        "value": candidate,
                        "field_type": "regex_fallback",
                        "source": str(attachment.get("filename") or "EPR申请表文本"),
                        "evidence": text[:300],
                    })

        unique: List[dict] = []
        seen = set()
        for index, candidate in enumerate(candidates, start=1):
            cleaned = self._sanitize_customer_result({
                "customer": candidate.get("value", ""),
                "source": "附件EPR申请表",
            }).get("customer", "")
            key = normalize_text(cleaned)
            if cleaned and key and key not in seen:
                seen.add(key)
                unique.append({
                    "id": f"epr-company-{index}",
                    "value": cleaned,
                    "field_type": candidate.get("field_type", "company_name"),
                    "source": candidate.get("source", "EPR申请表"),
                    "evidence": candidate.get("evidence", ""),
                })
        if not unique:
            return {}
        if len(unique) > 1:
            self._log(
                f"EPR申请表发现多个公司候选，转入一次性语义仲裁: {[item['value'] for item in unique]}",
                "warning",
            )
        field_priority = {"company_name_zh": 0, "company_name": 1, "applicant_name": 2, "company_name_en": 3, "regex_fallback": 4}
        preferred = min(unique, key=lambda item: field_priority.get(item.get("field_type"), 5))
        return {
            "customer": preferred["value"],
            # 兼容历史“来源”口径；抽取方式另记在可审计元数据中。
            "source": "附件EPR申请表",
            "company_extraction_method": "格式化字段优先" if preferred.get("field_type") != "regex_fallback" else "格式化字段未命中后正则回退",
            "company_candidates": unique if len(unique) > 1 else [],
        }

    def _extract_customer(
        self, subject: str, body: str, attachments: list
    ) -> dict:
        """
        提取客户(公司名称)
        优先级: EPR申请表 → 正文 → 主题 → 其它附件
        """
        # 正文是实际申请内容，优先于可能存在笔误或复制错误的主题。
        cust = self._extract_company_from_body(body)
        if cust:
            return {"customer": cust, "source": "正文"}

        # 从主题提取 — 主题格式通常是: 代理+客户号+公司名+项目+注册
        cust = self._extract_company_from_subject(subject)
        if cust:
            return {"customer": cust, "source": "主题"}

        # 从附件提取
        for att in attachments:
            # 从附件文件名提取
            fname = att.get("filename", "")
            cust = self._extract_company_from_filename(fname)
            if cust:
                return {"customer": cust, "source": f"附件文件名({fname})"}

            # 从附件内容提取
            text = att.get("text_content", "")
            cust = self._extract_company_from_attachment_text(text)
            if cust:
                return {"customer": cust, "source": "附件内容"}

        return {"customer": "", "source": "待确认"}

    def _cross_source_company_consensus(
        self, subject: str, body: str, attachments: list
    ) -> dict:
        """在主题、正文、附件名称/内容/结构化表格之间做保守的公司名印证。

        这里故意不把“出现次数最多”当成真值：同一正文可能重复复制一段，
        次数不能替代来源独立性。只有同一候选至少被两个不同来源支持，才允许
        修正一个冲突的规则候选；单一附件表格行仍可以作为强来源，但不会凭空
        推断出正文中不存在的其它公司。
        """
        source_candidates: Dict[str, List[str]] = {}

        def add_source(source: str, values) -> None:
            clean: List[str] = []
            for value in values or []:
                candidate = self._sanitize_customer_result({
                    "customer": str(value or ""), "source": source,
                }).get("customer", "")
                if not candidate:
                    continue
                # 只保留有公司主体证据的候选；表格客户列可以是品牌/商号，
                # 但不应把表单标签、联系人姓名或数量说明送入印证池。
                if self._customer_noise_reason(candidate):
                    continue
                key = normalize_text(candidate)
                if key and key not in {normalize_text(v) for v in clean}:
                    clean.append(candidate)
            if clean:
                source_candidates[source] = clean

        subject_candidates = self._company_substrings(subject)
        loose_subject_company = self._extract_company_from_subject(subject)
        if loose_subject_company and loose_subject_company not in subject_candidates:
            subject_candidates.append(loose_subject_company)
        add_source("主题", subject_candidates)
        # body_text 已是抽取用正文；body_original 由调用方在需要时保留到证据区，
        # 这里不把签名/免责声明重新当成公司候选。
        add_source("正文", self._company_substrings(body))

        attachment_names: List[str] = []
        attachment_text_candidates: List[str] = []
        attachment_records: List[str] = []
        for attachment in attachments or []:
            filename = str(attachment.get("filename", "") or "").strip()
            if filename:
                name = self._extract_company_from_filename(filename)
                if name:
                    attachment_names.append(name)
            text = str(attachment.get("text_content", "") or "")
            if text:
                attachment_text_candidates.extend(self._company_substrings(text))
            for record in attachment.get("structured_records") or []:
                if isinstance(record, dict) and record.get("customer"):
                    attachment_records.append(str(record.get("customer")))

        add_source("附件表格", attachment_records)
        add_source("附件文件名", attachment_names)
        add_source("附件内容", attachment_text_candidates)

        if not source_candidates:
            return {}

        support: Dict[str, dict] = {}
        source_priority = {"附件表格": 4, "附件内容": 3, "正文": 2, "主题": 2, "附件文件名": 1}
        for source, values in source_candidates.items():
            for value in values:
                key = normalize_text(value)
                if not key:
                    continue
                item = support.setdefault(key, {
                    "customer": value, "sources": set(), "mentions": 0,
                    "priority": 0,
                })
                item["sources"].add(source)
                item["mentions"] += 1
                item["priority"] = max(item["priority"], source_priority.get(source, 0))

        if not support:
            return {}
        ranked = sorted(
            support.values(),
            key=lambda item: (
                len(item["sources"]), item["priority"], item["mentions"],
                len(item["customer"]),
            ),
            reverse=True,
        )
        best = ranked[0]
        # 多来源共识才有资格纠正冲突值。只有附件表格单独提供时，规则提取
        # 本来也会优先使用它，不需要在这里重复覆盖。
        if len(best["sources"]) < 2:
            return {}
        source_label = "+".join(
            source for source in ("附件表格", "附件内容", "正文", "主题", "附件文件名")
            if source in best["sources"]
        )
        return {
            "customer": best["customer"],
            "source": f"{source_label}交叉印证",
            "sources": sorted(best["sources"]),
        }

    def _sanitize_customer_groups(self, groups: list) -> List[dict]:
        """清洗结构化公司—项目分组，禁止占位词穿透到业务明细。"""
        cleaned_groups: List[dict] = []
        for raw_group in groups or []:
            if not isinstance(raw_group, dict):
                continue
            group = dict(raw_group)
            raw_customer = str(group.get("customer", "") or "").strip()
            # 说明句中的公司后缀可能被正则截成“以下为某某有限公司”。
            # 这是发件方/代理的引导语，不是客户主体；确定为此类值时直接
            # 丢弃，避免同一封邮件多出一张假的客户明细。其他不确定候选仍留给人工。
            if self._customer_noise_reason(raw_customer) == "发件方说明句":
                self._log(f"忽略发件方说明句中的公司候选: {raw_customer}", "warning")
                continue
            result = self._sanitize_customer_result({
                "customer": raw_customer,
                "source": group.get("source", "结构化分组"),
            })
            group["customer"] = result.get("customer", "")
            if result.get("source"):
                group["source"] = result["source"]
            if not group["customer"]:
                group["needs_review"] = True
            cleaned_groups.append(group)
        return cleaned_groups

    def _attachment_company_names(self, attachments: list) -> List[str]:
        """从全部附件文件名提取去重后的公司名，保留出现顺序。"""
        names: List[str] = []
        seen = set()
        for att in attachments or []:
            candidate = self._extract_company_from_filename(
                str(att.get("filename", "") or "")
            )
            if not candidate:
                continue
            key = normalize_text(candidate)
            if key and key not in seen:
                seen.add(key)
                names.append(candidate)
        return names

    def _extract_company_from_subject(self, subject: str) -> str:
        """从主题中提取公司名。
        在完整主题上做公司后缀锚定, 取最长合法子串 — 避免把
        "注销（WEEE+电池）示例代理己——EG1970 示例2cc81081有限公司 德国WEEE"
        这类主题的流水号/项目前缀一起带出来。"""
        if not subject:
            return ""
        strict = self._company_substring(subject)
        if strict:
            return strict

        # 主题中的 CHZ482、EG3168、K-DED0943 等是客户/案件编号，不是公司名。
        # 对没有 SAS/LIMITED/有限公司后缀的英文商号，只在编号后的主题片段中
        # 读取，并在国家、项目或动作词前截断；不从普通正文自由猜测。
        code_match = self._CLIENT_CODE_RE.search(subject)
        if not code_match:
            return ""
        tail = subject[code_match.end():].strip(" \t-—_:：+＋")
        boundary = re.search(
            r"(?:德国|比利时|比利時|法国|法國|意大利|義大利|西班牙|荷兰|荷蘭|"
            r"波兰|波蘭|瑞典|爱尔兰|愛爾蘭|葡萄牙|奥地利|奧地利|芬兰|芬蘭|"
            r"WEEE|EEE|EPR|电池法|電池法|包装法|包裝法|一次性塑料|"
            r"注册|註冊|新注册|新註冊|申报|申報|注销|註銷|修改|变更|變更)",
            tail,
            re.I,
        )
        candidate = tail[:boundary.start()] if boundary else tail
        candidate = candidate.strip(" \t-—_:：+＋,，;；")
        if len(candidate) < 4 or len(candidate) > 100:
            return ""
        if self._customer_noise_reason(candidate) or self._is_agent_name(candidate):
            return ""
        if re.fullmatch(r"[A-Za-z]{1,8}[-_]?\d{4,}", candidate):
            return ""
        if not re.search(r"[A-Za-z]{2,}|[\u4e00-\u9fff]{2,}", candidate):
            return ""
        return candidate

    def _sanitize_customer_result(self, result: dict) -> dict:
        """去除制表符拼接的标题/表格行，防止它被当作公司名传给阶段二。"""
        raw = str(result.get("customer", "") or "")
        noise_reason = self._customer_noise_reason(raw)
        if noise_reason:
            self._log(f"客户候选不是公司主体({noise_reason})，转人工补全", "warning")
            result["customer"] = ""
            result["source"] = "待确认"
            return result
        if not re.search(r"[\t\r\n]", raw):
            # 结构化表格/LLM 常直接返回一行候选，不能因为没有换行就跳过
            # 公司前的客户编号清洗。例如 `SED5632 xxxxxxxx B.V.` 的主体是
            # `xxxxxxxx B.V.`，SED5632 只应记录为客户编号。
            anchored = self._company_substring(raw)
            if anchored and normalize_text(anchored) != normalize_text(raw):
                result["customer"] = anchored
                result["source"] = f"{result.get('source', '规则')}清洗"
            else:
                result["customer"] = raw.strip()
            return result

        candidates = []
        for segment in re.split(r"[\t\r\n]+", raw):
            candidate = self._company_substring(segment.strip())
            if candidate:
                candidates.append(candidate)
        if candidates:
            result["customer"] = max(candidates, key=len)
            result["source"] = f"{result.get('source', '规则')}清洗"
            return result

        self._log("客户候选含制表符但未找到合法公司名，转人工补全", "warning")
        result["customer"] = ""
        result["source"] = "待确认"
        return result

    @staticmethod
    def _customer_noise_reason(value: str) -> str:
        """识别明确不是公司主体的字段标签/联系方式/说明文字。

        这是拒绝式闸门：只能拦截确定不是公司名的内容，不能根据常识
        生成或改写一个公司名。真实公司名无法确认时返回空字符串并交人工。
        """
        text = str(value or "").strip()
        if not text:
            return "客户为空"
        compact = re.sub(r"\s+", "", text).lower()
        hints = (
            "poa", "legalrepresentative", "legalperson", "legalpositions",
            "nameoflegalperson", "placeofsignature", "signingtime",
            "registrationcapital", "companyname", "companyaddress", "companybusiness",
            "plz", "postcode", "amazonlink", "shoplink", "e-mail", "email", "tel", "phone",
            "legrepresentativename", "companyregistrationnumber", "registrationnumber", "uscc",
            "营业执照", "公司名称", "公司中文名", "公司英文名", "中文公司名", "英文公司名",
            "公司地址", "公司注册", "注册资本", "法人",
            "公司成立日期", "成立日期", "签字", "签署", "职位", "联系信息", "联系人", "联系电话", "邮箱",
            "邮政编码", "邮编", "地址", "姓名", "身份证", "护照", "性别", "店铺链接",
            "平台信息", "服务内容", "服务的国家", "销售量", "预计销售", "说明", "备注", "请提供", "请选择", "填写",
            "非中国公司", "中国公司", "外国公司", "国公司",
            "不能与其它公司", "不能与其他公司", "与其它公司申请", "与其他公司申请",
            "foreigncompany", "nonchinesecompany",
            "香港公司", "中国公司", "中國公司", "大陆公司", "大陸公司",
            "海外公司", "境外公司", "外国公司", "外國公司", "欧洲公司", "欧盟公司",
            "附件香港公司", "附件中国公司", "附件外国公司",
        )
        if any(hint in compact for hint in hints):
            return "表单字段标签或说明文字"
        if re.match(r"^(?:针对|適用於|适用于|适用|用于|仅限|適用).*(?:业务|服務|服务|资料|材料|项目)$", text):
            return "表单字段标签或说明文字"
        # “以下为/下列为……提交……名单”是代理或发件方对后续名单的
        # 引导语。公司后缀正则会从整句截出“以下为某某有限公司”，旧逻辑
        # 因而把代理公司误当客户；按上下文前缀拒绝该候选。
        if re.match(
            r"^(?:以下|下面|下列|现将|本次)(?:为|是)?",
            text,
            re.I,
        ) and re.search(
            r"(?:提交|报送|发送|提供|列出|名单|新注册|申请)",
            text,
            re.I,
        ):
            return "发件方说明句"
        # 上一步公司后缀提取后可能只剩“以下为某某有限公司”，提交/名单
        # 等后文已被截掉；该前缀 + 公司后缀组合仍是确定的说明句候选。
        if re.match(r"^(?:以下|下面|下列)(?:为|是)", text, re.I) and re.search(
            r"(?:有限责任公司|股份有限公司|集团有限公司|有限公司|责任公司|公司|企业)$",
            text,
            re.I,
        ):
            return "发件方说明句"
        if "@" in text or re.search(r"https?://|www\.", text, re.I):
            return "邮箱或链接"
        if re.fullmatch(r"[+()\-\s\d]{6,}", text):
            return "联系方式或纯数字"
        if re.fullmatch(r"[A-Za-z]{1,8}[-_]?\d{4,}", text):
            return "客户编号或注册号"
        if re.fullmatch(r"(?:不含|包含|含有|无|有)?回收公司", text):
            return "业务说明中的通用词"
        if re.fullmatch(r"[0-9一二三四五六七八九十多几]*\s*家\s*(?:公司|主体|企业)", text, re.I):
            return "公司数量说明"
        # “一家公司-德国包装法”“2家公司：…”属于数量/说明，不是公司主体。
        # 允许后面有项目说明，仍然整体拒绝，避免把数量误当客户名。
        if re.match(
            r"^[0-9一二三四五六七八九十多几]+\s*家\s*(?:公司|主体|企业)"
            r"(?:\s*[-—:：,，/／+＋].*)?$",
            text,
            re.I,
        ):
            return "公司数量说明"
        # 日期/数量后的“告知以下 2 家公司”“通知 3 家企业”会被公司后缀
        # 正则截成一个看似主体的短片段，例如“24告知以下2家公司”。它是
        # 说明句，不是客户公司；真正的英文/中文公司主体不会以这些叙述词开头。
        if re.match(
            r"^\d{1,4}\s*(?:告知|通知|提交|说明|列出|涉及|有)"
            r".*(?:\d+\s*家\s*)?(?:公司|主体|企业)",
            text,
            re.I,
        ):
            return "公司数量说明"
        # “麻烦安排以下公司”“请安排下列公司”等是发件方引导语，
        # 公司后缀正则可能从整句末尾截出“……公司”，不能当成客户主体。
        if re.search(r"(?:安排|以下|下列|上述|涉及|通知|告知).{0,16}公司$", text):
            return "发件方公司引导语"
        # 英文公司名经常正好是两个首字母大写的词（如
        # ``Fixture 4898ee8f Inc``）。必须先确认法定公司后缀，再判断
        # “两个英文单词”是否像联系人姓名，否则真实 Inc/LLC 主体会被误删。
        if re.search(
            r"(?i)(?:CO\.?\s*,?\s*LTD\.?|LIMITED|LTD\.?|LLC|INC\.?|GMBH|"
            r"GBR|B\.?V\.?|S\.?A\.?S|SAS|N\.?V\.?|S\.?R\.?O\.?|SRO|"
            r"S\.?R\.?L\.?|O\.?U\.?|U\.?A\.?B\.?|PTE\.?\s*LTD\.?|"
            r"PTY\.?\s*LTD\.?|CORPORATION|CORP\.?|S\.?L\.?)\s*$",
            text,
        ):
            return ""
        # 申请表常把法人姓名单独列成一行（如 Huiming Wu）。没有公司后缀、
        # 只有英文姓名结构的候选不能作为客户主体；不确定时留空交人工。
        if re.fullmatch(r"[A-Z][a-z]{1,24}(?:\s+[A-Z][a-z]{1,24}){1,3}", text):
            return "疑似法人/联系人姓名"
        if len(text) > 120:
            return "说明段落"
        return ""

    def _accept_llm_customer(self, value: str) -> str:
        """只接受 LLM 能明确证明为公司主体的候选，不能确认则返回空。"""
        cleaned = self._sanitize_customer_result({
            "customer": str(value or ""), "source": "LLM补充"
        }).get("customer", "")
        if not cleaned:
            return ""

        # 公司后缀是最可靠的主体证据；允许中英文、重音字符及括号。
        anchored = self._company_substring(cleaned)
        if anchored and normalize_text(anchored) == normalize_text(cleaned):
            return anchored
        if re.search(
            rf"(?:有限责任公司|股份有限公司|集团有限公司|有限公司|责任公司|公司|集团|"
            rf"（个体工商户）|\(个体工商户\)|经营部|门市部|服务部|商店|商贸行|商行|工厂|工作室|店|"
            rf"{self._COMPANY_EN_SUFFIX})\s*$",
            cleaned,
            re.I,
        ):
            return cleaned
        # 没有法定后缀时，仅接受明显的中文企业主体词；个人姓名/职位不接受。
        if re.search(r"(?:科技|贸易|商贸|电商|实业|工业|制造|企业|工厂|商贸行|商行)$", cleaned):
            return cleaned
        self._log("字段Agent返回的客户无法确认是公司主体，已忽略并转人工", "warning")
        return ""

    def _extract_company_from_body(self, body: str) -> str:
        """从正文中提取公司名"""
        if not body:
            return ""
        lines = body.split("\n")

        # “公司：”经常单独占一行，下一行才是主体名称。必须优先处理
        # 这种明确标签，不能先把上一句“麻烦安排以下公司”截成假客户。
        for index, raw_line in enumerate(lines):
            line = raw_line.strip()
            label = re.match(
                r"^(?:公司名称|公司名|客户公司|申请单位|公司)\s*[:：]?\s*(.*)$",
                line,
                re.I,
            )
            if not label:
                continue
            value = label.group(1).strip()
            if not value:
                for next_line in lines[index + 1:]:
                    value = next_line.strip()
                    if value:
                        break
            candidate = self._company_substring(value)
            if candidate and not self._customer_noise_reason(candidate):
                return candidate

        # 搜索含公司名特征的行。保留全部候选后取最长的法定主体，避免
        # “麻烦安排以下公司”示例cd2a9eae有限公司名称。
        candidates = []
        lines = body.split("\n")
        for line in lines:
            line = line.strip()
            if any(kw in line for kw in ["公司", "有限", "科技", "电商", "贸易", "实业"]):
                # 排除代理名和邮件签名
                if (
                    not self._is_agent_name(line)
                    and self._customer_noise_reason(line) != "发件方说明句"
                    and len(line) < 50
                ):
                    # 清理行内多余内容，并复用统一的拒绝闸门；否则
                    # “1家公司/不含回收公司”会遮住后面真正的主体名称。
                    candidate = self._company_substring(line)
                    if candidate and not self._customer_noise_reason(candidate):
                        candidates.append(candidate)
        if candidates:
            return max(candidates, key=len)
        return ""

    def _extract_company_from_filename(self, filename: str) -> str:
        """从附件文件名中提取公司名"""
        if not filename:
            return ""
        # 去除扩展名
        name = os.path.splitext(filename)[0] if "." in filename else filename
        # 统一走公司后缀解析，支持中文公司名、英文 LLC/LIMITED、
        # 西语/波兰语等带重音字符的法定后缀。
        return self._company_substring(name)

    def _extract_company_from_attachment_text(self, text: str) -> str:
        """从附件文本中提取公司名"""
        if not text:
            return ""
        # 搜索"公司名称"标签后的内容
        patterns = [
            r"公司名称[:\s：]*([^\n\r，,]{2,30})",
            r"甲方[:\s：]*([^\n\r，,]{2,30})",
            r"申请单位[:\s：]*([^\n\r，,]{2,30})",
            r"(?:company\s*name|company|applicant)\s*[:：]?\s*([^\n\r,，;；]{2,80})",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.I)
            if match:
                name = match.group(1).strip()
                candidate = self._company_substring(name)
                if candidate:
                    return candidate
        # 兜底: 无标签的裸公司名（OCR 截图常见, 如订单确认截图里直接印公司名）
        # 取最长匹配（避免截到公司名的前缀子串）
        bare = re.findall(
            r"[\u4e00-\u9fa5A-Za-z0-9（）()]{2,40}"
            r"(?:(?:(?:有限|股份)责任?|责任)?公司|（个体工商户）|\(个体工商户\)|"
            r"经营部|门市部|服务部|商店|商贸行|商行|工厂|工作室|店)",
            text,
        )
        # 排除表单说明文字(如香港公司注册表里的
        # "適用於只有一名成員而該成員同時亦是唯一董事的私人公司"、
        # "境外公司需提供注册证书及含法人信息和公司"、"(邮箱只能对应注册一个公司")
        bare = [b for b in bare if not any(h in b for h in self._DISCLAIMER_HINTS)]
        if bare:
            return max(bare, key=len)
        return ""

    def _is_agent_name(self, text: str) -> bool:
        """判断文本是否是代理名(而非公司名)"""
        return text.strip() in self._agent_alias_index

    def _strip_known_agent_prefix(self, value: str) -> str:
        """剥离粘在公司名前的已知代理别名。

        批量邮件有时写成 ``TBA示例b691064e有限公司``，而不是
        ``TBA+示例b691064e有限公司``。代理表已经提供了可靠的别名
        集合，因此只在候选以已知别名开头、且剩余部分本身通过公司主体规则
        时剥离；不会对未知前缀做猜测，也不会吞掉真实公司名。
        """
        raw = str(value or "").strip()
        if not raw or not self._agent_alias_index:
            return raw
        aliases = sorted(self._agent_alias_index, key=len, reverse=True)
        for alias in aliases:
            if len(raw) <= len(alias) or raw[:len(alias)].casefold() != alias.casefold():
                continue
            remainder = raw[len(alias):].lstrip(" \t-—_:：+＋/／")
            if remainder and self._company_fragment_is_valid(remainder):
                return remainder.strip(" .,-")
        return raw

    @classmethod
    def _company_fragment_is_valid(cls, value: str) -> bool:
        """判断代理别名后的剩余片段是否具备公司主体证据。"""
        text = str(value or "").strip()
        if len(text) < 4:
            return False
        if re.search(
            r"(?:有限责任公司|股份有限公司|集团有限公司|有限公司|责任公司|公司(?!名称|名)|"
            r"经营部|门市部|服务部|商店|商贸行|商行|工厂|工作室|店)$",
            text,
        ):
            return True
        if re.match(
            r"(?i)^(?:SIA|AG|SA|SARL)\s+[A-Za-zÀ-ÖØ-öø-ÿĄĆĘŁŃÓŚŹŻąćęłńóśźż]",
            text,
        ):
            return True
        return bool(re.search(rf"{cls._COMPANY_EN_SUFFIX}\s*$", text, re.I))

    _CLIENT_CODE_RE = re.compile(
        # 同时覆盖 EG3164、BG1945 和 K-DED0943 这类带前缀的案件编号。
        r"(?<![A-Za-z0-9])((?:[A-Za-z]{1,3}\s*[-_]\s*)?"
        r"[A-Za-z]{1,8}\s*[-_]?\s*\d{2,8})(?![A-Za-z0-9])"
    )
    _SERVICE_OR_CERT_RE = re.compile(
        r"(?:WEEE|(?<!W)EEE|电池(?:法)?|電池(?:法)?|包装(?:法)?|包裝(?:法)?|一次性塑料|EPR)",
        re.I,
    )

    @classmethod
    def _client_codes_in(cls, text: str) -> List[str]:
        """抽取 EG3164 / BG1945 一类客户或案件编号，保序去重。"""
        found: List[str] = []
        for match in cls._CLIENT_CODE_RE.finditer(text or ""):
            code = re.sub(r"[\s_-]+", "", match.group(1)).upper()
            if code not in found:
                found.append(code)
        return found

    @classmethod
    def _client_code_near_company(cls, text: str, company: str) -> str:
        """提取紧邻公司名称前的客户编号，避免把同封另一项目的编号错配。"""
        if not text or not company:
            return ""
        position = text.lower().find(str(company).lower())
        if position < 0:
            return ""
        # 只看公司前的局部窗口，并取最后一个编号：
        # `德国WEEE EG3164 RONG FANG ...` → EG3164。
        window = text[max(0, position - 100):position]
        codes = cls._client_codes_in(window)
        return codes[-1] if codes else ""

    @classmethod
    def _strip_company_prefix_noise(cls, candidate: str) -> tuple[str, int, str]:
        """剥离英文公司候选前的项目/证书类型和客户编号。

        `德国WEEE EG0001 Example Trading LIMITED` 中，德国WEEE 是
        项目/证书类型，EG3164 是客户编号，只有 RONG FANG... 才是公司。
        返回清洗后的名称、相对原候选的起点偏移和识别出的编号。
        """
        raw = str(candidate or "")
        code_match = cls._CLIENT_CODE_RE.search(raw)
        if not code_match:
            return raw.strip(" .,-"), 0, ""
        prefix = raw[:code_match.start()]
        # 仅当编号位于候选起始处，或其前方是项目/证书词时才清理；公司名中
        # 偶然出现的字母数字组合不会被误删。
        if prefix.strip() and not cls._SERVICE_OR_CERT_RE.search(prefix):
            # 编号位于候选最前方时，即使前面没有 WEEE/证书词，也要剥离。
            # 但必须看到一个明确的公司法定后缀，避免误删合法的字母数字公司名。
            if code_match.start() != 0:
                return raw.strip(" .,-"), 0, ""
        remainder = raw[code_match.end():]
        leading = len(remainder) - len(remainder.lstrip(" \t-—_:："))
        start = code_match.end() + leading
        cleaned = remainder.lstrip(" \t-—_:：").strip(" .,-")
        if not prefix.strip() and not cls._company_fragment_is_valid(cleaned):
            return raw.strip(" .,-"), 0, ""
        if prefix.strip() and not cls._SERVICE_OR_CERT_RE.search(prefix):
            if not cls._company_fragment_is_valid(cleaned):
                return raw.strip(" .,-"), 0, ""
        code = re.sub(r"[\s_-]+", "", code_match.group(1)).upper()
        return cleaned, start, code

    @staticmethod
    def _split_top_level(text: str, separators: str = "+＋") -> List[str]:
        """只在括号外分段，保留 `（大型设备+小型设备）` 的内部加号。"""
        if not text:
            return []
        openers = "(（[【"
        closers = ")）]】"
        depth = 0
        current: List[str] = []
        parts: List[str] = []
        for char in text:
            if char in openers:
                depth += 1
            elif char in closers and depth:
                depth -= 1
            if char in separators and depth == 0:
                value = "".join(current).strip()
                if value:
                    parts.append(value)
                current = []
            else:
                current.append(char)
        value = "".join(current).strip()
        if value:
            parts.append(value)
        return parts

    def _company_substring(self, text: str) -> str:
        """从一段文本中提取公司名子串(而非整段)。
        中文: 后缀锚定(…有限公司/…公司); 英文: 词边界后缀锚定(LTD/GMBH/B.V.…)
        排除代理名与表单说明文字。"""
        if not text:
            return ""

        # 批量清单常以 BG1892、EG3107 等客户编号开头。编号不是公司名；
        # 题目里还会出现“德国WEEE EG3164 公司”的证书/编号前缀，后续英文
        # 候选清洗会按实体角色剥离，不再把它们写入客户列。
        text = re.sub(r"^\s*[A-Za-z]{1,8}[-_]?\d{2,}\s+", "", text)

        def _ok(name: str) -> bool:
            name = self._strip_known_agent_prefix(name)
            return (
                len(name) >= 4
                and not self._is_agent_name(name)
                and not any(h in name for h in self._DISCLAIMER_HINTS)
                and not re.fullmatch(r"[0-9一二三四五六七八九十多几]*\s*家\s*(?:公司|主体|企业)", name, re.I)
                and not re.fullmatch(r"(?:不含|包含|含有|无|有)?回收公司", name)
            )

        zh = re.findall(
            # 使用贪心主体，避免“示例23f9e61e有限公司”被惰性匹配截成仅“有限公司”。
            r"[\u4e00-\u9fa5A-Za-z0-9·（）() \t]{2,30}"
            r"(?:有限责任公司|股份有限公司|集团有限公司|有限公司|责任公司|公司(?!名称|名)|"
            r"经营部|门市部|服务部|商店|商贸行|商行|工厂|工作室|店)"
            r"(?:（个体工商户）|\(个体工商户\))?",
            text,
        )
        zh = [self._clean_company_candidate(self._strip_company_prefix_noise(c)[0]) for c in zh]
        zh = [self._strip_known_agent_prefix(c) for c in zh]
        zh = [c for c in zh if _ok(c)]
        if zh:
            return max(zh, key=len)

        en_pattern = (
            rf"([{self._COMPANY_EN_LETTER_CHARS}][{self._COMPANY_EN_CHARS}]{{2,80}}"
            rf"{self._COMPANY_EN_SUFFIX})(?![A-Za-z])"
        )
        # 同一名称可能同时含有 ``CORP`` 与 ``S.R.O.`` 等后缀。
        # 选择最长命中，避免惰性正则在前一个短后缀处提前截断。
        en_matches = list(re.finditer(en_pattern, text, re.I))
        if en_matches:
            m = max(en_matches, key=lambda item: len(item.group(1)))
            candidate, _, _ = self._strip_company_prefix_noise(m.group(1))
            candidate = self._clean_company_candidate(candidate)
            candidate = self._strip_known_agent_prefix(candidate)
            if _ok(candidate):
                return candidate

        # 部分欧洲公司名只有法定前缀，没有 LTD/GMBH 等后缀，例如
        # `SIA Andistef`。这类候选只能在明确的法律前缀后读取，不能把
        # 普通英文姓名或说明句泛化成公司名。
        prefix = re.search(
            rf"\b(?:SIA|AG|SA|SARL)\s+[{self._COMPANY_EN_LETTER_CHARS}]"
            rf"[{self._COMPANY_EN_CHARS}]{{2,80}}",
            text,
            re.I,
        )
        if prefix:
            candidate = prefix.group(0).strip(" .,-")
            candidate = self._strip_known_agent_prefix(candidate)
            if _ok(candidate):
                return candidate
        return ""

    def _company_substrings(self, text: str) -> List[str]:
        """提取一段文本中的全部公司名，保留出现顺序并去除重叠候选。

        ``_company_substring`` 只返回最长的一个，适合单公司主题；批量邮件
        需要保留 ``公司A、公司B`` 或 ``项目A 公司A & 项目B 公司B`` 的全部
        客户，不能再用单值接口。
        """
        if not text:
            return []

        text = re.sub(r"^\s*[A-Za-z]{1,8}[-_]?\d{2,}\s+", "", text)
        matches = []
        zh_pattern = (
            r"[\u4e00-\u9fa5A-Za-z0-9·（）() \t]{2,30}?"
            r"(?:有限责任公司|股份有限公司|集团有限公司|有限公司|责任公司|公司|"
            r"经营部|门市部|服务部|商店|商贸行|商行|工厂|工作室|店)"
            r"(?:（个体工商户）|\(个体工商户\))?"
        )
        for match in re.finditer(zh_pattern, text):
            value = match.group(0).strip(" ,，;；+＋&＆")
            if value:
                cleaned = self._clean_company_candidate(self._strip_company_prefix_noise(value)[0])
                cleaned = self._strip_known_agent_prefix(cleaned)
                offset = match.start() + (len(value) - len(cleaned)) if cleaned != value else match.start()
                matches.append((offset, match.end(), cleaned))

        en_pattern = (
            rf"[{self._COMPANY_EN_LETTER_CHARS}][{self._COMPANY_EN_CHARS}]{{2,80}}"
            rf"{self._COMPANY_EN_SUFFIX}(?![A-Za-z])"
        )
        for match in re.finditer(en_pattern, text, re.I):
            raw_value = match.group(1).strip(" .,-") if match.lastindex else match.group(0).strip(" .,-")
            value, offset, _ = self._strip_company_prefix_noise(raw_value)
            value = self._clean_company_candidate(value)
            value = self._strip_known_agent_prefix(value)
            if value:
                matches.append((match.start() + offset, match.end(), value))

        prefix_pattern = (
            rf"\b(?:SIA|AG|SA|SARL)\s+[{self._COMPANY_EN_LETTER_CHARS}]"
            rf"[{self._COMPANY_EN_CHARS}]{{2,80}}"
        )
        for match in re.finditer(prefix_pattern, text, re.I):
            value = self._strip_known_agent_prefix(match.group(0).strip(" .,-"))
            if value:
                matches.append((match.start(), match.end(), value))

        # 先按原文位置，再按长度降序；同一位置的短子串被长候选覆盖。
        matches.sort(key=lambda item: (item[0], -(item[1] - item[0])))
        selected = []
        seen = set()
        for start, end, value in matches:
            if len(value) < 4 or self._is_agent_name(value):
                continue
            if self._customer_noise_reason(value):
                continue
            if re.fullmatch(r"[0-9一二三四五六七八九十多几]*\s*家\s*(?:公司|主体|企业)", value, re.I):
                continue
            if re.fullmatch(r"(?:不含|包含|含有|无|有)?回收公司", value):
                continue
            if any(start >= old_start and end <= old_end for old_start, old_end, _ in selected):
                continue
            norm = normalize_text(value)
            if not norm or norm in seen:
                continue
            if any(not (end <= old_start or start >= old_end) for old_start, old_end, _ in selected):
                continue
            selected.append((start, end, value))
            seen.add(norm)
        selected.sort(key=lambda item: item[0])
        return [value for _, _, value in selected]

    @staticmethod
    def _clean_company_candidate(value: str) -> str:
        """从公司后缀正则命中的整句中剥离业务动作前缀。

        例如“示例81b28bc8有限公司”只能把后半段作为客户，
        不能把整句或仅剩“有限公司”写入客户字段。
        """
        text = str(value or "").strip(" \t\r\n,，;；:+＋-—")
        text = re.sub(
            r"^.*?(?:申请撤销|正式撤销|申请注销|正式注销|撤销|注销|变更|修改)\s*",
            "",
            text,
            flags=re.I,
        )
        text = re.sub(
            r"^(?:您好[，,：:]?\s*)?(?:公司名称|公司名|客户公司|附件(?:中的?)?公司|公司)\s*[:：]?\s*",
            "",
            text,
            flags=re.I,
        )
        return text.strip(" \t\r\n,，;；:+＋-—")

    def _structured_groups(self, subject: str, body: str) -> List[dict]:
        """解析主题/正文中的结构化「公司+项目」清单行。

        覆盖两类高频格式:
          1) 代理群发邮件正文逐行列出多家公司:
             "TBA+Fixture 5b2a6d2f Limited+爱尔兰包装法
              TBA+PLConcepts GmbH & Co KG+爱尔兰包装法 ..."
          2) 主题式单公司多项目:
             "示例代理庚+示例bdd21359有限公司+西班牙包装法+法国包装法+意大利包装法"
             "示例代理己——德国WEEE EG3107 示例ddb56611有限公司+德国电池 BG1878 示例客户..."

        做法: 把文本按 "+" 切段, 逐段识别 公司子串 与 国家×业务组合,
        组合归属最近的 公司段; 同一公司多条记录合并。
        返回: [{"customer","source","projects":[标准化项目名]}]; 无结构化记录时返回 []
        """
        groups: Dict[str, dict] = {}
        # 邮件纯文本有时仍含 &nbsp;，先还原用于分组；原文证据保持不变。
        subject, body = html.unescape(subject), html.unescape(body)

        for source, text in (("主题", subject), ("正文", body)):
            if not text or ("+" not in text and "＋" not in text):
                continue
            # 换行是记录边界，不能压成普通空格，否则上一行的“德国包装法”
            # 会粘到下一行公司名前，形成额外的伪申请人。
            norm = "+".join(" ".join(line.split()) for line in text.splitlines() if line.strip())
            # `德国WEEE（大型设备+小型设备）` 的括号内加号属于规格，不是
            # 公司/项目分隔符；只对括号外的 + 做结构化分段。
            segs = self._split_top_level(norm)

            pending_company = ""
            pending_projects: List[str] = []
            pending_project_codes: Dict[str, List[str]] = {}
            pending_company_codes: List[str] = []
            # 连续分段经常写成「荷兰WEEE + 包装法」。后一段没有国家，
            # 需要继承前一业务段的国家，而不是丢掉包装法。
            active_country = ""
            records: List[tuple] = []

            def _flush():
                nonlocal pending_company, pending_projects, pending_project_codes, pending_company_codes, active_country
                if pending_company and pending_projects:
                    records.append((
                        pending_company,
                        list(pending_projects),
                        {name: list(codes) for name, codes in pending_project_codes.items()},
                    ))
                pending_company = ""
                pending_projects = []
                pending_project_codes = {}
                pending_company_codes = []
                active_country = ""

            for seg in segs:
                seg = seg.strip()
                if not seg:
                    continue
                combos = [
                    p["standard_name"] for p in self._extract_projects_by_rules(seg)
                ]
                countries = self._country_names_in(seg)
                if countries:
                    active_country = countries[-1]
                elif pending_company and active_country and not combos:
                    combos = [
                        p["standard_name"]
                        for p in self._extract_projects_by_rules(active_country + seg)
                    ]
                company = self._company_substring(seg)
                # 公司后缀正则可能从说明句中截出“24告知以下2家公司”一类
                # 假主体；这类候选不能参与公司—项目分组，否则会把一封邮件
                # 错误展开成额外业务明细。英文公司后缀(Inc/LLC 等)由噪声闸门
                # 明确放行，真实英文公司不会再被“疑似联系人姓名”误删。
                if company and self._customer_noise_reason(company):
                    company = ""
                codes = self._client_codes_in(seg)
                if company:
                    inherited_country = active_country if (
                        pending_company and normalize_text(company) == normalize_text(pending_company)
                    ) else ""
                    # 新公司段: 若当前记录已凑齐(公司+组合)则先落账
                    if pending_company and pending_projects:
                        _flush()
                    active_country = countries[-1] if countries else inherited_country
                    pending_company = company
                    pending_company_codes = list(codes)
                    # 段内自含组合 (如 "德国WEEE EG3107 示例ddb56611有限公司")
                    pending_projects = combos
                    pending_project_codes = {
                        project: list(codes) for project in combos if codes
                    }
                elif combos and pending_company:
                    for c in combos:
                        if c not in pending_projects:
                            pending_projects.append(c)
                        local_codes = codes or pending_company_codes
                        if local_codes:
                            pending_project_codes.setdefault(c, [])
                            for code in local_codes:
                                if code not in pending_project_codes[c]:
                                    pending_project_codes[c].append(code)
                # 其他段(代理名/品牌/流水号)忽略
            _flush()

            for company, projs, project_codes in records:
                key = company
                if key in groups:
                    for p in projs:
                        if p not in groups[key]["projects"]:
                            groups[key]["projects"].append(p)
                    existing_codes = groups[key].setdefault("project_codes", {})
                    for project, codes in project_codes.items():
                        bucket = existing_codes.setdefault(project, [])
                        for code in codes:
                            if code not in bucket:
                                bucket.append(code)
                else:
                    groups[key] = {
                        "customer": company,
                        "source": f"{source}结构化",
                        "projects": list(projs),
                        "project_codes": project_codes,
                    }

        # 泛称“国家EPR”只在没有更具体业务时保留。主题/正文同时出现
        # “意大利WEEE/电池法/包装法”时，正文里单独出现的“意大利EPR”
        # 只是概括语，不能再生成额外业务明细。
        specific_countries = {
            country
            for group in groups.values()
            for project in group.get("projects") or []
            for country in COUNTRIES
            if project == country + "EPR" or project.startswith(country)
            if project != country + "EPR"
        }
        for key in list(groups):
            group = groups[key]
            filtered = []
            for project in group.get("projects") or []:
                if project.endswith("EPR") and any(
                    project == country + "EPR" for country in specific_countries
                ):
                    continue
                if project not in filtered:
                    filtered.append(project)
            group["projects"] = filtered
            if not filtered:
                groups.pop(key, None)

        # 另一种真实邮件格式不使用“+”作为行分隔，而是使用制表符、&、
        # 中文顿号或换行，例如：
        #   德国电池法 BG2002 公司A & 德国WEEE EG3205 公司A
        #   公司A、公司B申请注册爱尔兰包装法
        # 旧逻辑会回退到单客户提取，导致第二家公司消失。这里先按分隔符
        # 形成局部窗口；若公司与项目无法一一对应，则保守地为每家公司保留
        # 全部声明项目，并标记需人工确认，优先保证不漏掉客户/项目。
        for source, text in (("主题", subject), ("正文", body)):
            if not text:
                continue
            delimiters = r"[\r\n\t&＆；;、，,]"
            if not re.search(delimiters, text):
                continue

            chunks = [c.strip() for c in re.split(delimiters, text) if c.strip()]
            if len(chunks) < 2:
                continue

            all_companies = self._company_substrings(text)
            # 说明句中也会出现一个带“公司”后缀的代理/发件方名称，
            # 例如“示例9ad3384d有限公司……提交……名单”。
            # 该名称不是客户主体，不能参与批量公司—项目映射。
            all_companies = [
                company for company in all_companies
                if not self._customer_noise_reason(company)
            ]
            global_projects = [
                p["standard_name"] for p in self._extract_projects_by_rules(text)
            ]
            if len(all_companies) < 2 or not global_projects:
                continue

            local = {}
            for chunk in chunks:
                chunk_companies = self._company_substrings(chunk)
                chunk_companies = [
                    company for company in chunk_companies
                    if not self._customer_noise_reason(company)
                ]
                chunk_projects = [
                    p["standard_name"] for p in self._extract_projects_by_rules(chunk)
                ]
                for company in chunk_companies:
                    local.setdefault(company, [])
                    for project in chunk_projects:
                        if project not in local[company]:
                            local[company].append(project)

            mapped_companies = set(local)
            ambiguous = False
            if not local or mapped_companies != set(all_companies):
                ambiguous = True
                for company in all_companies:
                    local.setdefault(company, list(global_projects))
            elif len(local) > 1 and any(not projects for projects in local.values()):
                ambiguous = True
                for company in local:
                    if not local[company]:
                        local[company] = list(global_projects)

            for company, projects in local.items():
                if not projects:
                    continue
                entry = groups.setdefault(
                    company,
                    {
                        "customer": company,
                        "source": f"{source}结构化",
                        "projects": [],
                    },
                )
                for project in projects:
                    if project not in entry["projects"]:
                        entry["projects"].append(project)
                if ambiguous:
                    entry["needs_review"] = True
                    entry["source"] = f"{source}结构化(公司-项目关联待确认)"

        # 另一类常见群发格式没有“+”分隔符，而是以客户编号逐行列出公司：
        #   BG1892 甲公司
        #   BG1893 乙公司
        # 项目只在主题中声明一次。旧逻辑会回退到普通客户提取，最终只取首家公司。
        # 仅在主题能唯一确定一个项目、且正文至少有两行编号+公司时启用，避免把
        # 普通正文中零散提及的多家公司误扩展成多个工单。
        subject_projects = [
            p["standard_name"] for p in self._extract_projects_by_rules(subject)
        ]
        subject_projects = list(dict.fromkeys(subject_projects))
        if len(subject_projects) == 1:
            # 无编号的逐行主体清单也可以继承唯一主题项目。只接受整行
            # 公司名，不把正文叙述、联系人或常见签名区域中的公司扩成工单。
            plain_companies = []
            for raw_line in body.splitlines():
                line = raw_line.strip()
                if re.match(
                    r"^(?:--+|此致|致礼|顺祝|祝好|祝商祺|best\s+regards|kind\s+regards|"
                    r"regards\b|thanks\b|sent\s+from\b|发件人[:：]|from[:：])",
                    line, re.I,
                ):
                    break
                line = re.sub(r"^\s*(?:\d{1,3}[.、)）]|[（(]\d{1,3}[)）])\s*", "", line)
                company = self._company_substring(line)
                if (company and not self._customer_noise_reason(company)
                        and normalize_text(company) == normalize_text(line)
                        and company not in plain_companies):
                    plain_companies.append(company)
            if len(plain_companies) >= 2:
                for company in plain_companies:
                    entry = groups.setdefault(company, {
                        "customer": company, "source": "正文批量清单",
                        "projects": [], "needs_review": False,
                    })
                    if subject_projects[0] not in entry["projects"]:
                        entry["projects"].append(subject_projects[0])

            numbered_companies: List[tuple[str, bool]] = []
            for line in body.splitlines():
                line = line.strip()
                code_match = re.match(r"^[A-Za-z]{1,8}[-_]?\d{2,}\s+(.+?)\s*$", line)
                if not code_match:
                    continue
                company = self._company_substring(line)
                needs_review = False
                if not company:
                    # 编号清单中也会出现品牌名、个体商号等没有 Ltd/公司后缀的名称。
                    # 在“唯一主题项目 + 至少两条编号行”的严格前提下，保留原名称，
                    # 但显式标记人工确认，而不是静默漏掉该客户。
                    tail = code_match.group(1).strip(" ,;，；")
                    if (
                        4 <= len(tail) <= 80
                        and re.fullmatch(r"[A-Za-z0-9 .,&'’()\-]+", tail)
                    ):
                        company = tail
                        needs_review = True
                if company and company not in {name for name, _ in numbered_companies}:
                    numbered_companies.append((company, needs_review))
            if len(numbered_companies) >= 2:
                return [
                    {
                        "customer": company,
                        "source": "正文批量清单(名称待确认)" if needs_review else "正文批量清单",
                        "projects": list(subject_projects),
                        "needs_review": needs_review,
                    }
                    for company, needs_review in numbered_companies
                ]
        return list(groups.values())

    def _project_coverage_gap(self, subject: str, body: str, project_result: dict) -> List[str]:
        """检查主题/正文明确声明的项目是否全部进入输出。

        仅比较邮件正文和主题，不把旧式申请表整张国家×业务网格算入，避免
        附件模板文字造成假漏项。多公司关联不明确时该检查只负责阻止静默漏项，
        不擅自把项目分配给错误的公司。
        """
        if project_result.get("epr_form") is not None:
            return []
        declared = {
            p.get("standard_name")
            for p in self._drop_generic_epr(
                self._extract_projects_by_rules(subject + " " + body)
            )
            if p.get("standard_name")
        }
        emitted = {
            p.get("standard_name")
            for p in project_result.get("projects") or []
            if p.get("standard_name")
        }
        for group in project_result.get("groups") or []:
            emitted.update(str(p) for p in group.get("projects") or [] if p)
        return sorted(declared - emitted)

    def _application_record_groups(self, attachments, subject, body, text_groups):
        """按附件成员及申请人保留表单，无复选框的申请表也必须参与。

        同一工作表的中英文名是同一申请人的候选；不同表单不互相竞争。
        仅使用该表勾选、局部文件名、明确正文关系或唯一邮件项目，
        不扫描申请表里的国家业务模板网格来制造项目。
        """
        declared = list(dict.fromkeys(
            item["standard_name"] for item in self._drop_generic_epr(
                self._extract_projects_by_rules(subject + " " + body)
            )
        ))
        output = []
        covered_aliases = set()
        priority = {"company_name_zh": 0, "company_name": 1, "applicant_name": 2, "company_name_en": 3}
        for attachment_index, attachment in enumerate(attachments or []):
            buckets = {}
            for record in attachment.get("structured_records") or []:
                if not isinstance(record, dict) or record.get("record_type") != "epr_application":
                    continue
                member = str(record.get("attachment_name") or attachment.get("filename") or "").replace("\\", "/")
                sheet = str(record.get("sheet_name") or "")
                buckets.setdefault((member, sheet), []).append(record)
            # 相同申请人可能在注册信息页、产品页重复填写；仅在已有名称
            # 明确重合时合并，不能仅因位于同一个 ZIP 就合并主体。
            applicants = []
            for (member, sheet), records in buckets.items():
                candidates = []
                seen = set()
                for record in sorted(records, key=lambda r: priority.get(r.get("company_field"), 5)):
                    clean = self._sanitize_customer_result({
                        "customer": record.get("customer", ""), "source": "附件EPR申请表格式化字段",
                    }).get("customer", "")
                    key = normalize_text(clean)
                    if not key or key in seen:
                        continue
                    seen.add(key)
                    candidates.append({
                        "id": f"epr-company-{len(candidates) + 1}", "value": clean,
                        "field_type": record.get("company_field") or "company_name",
                        "source": member,
                        "evidence": str(record.get("raw_text") or record.get("field_label") or "")[:300],
                    })
                # 资料清单中的“针对某业务”等说明即使误被上游识别为字段，
                # 也不能证明这一页存在一个申请人。真正无主体的勾选表由下面保留。
                if not candidates:
                    continue
                covered_aliases.update(seen)
                existing = next((entry for entry in applicants
                    if entry["member"] == member and seen.intersection(entry["aliases"])), None)
                if existing:
                    existing["sheets"].add(sheet)
                    existing["aliases"].update(seen)
                    known = {normalize_text(item["value"]) for item in existing["candidates"]}
                    existing["candidates"].extend(item for item in candidates if normalize_text(item["value"]) not in known)
                    existing["candidates"].sort(key=lambda item: priority.get(item["field_type"], 5))
                else:
                    applicants.append({"member": member, "sheets": {sheet}, "aliases": seen, "candidates": candidates})

            forms = [form for form in attachment.get("epr_forms") or [] if isinstance(form, dict)]
            used_forms = set()
            for applicant in applicants:
                member = applicant["member"]
                candidates = applicant["candidates"]
                peers = [item for item in applicants if item["member"] == member]
                matching_forms = []
                for index, form in enumerate(forms):
                    form_member = str(form.get("filename") or attachment.get("filename") or "").replace("\\", "/")
                    if form_member != member:
                        continue
                    # 一个工作簿若有多个申请人，必须能定位到同一工作表。
                    form_sheet = str(form.get("sheet_name") or form.get("sheet") or "")
                    # sheetN.xml 是内部文件编号，不是工作表显示顺序。
                    # 缺少实际名称映射时不能用数组下标把项目绑定给另一个申请人。
                    if len(peers) == 1 or form_sheet in applicant["sheets"]:
                        matching_forms.append(form)
                        used_forms.add(index)
                review = not bool(candidates)
                names = []
                if matching_forms:
                    for form in matching_forms:
                        checked = list(dict.fromkeys(form.get("projects") or []))
                        if (declared and len(checked) > len(declared)
                                and set(declared).issubset(checked)
                                and re.search(r"泛欧|8\s*国|pan[- ]?epr|pan[- ]?europe", member, re.I)):
                            checked = [name for name in checked if name in declared]
                        names.extend(checked)
                        review = review or bool(form.get("unmatched_countries") or form.get("orphan_business"))
                else:
                    # 先取已明确建立的该申请人关系，再取其文件名的项目证据。
                    for group in text_groups or []:
                        if normalize_text(group.get("customer", "")) in applicant["aliases"]:
                            names.extend(group.get("projects") or [])
                            review = review or bool(group.get("needs_review"))
                    if not names:
                        names = [item["standard_name"] for item in self._drop_generic_epr(
                            self._extract_projects_by_rules(member)
                        ) if not item["standard_name"].endswith("EPR")]
                    if not names and len(declared) == 1:
                        names = list(declared)
                names = list(dict.fromkeys(names))
                for index, candidate in enumerate(candidates, 1):
                    candidate["id"] = f"epr-company-{index}"
                sheets = "、".join(sorted(name for name in applicant["sheets"] if name))
                output.append({
                    "customer": candidates[0]["value"] if candidates else "",
                    "source": "附件EPR申请表格式化字段", "projects": names,
                    "company_candidates": candidates if len(candidates) > 1 else [],
                    "attachment_record_source": f"附件EPR申请表：{attachment_index + 1}:{member}" + (f" / {sheets}" if sheets else ""),
                    "needs_review": review or not names or len(candidates) > 1,
                })
            # 缺少主体字段的勾选表仍保留待确认，不因为新分组吞掉该表。
            for index, form in enumerate(forms):
                if index not in used_forms:
                    group = self._epr_form_applicant_group(form, [attachment])
                    form_member = str(form.get("filename") or attachment.get("filename") or "").replace("\\", "/")
                    if sum(item["member"] == form_member for item in applicants) > 1:
                        # 已知有多个独立申请人，不能把“无法定位工作表”
                        # 降级成从这些申请人中擅自挑选第一个。
                        group["customer"] = ""
                        group["company_candidates"] = []
                        group["source"] = "待确认"
                    group["projects"] = list(dict.fromkeys(form.get("projects") or []))
                    group["needs_review"] = True
                    output.append(group)
        if output:
            for group in text_groups or []:
                company = str(group.get("customer") or "")
                key = normalize_text(company)
                if (key and key not in covered_aliases and group.get("projects")
                        and str(group.get("source") or "").startswith("正文")
                        and not group.get("needs_review")
                        and not self._customer_noise_reason(company)):
                    # 附件只覆盖 A/B 时，正文明确列出的 C 仍保留原有关系。
                    output.append(dict(group))
                    covered_aliases.add(key)
        return output

    def _extract_projects(self, subject: str, body: str, attachments: list,
                          groups: List[dict] = None) -> dict:
        """
        提取项目。

        优先级 P0 — EPR 申请表勾选 (唯一权威):
            附件含「可勾选的泛欧 EPR 申请表」时, **只**取客户打勾的
            国家×业务组合, 其他一律不管 (未勾选的国家不查、也不做文本扫描)。
            若「国家勾了但业务没勾」「业务勾了但国家没勾」→ need_review=True
            (置信度落到"需人工确认", 由人工补全, 不静默跳过)。

        回落 P1 — 文本规则 (保持原逻辑):
            附件没有可勾选申请表(如德国ECOPV服务信息申请表/意大利EPR申请表),
            仍按 附件四项目名扫描 → 离线"国家+业务"正则 提取, 宁可多提不漏提。
        """
        search_text = subject + " " + body
        for att in attachments:
            search_text += " " + att.get("text_content", "")
            search_text += " " + att.get("filename", "")

        # ---------- P0: 多份 EPR 申请表逐表绑定主体和勾选项目 ----------
        all_forms = [
            form
            for attachment in attachments or []
            for form in (attachment.get("epr_forms") or [])
            if isinstance(form, dict)
        ]
        application_groups = self._application_record_groups(
            attachments, subject, body, self._structured_groups(subject, body)
        )
        if len(application_groups) > 1:
            names = list(dict.fromkeys(
                name for group in application_groups for name in group.get("projects") or []
            ))
            return {
                "projects": [{"raw_value": name, "standard_name": name} for name in names],
                "epr_form": all_forms[0] if all_forms else None,
                "epr_forms_count": len(all_forms), "epr_form_groups": True,
                "need_review": any(group.get("needs_review") for group in application_groups),
                "groups": application_groups,
            }
        if len(all_forms) > 1:
            form_groups = []
            flattened_projects = []
            for form in all_forms:
                group = self._epr_form_applicant_group(form, attachments)
                names = list(dict.fromkeys(
                    str(name).strip() for name in (form.get("projects") or [])
                    if str(name).strip()
                ))
                group["projects"] = names
                group["needs_review"] = bool(
                    group.get("needs_review")
                    or form.get("unmatched_countries")
                    or form.get("orphan_business")
                    or not names
                )
                form_groups.append(group)
                flattened_projects.extend(names)
            flattened_projects = list(dict.fromkeys(flattened_projects))
            self._log(
                f"检测到 {len(all_forms)} 份 EPR 申请表，按成员文件逐份绑定: "
                f"{[(g.get('customer'), g.get('projects')) for g in form_groups]}"
            )
            return {
                "projects": [{"raw_value": f"申请表勾选: {name}", "standard_name": name}
                             for name in flattened_projects],
                "epr_form": all_forms[0],
                "epr_forms_count": len(all_forms),
                "epr_form_groups": True,
                "need_review": any(group.get("needs_review") for group in form_groups),
                "groups": form_groups,
            }

        # ---------- P0: EPR 申请表勾选优先 ----------
        form = self._pick_epr_form(attachments)
        if form is not None:
            projects = [
                {"raw_value": f"申请表勾选: {name}", "standard_name": name}
                for name in form.get("projects") or []
            ]
            # 泛欧 8 国模板经常由代理复制上一单，旧国家/业务复选框会被
            # 一并保留。若邮件主题已经明确只声明一个具体项目，则该声明
            # 是本封邮件的业务范围闸门；仅对“泛欧/8国”模板启用，避免改变
            # 其它自定义申请表中“主题写一个、表里勾多个”的既有规则。
            form_filename = str(form.get("filename") or "")
            declared_text = subject + " " + body
            subject_declared = self._drop_generic_epr(
                self._extract_projects_by_rules(declared_text)
            )
            subject_names = list(dict.fromkeys(
                str(item.get("standard_name") or "").strip()
                for item in subject_declared
                if str(item.get("standard_name") or "").strip()
            ))
            # 泛欧/8国申请表经常沿用上一单的勾选状态。只要主题或正文
            # 已明确声明了具体业务，声明集合就是本封邮件的业务范围；
            # 不能因为表格残留勾选而凭空增加其它国家的明细。
            if (
                subject_names
                and len(projects) > len(subject_names)
                and re.search(r"泛欧|8\s*国|pan[- ]?epr|pan[- ]?europe", form_filename, re.I)
                and set(subject_names).issubset({p.get("standard_name") for p in projects})
            ):
                original_projects = [p.get("standard_name") for p in projects]
                projects = [
                    p for p in projects if p.get("standard_name") in set(subject_names)
                ]
                form = dict(form)
                form["raw_checked_projects"] = original_projects
                form["projects"] = list(subject_names)
                form["subject_gate"] = "；".join(subject_names)
                self._log(
                    f"泛欧申请表按主题/正文项目闸门收敛: {original_projects} → {subject_names}"
                )
            # 混合场景: 同封邮件里若还有「旧式(无复选框)申请表」, 用主题做闸门补充。
            # 旧式表(如 荷兰EPR申请表.xlsx)把整张 国家×业务 网格都印在单元格里,
            # 直接文本扫描会炸出十几个错误的组合, 故只保留主题中声明过的组合。
            extras = self._extract_from_plain_attachments(attachments, subject)
            added = []
            for e in extras:
                if e["standard_name"] not in {p["standard_name"] for p in projects}:
                    projects.append(e)
                    added.append(e["standard_name"])
            if added:
                self._log(f"旧式申请表补充(主题闸门过滤后): {added}")

            incomplete = bool(form.get("unmatched_countries") or form.get("orphan_business"))
            need_review = incomplete or not projects
            if projects:
                self._log(
                    f"EPR申请表勾选项目: {[p['standard_name'] for p in projects]}"
                    f" (来源: {form.get('filename', '')})"
                )
            if need_review:
                reason = []
                if form.get("unmatched_countries"):
                    reason.append(f"国家已勾选但未勾业务: {form['unmatched_countries']}")
                if form.get("orphan_business"):
                    reason.append(f"业务已勾选但未勾国家: {form['orphan_business']}")
                if not reason:
                    reason.append("申请表未勾选任何项目")
                self._log(f"EPR申请表勾选不完整 → 转人工: {'; '.join(reason)}", "warning")

            applicant_group = self._epr_form_applicant_group(form, attachments)
            if applicant_group.get("customer") or applicant_group.get("company_candidates"):
                applicant_group["projects"] = [p["standard_name"] for p in projects]
                applicant_group["needs_review"] = bool(
                    applicant_group.get("needs_review") or need_review
                )
                return {
                    "projects": projects,
                    "epr_form": form,
                    "need_review": applicant_group["needs_review"],
                    "groups": [applicant_group],
                }

            # 结构化行提供了更可靠的客户名/公司-项目对应关系
            if groups:
                if len(groups) == 1 and groups[0]["customer"]:
                    return {
                        "projects": projects, "epr_form": form,
                        "need_review": need_review,
                        "groups": [{
                            "customer": groups[0]["customer"],
                            "source": groups[0]["source"],
                            "projects": [p["standard_name"] for p in projects],
                        }],
                    }
                # 多家公司的结构化清单比单一勾选表更具体 → 以结构化行为准
                usable = [g for g in groups if g["projects"]]
                if usable:
                    self._log(f"勾选表与结构化行并存, 以结构化行为准: "
                              f"{[(g['customer'], g['projects']) for g in usable]}")
                    return {
                        "projects": [{"raw_value": p, "standard_name": p}
                                     for g in usable for p in g["projects"]],
                        "epr_form": form, "need_review": need_review,
                        "groups": usable,
                    }
            return {"projects": projects, "epr_form": form, "need_review": need_review,
                    "groups": None}

        # ---------- P1: 文本规则(无勾选式申请表) ----------

        # P1a: 结构化「公司+项目」清单行 (代理群发多公司邮件)
        if groups:
            declared = self._drop_generic_epr(
                self._extract_projects_by_rules(subject + " " + body))
            usable = [g for g in groups if g["projects"]]
            if len(groups) == 1 and declared:
                # 单公司主题中常把国家拆成“比利时+波兰+荷兰+丹麦包装法”，
                # 结构化分段只能在最后一段看到“丹麦包装法”，但全局规则已经
                # 识别出四个国家。单公司时这些声明项目都属于同一家公司，必须
                # 合并回分组，不能因 usable 非空而只保留最后一个项目。
                declared_names = [p["standard_name"] for p in declared]
                existing_names = list(groups[0].get("projects") or [])
                groups[0]["projects"] = list(dict.fromkeys(existing_names + declared_names))
                usable = [groups[0]]
            if usable:
                self._log(f"结构化行提取: {[(g['customer'], g['projects']) for g in usable]}")
                return {
                    "projects": [{"raw_value": p, "standard_name": p}
                                 for g in usable for p in g["projects"]],
                    "epr_form": None,
                    "need_review": any(g.get("needs_review") for g in usable),
                    "groups": usable,
                }

        # P1b: 主题/正文优先闸门
        #   邮件主题/正文是客户或代理自己写的业务声明, 可靠性远高于附件
        #   (旧式申请表把整张国家×业务网格印在单元格里, 附件文本扫描必多提)。
        #   主题+正文能提取到组合时, 附件文本不再参与项目提取。
        declared = self._extract_projects_by_rules(subject + " " + body)
        declared = self._drop_generic_epr(declared)
        if declared:
            self._log(f"主题/正文声明项目(附件文本不参与): "
                      f"{[p['standard_name'] for p in declared]}")
            return {"projects": declared, "epr_form": None, "need_review": False,
                    "groups": None}

        # P1c: 主题+正文无声明 → 全文扫描(含附件), 宁可多提不漏提
        matched_projects = []

        for proj_name in self.project_name_list:
            if not proj_name:
                continue
            if proj_name in search_text:
                matched_projects.append({
                    "raw_value": proj_name,
                    "standard_name": proj_name,
                })

        # 去重
        seen = set()
        unique = []
        for p in matched_projects:
            key = p["standard_name"]
            if key not in seen:
                seen.add(key)
                unique.append(p)

        # 附件四未命中（或未导入附件四）→ 离线规则提取: "国家+业务类型" 相邻模式
        if not unique:
            unique = self._drop_generic_epr(self._extract_projects_by_rules(search_text))
            if unique:
                self._log(f"规则提取项目(无附件四匹配): {[p['standard_name'] for p in unique]}")

        if not unique:
            self._log(f"项目未匹配: {subject[:40]}", "warning")

        return {"projects": unique, "epr_form": None, "need_review": False,
                "groups": None}

    @staticmethod
    def _pick_epr_form(attachments: list) -> Optional[dict]:
        """从附件(含压缩包内)中挑出 EPR 申请表勾选结果; 多份时取勾选最多的那份"""
        forms = []
        for att in attachments or []:
            for f in att.get("epr_forms") or []:
                if f:
                    forms.append(f)
        if not forms:
            return None
        forms.sort(
            key=lambda f: (len(f.get("projects") or []) + len(f.get("unmatched_countries") or [])),
            reverse=True,
        )
        return forms[0]

    @staticmethod
    def _epr_form_applicant_group(form: dict, attachments: list) -> dict:
        """按表单自己的成员路径找主体候选，避免把压缩包中不同公司混在一起。"""
        form_name = str(form.get("filename") or "").replace("\\", "/").strip().casefold()
        matched = []
        for attachment in attachments or []:
            for record in attachment.get("structured_records") or []:
                if not isinstance(record, dict) or record.get("record_type") != "epr_application":
                    continue
                record_name = str(
                    record.get("attachment_name") or attachment.get("filename") or ""
                ).replace("\\", "/").strip().casefold()
                if record_name == form_name:
                    matched.append(record)
        field_priority = {
            "company_name_zh": 0,
            "company_name": 1,
            "applicant_name": 2,
            "company_name_en": 3,
        }
        matched.sort(key=lambda item: field_priority.get(str(item.get("company_field") or ""), 5))
        candidates = []
        seen = set()
        for record in matched:
            value = str(record.get("customer") or "").strip()
            key = normalize_text(value)
            if not value or not key or key in seen:
                continue
            seen.add(key)
            candidates.append({
                "id": f"epr-company-{len(candidates) + 1}",
                "value": value,
                "field_type": str(record.get("company_field") or "company_name"),
                "source": form_name or "EPR申请表",
                "evidence": str(record.get("raw_text") or record.get("field_label") or "")[:300],
            })
        preferred = candidates[0] if candidates else None
        return {
            "customer": preferred.get("value", "") if preferred else "",
            "source": "附件EPR申请表格式化字段" if preferred else "待确认",
            "company_candidates": candidates if len(candidates) > 1 else [],
            "attachment_record_source": f"附件EPR申请表：{form.get('filename', '')}",
            "needs_review": not bool(preferred) or len(candidates) > 1,
        }

    def _extract_from_plain_attachments(self, attachments: list, subject: str) -> List[Dict]:
        """混合场景补充: 从「不含复选框的附件」里提取项目, 并用主题做闸门过滤。

        背景:
          同一封邮件可能既有可勾选式申请表(如 法国电池法申请表.xlsx), 又有旧式无勾选
          申请表(如 荷兰EPR申请表.xlsx / 意大利EPR申请表.xlsx)。旧式表把整张
          「国家 × 业务」网格(荷兰WEEE产品信息 / 荷兰电池法信息 / 荷兰包装法信息)
          都写进了单元格, 纯文本扫描会炸出一堆并不存在的组合。
          做法: 只在这些旧式附件里扫, 且**仅保留主题中同时出现「国家」与「业务」的组合**,
                主题是客户/代理自己写的业务声明, 是这一场景下最可靠的锚点。
        """
        if not subject:
            return []
        plain_text = []
        for att in attachments or []:
            if att.get("epr_forms"):
                continue  # 勾选表的文本(含整张网格)不参与, 否则必然多提
            plain_text.append(att.get("text_content") or "")
            plain_text.append(att.get("filename") or "")
        text = " ".join(plain_text)
        if not text.strip():
            return []

        hits = self._extract_projects_by_rules(text)
        kept = []
        for h in hits:
            if self._subject_declares(h["standard_name"], subject):
                kept.append(h)
        return kept

    @staticmethod
    def _subject_declares(standard_name: str, subject: str) -> bool:
        """主题里是否同时声明了该项目对应的「国家」与「业务」"""
        if not standard_name or not subject:
            return False
        subj = subject.lower()
        country = next((c for c in COUNTRIES if standard_name.startswith(c)), None)
        if not country:
            return False
        rest = standard_name[len(country):]
        if country.lower() not in subj:
            return False
        if not rest:
            return False
        for std_bt, pat in FieldExtractor._BT_PATTERNS:
            if rest == std_bt:
                # 用该业务的全部别名回主题里找一遍
                aliases = [a.replace("\\s*", "").replace("(?:法)?", "").strip("()|")
                           for a in pat.split("|")]
                return any(a and a.lower() in subj for a in aliases)
        return False


    # 业务类型别名表: 标准名 → 正则别名
    _BT_PATTERNS = [
        # 部分代理在主题里把 WEEE 简写成 EEE；只在国家×业务的实体模式中
        # 使用这个别名，随后仍规范为“国家WEEE”。
        ("WEEE", r"weee|eee"),
        ("电池法", r"电池法|便携式电池|电池"),
        ("包装法", r"包装法|包装"),
        ("一次性塑料", r"一次性塑料(?:法)?|塑料法"),
        ("EPR", r"epr"),
    ]

    # 国家二字码别名 (主题常用缩写: "PL包装法" / "IT电池&WEEE" / "DE一次性塑料")
    # 排除易撞英文单词的 no(挪威)/at(奥地利), 遇到它们仍走中文全名
    _COUNTRY_CODES = {
        "de": "德国", "nl": "荷兰", "ie": "爱尔兰", "it": "意大利", "be": "比利时",
        "pl": "波兰", "dk": "丹麦", "fr": "法国", "cz": "捷克", "pt": "葡萄牙",
        "se": "瑞典", "es": "西班牙", "lu": "卢森堡", "hu": "匈牙利", "fi": "芬兰",
        "ro": "罗马尼亚", "ee": "爱沙尼亚", "ch": "瑞士", "gr": "希腊",
        "uk": "英国", "gb": "英国", "ca": "加拿大", "lv": "拉脱维亚",
    }

    # 英文公司后缀 (词边界锚定, 防止 "attachments" 之类误命中)
    _COMPANY_EN_LETTER_CHARS = "A-Za-zÀ-ÖØ-öø-ÿĄĆĘŁŃÓŚŹŻąćęłńóśźż"
    _COMPANY_EN_CHARS = _COMPANY_EN_LETTER_CHARS + r"0-9 .,&'’()\-"
    _COMPANY_EN_SUFFIX = (
        r"(?:CO\.,?\s*LTD\.?|LIMITED|LTD\.?|LLC|INC\.?|GMBH|GBR|"
        r"B\.V\.?|S\.A\.S|SAS|N\.V|S\.?\s*R\.?\s*O\.?|SRO|"
        r"S\.?\s*R\.?\s*L\.?|"
        r"O\.U\.?|U\.A\.B\.?|PTE\.?\s*LTD\.?|PTY\.?\s*LTD\.?|"
        r"CORPORATION|CORP\.?|S\.L\.?|"
        r"SP[ÓO]ŁKA\s+Z\s+OGRANICZON[ĄA]\s+ODPOWIEDZIALNO[ŚS]CI[ĄA])"
    )

    # 公司名候选里出现这些词 → 是表单说明文字, 不是公司名
    _DISCLAIMER_HINTS = (
        "适用", "適用", "说明", "說明", "如下", "董事", "成员", "成員",
        "邮箱", "郵箱", "法人", "上述", "授權", "授权", "確認", "确认",
        "提供", "境外", "注册证书", "註冊證書", "注册一个", "申請表", "申请表",
        "公司名称", "企业名称", "营业执照", "所需资料", "申请资料", "填写", "填报",
        "其他欧盟", "第三国", "设立的公司",
        "非中国公司", "foreign company", "non-chinese company",
        "客户信息", "公司信息", "贵司", "本公司", "申请人",
    )

    def _country_names_in(self, seg: str) -> List[str]:
        """一段文本里出现哪些国家(中文全名或二字码), 返回中文国名(去重保序)"""
        out = []
        for c in COUNTRIES:
            if c in seg and c not in out:
                out.append(c)
        for code, c in self._COUNTRY_CODES.items():
            if re.search(rf"(?<![a-z0-9]){code}(?![a-z0-9])", seg, re.I) and c not in out:
                out.append(c)
        return out

    def _extract_projects_by_rules(self, text: str) -> List[Dict]:
        """
        离线规则提取项目: 不依赖附件四，扫描 "国家+业务类型" 相邻模式
        覆盖: "德国WEEE"、"法国包装法"、"波兰荷兰包装法"(国家连排)、
              "德国和奥地利包装法"(和/与/及连排)、"WEEE德国"(反序)、
              "PL包装法"/"IT电池&WEEE"(二字码)
        """
        if not text:
            return []
        text_l = text.lower()
        # 国家 alternation: 中文全名 + 带词边界的二字码
        country_alt = "|".join(COUNTRIES) + "|" + "|".join(
            rf"(?<![a-z0-9]){code}(?![a-z0-9])" for code in self._COUNTRY_CODES
        )
        # 国家连排: 德国 / 德国荷兰 / 德国和奥地利 (允许 和/与/及/、 分隔及空白)
        country_run = rf"(?:{country_alt})(?:\s*[和与及、,，+＋/／&＆;；]\s*(?:{country_alt}))*"
        found = []
        seen = set()

        def _add(raw: str, std: str):
            if std and std not in seen:
                seen.add(std)
                found.append({"raw_value": raw, "standard_name": std})

        for std_bt, bt_pat in self._BT_PATTERNS:
            # 模式1: 国家+业务 (含国家连排与二字码)
            for m in re.finditer(rf"({country_run})\s*(?:{bt_pat})", text_l):
                for c in self._country_names_in(m.group(1)):
                    _add(m.group(0), c + std_bt)
            # 模式2: 业务+国家 (反序, 如 "WEEE德国")
            for m in re.finditer(rf"(?:{bt_pat})\s*({country_run})", text_l):
                for c in self._country_names_in(m.group(1)):
                    _add(m.group(0), c + std_bt)

        # 模式3: 国家+业务1(+业务2...) 共享国家前缀
        #   (如 "荷兰WEEE+包装法"、"IT电池&WEEE"、"德国WEEE/电池法/包装法")
        all_bt = "|".join(pat for _, pat in self._BT_PATTERNS)
        alias_to_std = {}
        for std_bt, pat in self._BT_PATTERNS:
            for alias in pat.split("|"):
                alias_to_std[alias] = std_bt
        sep = r"\s*[+＋/／、，,&＆;；和与及]\s*"
        for m in re.finditer(rf"({country_alt})\s*((?:{all_bt})(?:{sep}(?:{all_bt}))+)", text_l):
            c = self._country_names_in(m.group(1))[0]
            for bt_m in re.finditer(all_bt, m.group(2)):
                _add(m.group(0), c + alias_to_std[bt_m.group(0)])

        return found

    @staticmethod
    def _drop_generic_epr(projects: List[Dict]) -> List[Dict]:
        """同国家已有具体业务(WEEE/电池法/包装法/一次性塑料)时, 去掉泛称的"国家EPR" """
        specific = set()
        for p in projects:
            n = p.get("standard_name", "")
            for c in COUNTRIES:
                if n.startswith(c) and n != c + "EPR":
                    specific.add(c)
        out = []
        for p in projects:
            n = p.get("standard_name", "")
            if n.endswith("EPR") and any(n.startswith(c) for c in specific):
                continue
            out.append(p)
        return out

    def _calc_confidence(self, agent_r: dict, customer_r: dict, proj_r: dict) -> str:
        """计算置信度"""
        score = 0
        if agent_r.get("agent"):
            score += 1
        if customer_r.get("customer"):
            score += 1
        if proj_r.get("standard_name"):
            score += 1

        if agent_r.get("match_method", "").startswith("多匹配"):
            return "low"
        if score == 3:
            return "high"
        if score >= 2:
            return "medium"
        return "low"
