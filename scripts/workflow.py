"""Rule evaluation, workbook/PDF rendering, and versioned packaging."""

from __future__ import annotations

import csv
import io
import json
import re
import shutil
from collections import Counter, defaultdict
from datetime import date
from decimal import Decimal
from html import escape
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from engine import CATEGORY_ORDER, _seller, money, money_str, natural_key, read_source, write_json

try:
    import fitz
except ImportError:  # pragma: no cover
    fitz = None

from openpyxl import Workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.pagebreak import Break
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


EXPECTED_BUYER = "西安交通大学"
EXPECTED_TAX_ID = "12100000435230200R"
INVENTORY_THRESHOLD = Decimal("500.00")
HIGH_VALUE_THRESHOLD = Decimal("1000.00")
MATERIAL_CATEGORIES = {"材料费", "日用及两用物品"}
FORMAL_STATUSES = {"ready", "missing", "pending"}
CHINESE_FONT = "黑体"
LATIN_FONT = "Times New Roman"
BLACK = "FF000000"
# V6 原版明细表的 Accent 6（#70AD47）加 80% 白色，与 E2EFDA 等效。
HIGHLIGHT = "FFE2EFDA"
STATUS_LABELS = {
    "ready": "可提交", "missing": "缺件", "pending": "待核实",
    "ineligible": "不可报销", "replaced": "已替换",
}
ATTACHMENT_LABELS = {
    "waybill": "运单或运费清单", "itinerary": "网约车行程单", "metro_trip": "地铁行程截图",
    "print_list": "打印清单", "print_proof": "打印内容证明", "payment": "支付记录",
    "fee_notice": "缴费通知", "acceptance": "审稿/录用通知", "paper_first_page": "论文首页",
    "test_report": "测试清单或报告", "sales_list": "盖章销货清单",
    "inventory": "入库单", "high_value": "高值耗材登记材料", "meeting_notice": "会议通知",
}


def _safe_text(value: Any) -> str:
    text = str(value or "")
    return "'" + text if text[:1] in {"=", "+", "-", "@"} else text


def _safe_name(value: str, limit: int = 80) -> str:
    value = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", value).strip(" .")
    return value[:limit] or "未命名"


def _compact_money(value: Decimal) -> str:
    return format(value, ".2f").rstrip("0").rstrip(".")


def _attachment_name(att: dict[str, Any]) -> str:
    return Path(att["ref"].split("::")[-1]).name


def _workbook_attachment_label(att: dict[str, Any]) -> str:
    name = _attachment_name(att)
    evidence = name + " " + str(att.get("text_excerpt") or "")
    kind = att.get("attachment_type")
    if kind == "payment":
        ext = str(att.get("ext") or Path(name).suffix).lower()
        return "支付记录截图" if ext in {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".webp"} else "支付记录"
    if kind == "itinerary":
        return "行程单"
    if kind == "waybill":
        return "订单详情" if "订单详情" in name else "运单明细" if any(x in name for x in ("明细", "清单")) else "运单"
    if kind == "acceptance":
        return "论文录用通知" if "录用" in evidence else "审稿通知" if "审稿" in evidence else "审稿/录用通知"
    labels = {
        "metro_trip": "地铁行程截图", "print_list": "打印清单", "print_proof": "打印内容证明",
        "fee_notice": "缴费通知", "paper_first_page": "论文首页", "test_report": "测试报告",
        "sales_list": "销货清单", "inventory": "入库单", "high_value": "高值材料审批材料",
        "meeting_notice": "会议通知",
    }
    return labels.get(kind, Path(name).stem[:24] or "其他附件")


def _workbook_attachment_summary(attached: list[dict[str, Any]]) -> str:
    unique = {att["ref"]: att for att in attached}
    labels = [_workbook_attachment_label(att) for att in unique.values()]
    counts = Counter(labels)
    return "附件：" + "、".join(f"{label}×{counts[label]}" if counts[label] > 1 else label for label in dict.fromkeys(labels)) if labels else ""


def _category(record: dict[str, Any]) -> str:
    manual = record.get("manual", {})
    return str(manual.get("category") or record.get("category_guess") or "其他")


def _description(record: dict[str, Any]) -> str:
    manual = record.get("manual", {})
    subject = str(manual.get("service_subject") or "").strip()
    if _category(record) == "网络服务费" and subject:
        return subject if subject.endswith("网络服务费") else subject + "网络服务费"
    if _category(record) == "网络服务费":
        return "平台待核实网络服务费"
    return str(manual.get("description") or record.get("description_guess") or "发票")


def _record_amount(record: dict[str, Any]) -> Decimal | None:
    manual = record.get("manual", {})
    return money(manual.get("claim_amount") or record.get("amount"))


def _record_number(record: dict[str, Any]) -> str | None:
    return str(record.get("manual", {}).get("invoice_number") or record.get("invoice_number") or "") or None


def _record_date(record: dict[str, Any]) -> str | None:
    return str(record.get("manual", {}).get("issue_date") or record.get("issue_date") or "") or None


def _is_public(record: dict[str, Any], draft: dict[str, Any]) -> bool:
    return record["id"] in set(draft.get("public_record_ids", [])) or _record_number(record) in set(draft.get("public_invoice_numbers", []))


def _attachment_targets(records: list[dict[str, Any]], attachments: list[dict[str, Any]]) -> tuple[dict[str, list[str]], dict[str, list[str]]]:
    by_id: dict[str, list[str]] = {r["id"]: [] for r in records}
    by_ref: dict[str, list[str]] = {}
    for record in records:
        for ref in record.get("manual", {}).get("attachment_refs", []):
            if ref in {a["ref"] for a in attachments}:
                by_id[record["id"]].append(ref)
                by_ref.setdefault(ref, []).append(record["id"])
    for att in attachments:
        if att["ref"] in by_ref:
            continue
        name = _attachment_name(att)
        text = att.get("text_excerpt", "")
        candidates = [r for r in records if _record_number(r) and _record_number(r) in name + text]
        if len(candidates) != 1:
            amount_tokens = re.findall(r"\d+(?:\.\d{1,2})?\s*元", name)
            amounts = {money(x.replace("元", "")) for x in amount_tokens}
            if amounts:
                candidates = [r for r in records if _record_amount(r) in amounts]
        if len(candidates) != 1:
            prefix = re.match(r"^(\d+)[-_－]", name)
            if prefix:
                candidates = [r for r in records if re.match(r"^" + re.escape(prefix.group(1)) + r"[-_－]", Path(r["ref"].split("::")[-1]).name)]
        if len(candidates) != 1:
            kind = att.get("attachment_type")
            possible = [r for r in records if (
                kind == "waybill" and _category(r) == "快递费"
                or kind == "metro_trip" and "地铁" in _description(r)
                or kind == "itinerary" and _category(r) == "交通费" and "地铁" not in _description(r)
                or kind in {"inventory", "high_value"} and _category(r) == "材料费"
                or kind in {"fee_notice", "acceptance", "paper_first_page"} and _category(r) == "版面费/审稿费"
            )]
            candidates = possible if len(possible) == 1 else []
        if len(candidates) == 1:
            record = candidates[0]
            by_id[record["id"]].append(att["ref"])
            by_ref.setdefault(att["ref"], []).append(record["id"])
    for refs in by_id.values():
        refs.sort(key=natural_key)
    return by_id, by_ref


def _item_thresholds(record: dict[str, Any]) -> tuple[list[str], list[str]]:
    inventory: list[str] = []
    high_value: list[str] = []
    for item in record.get("items", []):
        line_gross = money(item.get("line_gross"))
        unit_gross = money(item.get("unit_gross"))
        if line_gross is not None and line_gross >= INVENTORY_THRESHOLD:
            inventory.append(f"{item['name']}（含税行金额 {_compact_money(line_gross)} 元）")
        if unit_gross is not None and unit_gross >= HIGH_VALUE_THRESHOLD:
            high_value.append(f"{item['name']}（含税单价 {_compact_money(unit_gross)} 元）")
    return inventory, high_value


def _high_value_reasons(record: dict[str, Any]) -> list[str]:
    return _item_thresholds(record)[1]


def _missing_docs(record: dict[str, Any], attached: list[dict[str, Any]], public: bool) -> tuple[list[str], list[str]]:
    kinds = {a.get("attachment_type") for a in attached}
    missing: list[str] = []
    manual_tasks: list[str] = []
    category = _category(record)
    description = _description(record)
    amount = _record_amount(record)
    invoice_total = money(record.get("amount"))
    if invoice_total is None:
        invoice_total = amount
    if category == "快递费" and "waybill" not in kinds:
        missing.append(ATTACHMENT_LABELS["waybill"])
    if category == "交通费":
        if any(x in description for x in ("网约车", "滴滴", "高德", "打车")) and "itinerary" not in kinds:
            missing.append(ATTACHMENT_LABELS["itinerary"])
        if "地铁" in description and "metro_trip" not in kinds:
            missing.append(ATTACHMENT_LABELS["metro_trip"])
        if "出租车" in description:
            manual_tasks.append("出租车纸票按西安市票据要求粘贴，并用铅笔写明起点—终点")
    if category == "文印费":
        if "print_list" not in kinds:
            missing.append(ATTACHMENT_LABELS["print_list"])
        if amount is not None and amount > Decimal("200") and "print_proof" not in kinds:
            missing.append(ATTACHMENT_LABELS["print_proof"])
    if category == "版面费/审稿费":
        for kind in ("payment", "fee_notice", "acceptance", "paper_first_page"):
            if kind not in kinds:
                missing.append(ATTACHMENT_LABELS[kind])
    if category == "测试化验加工" and "test_report" not in kinds:
        missing.append(ATTACHMENT_LABELS["test_report"])
    if category == "查收查引/专利" and "查" in description:
        if amount is not None and amount > Decimal("200") and "sales_list" not in kinds and "test_report" not in kinds:
            missing.append("查收查引清单")
    if category == "日用及两用物品" and not record.get("manual", {}).get("research_use_confirmed"):
        missing.append("科研用途待负责人确认")
    if category in MATERIAL_CATEGORIES:
        inventory, _ = _item_thresholds(record)
        high_value = _high_value_reasons(record)
        if inventory and "inventory" not in kinds:
            missing.append("入库单：" + "；".join(inventory))
        if high_value and "high_value" not in kinds:
            missing.append("高值耗材登记材料：" + "；".join(high_value))
    if category == "书籍" and not any("书" in item.get("name", "") for item in record.get("items", [])):
        manual_tasks.append("若发票未列书名，在纸质 A4 空白处用铅笔写明书名")
    if "定额发票" in record.get("text_excerpt", "") and "sales_list" not in kinds:
        missing.append(ATTACHMENT_LABELS["sales_list"])
    if category == "差旅费":
        manual_tasks.append("同一行程的交通、住宿、会议材料放在一起；请负责人确认具体附件")
        if any(x in description for x in ("会议费", "注册费")) and "meeting_notice" not in kinds:
            missing.append(ATTACHMENT_LABELS["meeting_notice"])
        if any(x in description for x in ("火车", "高铁", "飞机", "机票", "注册费")):
            manual_tasks.append("机票、高铁或注册费发票需提交双份")
    if public and "payment" not in kinds:
        missing.append("对公转账付款凭证")
    elif invoice_total is not None and invoice_total > HIGH_VALUE_THRESHOLD and "payment" not in kinds:
        missing.append("支付记录（票面总价超 1000 元）")
    if record["ext"] in {".jpg", ".jpeg", ".png", ".tif", ".tiff"} and not record.get("manual", {}).get("paper_original_confirmed"):
        missing.append("纸质原件或电子发票原 PDF（图片仅作对照）")
    return list(dict.fromkeys(missing)), list(dict.fromkeys(manual_tasks))


def evaluate(draft: dict[str, Any]) -> tuple[list[dict[str, Any]], dict[str, list[str]]]:
    profile = draft.get("profile", {})
    for key in ("name", "student_id", "phone", "cutoff_date"):
        if not str(profile.get(key, "")).strip():
            raise ValueError(f"个人信息缺少 {key}；先在 review.json 中填写")
    try:
        cutoff = date.fromisoformat(profile["cutoff_date"])
    except ValueError as exc:
        raise ValueError("截止日期须为 YYYY-MM-DD") from exc
    records = draft.get("records", [])
    ignored_refs = set(draft.get("ignored_attachment_refs", []))
    attachments = [a for a in draft.get("attachments", []) if a["ref"] not in ignored_refs]
    attached_by_id, targets_by_ref = _attachment_targets(records, attachments)
    att_by_ref = {a["ref"]: a for a in attachments}
    replaced_ids = {r.get("manual", {}).get("replacement_for") for r in records if r.get("manual", {}).get("replacement_for")}
    numbers: dict[str, list[str]] = defaultdict(list)
    for record in records:
        number = _record_number(record)
        if number:
            numbers[number].append(record["id"])
    result: list[dict[str, Any]] = []
    for record in records:
        item = dict(record)
        item["category"] = _category(record)
        item["description"] = _description(record)
        item["claim_amount"] = money_str(_record_amount(record))
        item["invoice_number_final"] = _record_number(record)
        item["issue_date_final"] = _record_date(record)
        extracted_seller_name, extracted_seller_tax_id = _seller(str(record.get("text_excerpt") or ""))
        manual = record.get("manual", {})
        item["seller_name"] = manual.get("seller_name") or record.get("seller_name") or extracted_seller_name
        item["seller_tax_id"] = manual.get("seller_tax_id") or record.get("seller_tax_id") or extracted_seller_tax_id
        item["public_transfer"] = _is_public(record, draft)
        item["attached_refs"] = attached_by_id.get(record["id"], [])
        linked = [att_by_ref[ref] for ref in item["attached_refs"] if ref in att_by_ref]
        item["attachment_summary"] = _workbook_attachment_summary(linked)
        item["missing_docs"], item["manual_tasks"] = _missing_docs(record, linked, item["public_transfer"])
        pending: list[str] = []
        ineligible: list[str] = []
        warnings: list[str] = []
        manual_note = str(manual.get("notes") or "")
        if manual_note.startswith("待核实："):
            pending.append(manual_note)
        amount = money(item["claim_amount"])
        printed = money(record.get("amount"))
        if record["id"] in replaced_ids:
            item["status"] = "replaced"
            item["issues"] = ["已由更正发票替换，旧版留档"]
            result.append(item)
            continue
        if record.get("scan_errors"):
            pending.extend(record["scan_errors"])
        for attachment in linked:
            if attachment.get("scan_errors"):
                pending.append("关联附件无法可靠读取：" + attachment["ref"])
        if record.get("multiple_numbers"):
            pending.append("同一文件片段含多个发票号码")
        if not item["invoice_number_final"]:
            pending.append("未确认发票号码")
        if not item["issue_date_final"]:
            pending.append("未确认开票日期")
        if printed is None and money(manual.get("claim_amount")) is None:
            pending.append("未确认票面金额")
        if amount is None:
            pending.append("未确认报销金额")
        filename_amount = money(record.get("filename_amount"))
        if printed is not None and filename_amount is not None and printed != filename_amount and not manual.get("confirm_filename_amount"):
            pending.append(f"旧文件名金额 {_compact_money(filename_amount)} 元与票面 {_compact_money(printed)} 元不符")
        if printed is not None and amount is not None and printed != amount and not manual.get("confirm_claim_amount"):
            pending.append("报销金额与票面不符，需支付凭证并人工确认")
        if printed is not None and amount is not None and printed != amount and not any(a.get("attachment_type") == "payment" for a in linked):
            pending.append("报销金额与票面不同但未发现支付凭证")
        for attachment in linked:
            if attachment.get("attachment_type") == "payment":
                paid = money(attachment.get("payment_amount"))
                if paid is not None and amount is not None and paid != amount and not manual.get("confirm_payment_amount"):
                    pending.append(f"付款凭证金额 {_compact_money(paid)} 元与报销金额 {_compact_money(amount)} 元不符")
        if item["invoice_kind"] == "vat":
            buyer = manual.get("buyer_name") or record.get("buyer_name")
            tax_id = manual.get("buyer_tax_id") or record.get("buyer_tax_id")
            if manual.get("confirmed_buyer_absent") or manual.get("confirmed_wrong_buyer"):
                ineligible.append("普通增值税发票购买方抬头或税号经核对缺失/不符")
            elif (manual.get("buyer_name") or record.get("buyer_explicit")) and buyer and buyer != EXPECTED_BUYER:
                ineligible.append("普通增值税发票购买方抬头不符：" + str(buyer))
            elif (manual.get("buyer_tax_id") or record.get("buyer_explicit")) and tax_id and tax_id != EXPECTED_TAX_ID:
                ineligible.append("普通增值税发票购买方税号不符：" + str(tax_id))
            elif buyer == EXPECTED_BUYER and tax_id == EXPECTED_TAX_ID:
                pass
            elif manual.get("buyer_verified") and buyer == EXPECTED_BUYER and tax_id == EXPECTED_TAX_ID:
                pass
            else:
                pending.append("购买方抬头或税号未可靠确认为西安交通大学")
        elif item["invoice_kind"] == "unknown":
            pending.append("票据类型未确认，无法应用抬头例外")
        if item["category"] in MATERIAL_CATEGORIES and not record.get("items_verified"):
            pending.append("商品行金额无法与票面总额核对，不能判断入库/高值门槛")
        if item["category"] == "网络服务费" and not str(manual.get("service_subject") or "").strip():
            pending.append("网络服务平台或产品主体未确认；核对订单、开票方及其官方产品信息")
        if record.get("mixed_ambiguous") and not manual.get("confirm_mixed_category"):
            pending.append("同票含多种费用且主类别不明确，请确认类别")
        elif record.get("mixed_categories"):
            warnings.append("同票含" + "、".join(record["mixed_categories"]) + "，按主类别一票一行登记")
        if item["category"] in MATERIAL_CATEGORIES and record.get("items_verified"):
            for line in record.get("items", []):
                gross = money(line.get("line_gross"))
                if gross is not None and gross >= HIGH_VALUE_THRESHOLD and not line.get("quantity"):
                    pending.append("商品行金额满 1000 元但缺数量，不能判断含税单价和高值手续")
                    break
        if manual.get("confirmed_nonreimbursable"):
            ineligible.append("商品经人工核对属于 V6 限制或禁止项目")
        if record.get("items_verified"):
            for line in record.get("items", []):
                # The tax category can be "计算机配套产品" even for an HDMI adapter or
                # toner.  Only the actual product name is tested here.
                product = line.get("name", "").split("*", 1)[-1]
                if any(term in product for term in ("数据线", "手机充电线")):
                    ineligible.append("商品明细含 V6 明示不可报销品名：" + product)
                    break
                if "计算机" in product or any(term in product for term in ("化学试剂助剂", "化学合成材料")):
                    pending.append("商品名称涉及 V6 限制项目，请负责人确认：" + product)
                    break
        issue_date = item["issue_date_final"]
        if issue_date:
            try:
                issued = date.fromisoformat(issue_date)
                if issued > cutoff:
                    pending.append("开票日期晚于本次截止日期")
                elif issued.year < cutoff.year:
                    warnings.append("跨年度发票：请负责人确认")
            except ValueError:
                pending.append("开票日期格式错误")
        number = item["invoice_number_final"]
        if number and number in set(draft.get("history_invoice_numbers", [])):
            pending.append("历史目录中发现相同发票号码，疑似重复")
        active_same_number = [rid for rid in numbers.get(number, []) if rid not in replaced_ids] if number else []
        if number and len(active_same_number) > 1:
            pending.append("本次材料中出现相同发票号码")
        if item["category"] == "文印费" and "论文" in record.get("text_excerpt", ""):
            pending.append("打印费发票内容含“论文”，请负责人核查")
        if item["category"] == "交通费" and any(x in record.get("text_excerpt", "") for x in ("西安北站", "西安站", "咸阳机场")):
            warnings.append("车站/机场交通应与差旅费材料一起提交")
        item["warnings"] = list(dict.fromkeys(warnings))
        item["pending_reasons"] = list(dict.fromkeys(pending))
        if ineligible:
            item["status"] = "ineligible"
        elif pending:
            item["status"] = "pending"
        elif item["missing_docs"]:
            item["status"] = "missing"
        else:
            item["status"] = "ready"
        item["issues"] = list(dict.fromkeys(ineligible + pending + item["missing_docs"]))
        result.append(item)
    same_seller_day: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for item in result:
        if item["status"] not in FORMAL_STATUSES or not item.get("issue_date_final"):
            continue
        seller_id = str(item.get("seller_tax_id") or "").strip().upper()
        seller_name = re.sub(r"\s+", "", str(item.get("seller_name") or ""))
        if seller_id or seller_name:
            same_seller_day[(item["issue_date_final"], seller_id or "名称:" + seller_name)].append(item)
    grouped = sorted((key, members) for key, members in same_seller_day.items() if len(members) > 1)
    for group_number, (key, members) in enumerate(grouped, 1):
        seller = next((str(m.get("seller_name")) for m in members if m.get("seller_name")), key[1])
        note = f"同开票方同日组 G{group_number}（{len(members)}张，{seller}，{key[0]}）"
        for item in members:
            item["same_seller_date_group"] = f"G{group_number}"
            item["same_seller_date_note"] = note
    return result, targets_by_ref


def _sort_key(record: dict[str, Any]) -> tuple[Any, ...]:
    category = record["category"]
    order = CATEGORY_ORDER.index(category) if category in CATEGORY_ORDER else CATEGORY_ORDER.index("其他")
    if record["public_transfer"]:
        order = CATEGORY_ORDER.index("对公转账")
    threshold_flag = 0
    if category in MATERIAL_CATEGORIES:
        inventory, _ = _item_thresholds(record)
        high_value = _high_value_reasons(record)
        threshold_flag = 1 if inventory or high_value else 0
    return order, category, _display_name_base(record), threshold_flag, record.get("issue_date_final") or "9999", natural_key(record["ref"])


def _display_name_base(record: dict[str, Any]) -> str:
    name = re.sub(r"\s+(?:个|箱|台|套|件|张|次|份|包)$", "", record["description"].strip())
    if "收纳箱" in name or "整理箱" in name:
        return "收纳箱"
    if re.fullmatch(r"滴滴(?:网约车|出行)?[A-Za-z0-9]*", name):
        return "滴滴网约车"
    return name


def _number_repeated_names(records: list[dict[str, Any]]) -> None:
    formal = [r for r in records if r["status"] in FORMAL_STATUSES]
    bases = {r["id"]: _display_name_base(r) for r in formal}
    totals = Counter((r["category"], bases[r["id"]]) for r in formal)
    seen: Counter[tuple[str, str]] = Counter()
    for record in formal:
        base = bases[record["id"]]
        key = (record["category"], base)
        seen[key] += 1
        record["description"] = f"{base}{seen[key]}" if totals[key] > 1 else base


def _is_cjk_character(char: str) -> bool:
    code = ord(char)
    return 0x2E80 <= code <= 0x9FFF or 0xF900 <= code <= 0xFAFF or 0xFF00 <= code <= 0xFFEF


def _style_workbook_fonts(ws: Any) -> None:
    for row in ws:
        for cell in row:
            if isinstance(cell, MergedCell) or cell.value is None:
                continue
            size = cell.font.sz or 11
            bold = bool(cell.font.bold)
            value = cell.value
            if cell.data_type == "f" or not isinstance(value, str):
                cell.font = Font(name=LATIN_FONT, size=size, bold=bold, color=BLACK)
                continue
            runs: list[tuple[str, str]] = []
            leading_space = ""
            for char in value:
                if char.isspace():
                    if runs:
                        runs[-1] = (runs[-1][0], runs[-1][1] + char)
                    else:
                        leading_space += char
                    continue
                font_name = CHINESE_FONT if _is_cjk_character(char) else LATIN_FONT
                char = leading_space + char
                leading_space = ""
                if runs and runs[-1][0] == font_name:
                    runs[-1] = (font_name, runs[-1][1] + char)
                else:
                    runs.append((font_name, char))
            if leading_space and runs:
                runs[-1] = (runs[-1][0], runs[-1][1] + leading_space)
            cell.font = Font(name=runs[0][0] if runs else CHINESE_FONT, size=size, bold=bold, color=BLACK)
            if len(runs) > 1:
                cell.value = CellRichText(*[
                    TextBlock(InlineFont(rFont=font_name, sz=size, b=bold, color=BLACK), text)
                    for font_name, text in runs
                ])


def _write_workbook(path: Path, profile: dict[str, str], records: list[dict[str, Any]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "报销明细表（个人）"
    ws.sheet_view.showGridLines = False
    widths = {"A": 19, "B": 30, "C": 16, "D": 58, "E": 12, "F": 21}
    for column, width in widths.items():
        ws.column_dimensions[column].width = width
    ws.append(["姓名", _safe_text(profile["name"]), "学号", _safe_text(profile["student_id"]), "电话", _safe_text(profile["phone"])])
    ws.row_dimensions[1].height = 26
    border = Border(bottom=Side(style="hair", color="BFC9D2"))
    for cell in ws[1]:
        cell.font = Font(name=CHINESE_FONT, size=11, color=BLACK)
        cell.alignment = Alignment(vertical="center")
    groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        if record["status"] in FORMAL_STATUSES:
            key = "对公转账" if record["public_transfer"] else record["category"]
            groups[key].append(record)
    categories = [x for x in CATEGORY_ORDER if x != "对公转账" and groups.get(x)]
    categories += sorted([x for x in groups if x not in CATEGORY_ORDER and x != "对公转账"])
    if groups.get("对公转账"):
        categories.append("对公转账")
    subtotal_rows: dict[str, int] = {}
    person_subtotals: list[int] = []
    public_subtotal: int | None = None
    for category in categories:
        header_row = ws.max_row + 1
        ws.cell(header_row, 1, category)
        ws.cell(header_row, 2, "小计")
        ws.merge_cells(start_row=header_row, start_column=4, end_row=header_row, end_column=6)
        labels = ("编号", "简要名称", "金额", "备注")
        for col, label in enumerate(labels, 1):
            ws.cell(header_row + 1, col, label)
        ws.merge_cells(start_row=header_row + 1, start_column=4, end_row=header_row + 1, end_column=6)
        start = header_row + 2
        for record in groups[category]:
            row = ws.max_row + 1
            ws.cell(row, 1, record["number"])
            ws.cell(row, 2, _safe_text(record["description"]))
            if record["claim_amount"] is not None:
                ws.cell(row, 3, float(Decimal(record["claim_amount"])))
            notes = []
            if record.get("same_seller_date_note"):
                notes.append(record["same_seller_date_note"])
            if record.get("attachment_summary"):
                notes.append(record["attachment_summary"])
            if record["status"] == "pending":
                notes.extend(
                    reason if reason.startswith("待核实：") else "待核实：" + reason
                    for reason in record.get("pending_reasons", [])
                )
            if record["status"] in {"pending", "missing"}:
                missing = list(record["missing_docs"])
                if "科研用途待负责人确认" in missing:
                    notes.append("待确认：科研用途")
                    missing.remove("科研用途待负责人确认")
                if missing:
                    notes.append("缺件：" + "；".join(missing))
            inventory = _item_thresholds(record)[0] if record["category"] in MATERIAL_CATEGORIES else []
            high_value = _high_value_reasons(record) if record["category"] in MATERIAL_CATEGORIES else []
            if inventory and not any(x.startswith("入库单：") for x in record["missing_docs"]):
                notes.append("需入库：" + "；".join(inventory))
            if high_value and not any(x.startswith("高值耗材登记材料：") for x in record["missing_docs"]):
                notes.append("高值耗材登记：" + "；".join(high_value))
            if record.get("warnings"):
                notes.extend(record["warnings"])
            if record.get("manual", {}).get("notes") and not (
                record["status"] == "pending" and record["manual"]["notes"] in record.get("pending_reasons", [])
            ):
                notes.append(str(record["manual"]["notes"]))
            ws.cell(row, 4, _safe_text("；".join(notes)))
            ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=6)
            ws.cell(row, 3).number_format = '#,##0.00'
            ws.cell(row, 4).alignment = Alignment(wrap_text=True, vertical="center")
            ws.row_dimensions[row].height = min(120, max(24, 24 + len("；".join(notes)) // 75 * 14))
            for cell in ws[row][:4]:
                cell.border = border
        end = ws.max_row
        ws.cell(header_row, 3, f"=SUM(C{start}:C{end})" if end >= start else "=0")
        ws.cell(header_row, 3).number_format = '#,##0.00'
        subtotal_rows[category] = header_row
        if category == "对公转账":
            public_subtotal = header_row
        else:
            person_subtotals.append(header_row)
        for cell in ws[header_row][:6]:
            cell.fill = PatternFill("solid", fgColor=HIGHLIGHT)
            cell.font = Font(name=CHINESE_FONT, size=11, bold=True, color=BLACK)
        for cell in ws[header_row + 1][:4]:
            cell.font = Font(name=CHINESE_FONT, size=10, bold=True, color=BLACK)
            cell.alignment = Alignment(vertical="center")
    total_row = ws.max_row + 1
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    ws.cell(total_row, 1, "个人报销暂计合计")
    ws.cell(total_row, 3, "=" + "+".join(f"C{r}" for r in person_subtotals) if person_subtotals else "=0")
    ws.cell(total_row, 3).number_format = '#,##0.00'
    ws.cell(total_row, 4, "含缺件和待核实票的已知金额；不可报销票未计入")
    ws.cell(total_row, 1).font = Font(name=CHINESE_FONT, size=11, bold=True, color=BLACK)
    ws.cell(total_row, 3).font = Font(name=LATIN_FONT, size=11, bold=True, color=BLACK)
    if public_subtotal:
        row = ws.max_row + 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws.cell(row, 1, "对公转账暂计合计")
        ws.cell(row, 3, f"=C{public_subtotal}")
        ws.cell(row, 3).number_format = '#,##0.00'
        ws.cell(row, 1).font = Font(name=CHINESE_FONT, size=11, bold=True, color=BLACK)
        ws.cell(row, 3).font = Font(name=LATIN_FONT, size=11, bold=True, color=BLACK)
    section_row = ws.max_row + 2
    ws.row_breaks.append(Break(id=section_row - 1))
    ws.cell(section_row, 1, "报销单制作")
    ws.cell(section_row, 2, "（不要删这一部分，打印出来）")
    for cell in ws[section_row][:6]:
        cell.fill = PatternFill("solid", fgColor=HIGHLIGHT)
        cell.font = Font(name=CHINESE_FONT, size=11, bold=cell.column == 1, color=BLACK)
    for category in categories:
        row = ws.max_row + 1
        ws.cell(row, 1, category)
        ws.cell(row, 1).fill = PatternFill("solid", fgColor=HIGHLIGHT)
        ws.cell(row, 1).font = Font(name=CHINESE_FONT, size=11, bold=True, color=BLACK)
        ws.cell(row, 3, f"=C{subtotal_rows[category]}")
        ws.cell(row, 3).number_format = '#,##0.00'
    row = ws.max_row + 1
    ws.cell(row, 1, "总计（含对公）")
    ws.cell(row, 3, "=" + "+".join(f"C{subtotal_rows[c]}" for c in categories) if categories else "=0")
    ws.cell(row, 3).number_format = '#,##0.00'
    ws.cell(row, 1).fill = PatternFill("solid", fgColor=HIGHLIGHT)
    ws.cell(row, 1).font = Font(name=CHINESE_FONT, size=11, bold=True, color=BLACK)
    ws.cell(row, 3).font = Font(name=LATIN_FONT, size=11, bold=True, color=BLACK)
    ws.freeze_panes = "A2"
    ws.print_options.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_area = f"A1:F{ws.max_row}"
    ws.print_title_rows = "1:1"
    ws.oddFooter.center.text = "第 &P 页 / 共 &N 页"
    _style_workbook_fonts(ws)
    wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _pdf_styles() -> tuple[ParagraphStyle, ParagraphStyle, ParagraphStyle]:
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    title = ParagraphStyle("TitleCN", fontName="STSong-Light", fontSize=16, leading=22, alignment=TA_CENTER, spaceAfter=8)
    heading = ParagraphStyle("HeadingCN", fontName="STSong-Light", fontSize=11, leading=15, spaceBefore=6, spaceAfter=3, keepWithNext=True)
    body = ParagraphStyle("BodyCN", fontName="STSong-Light", fontSize=9, leading=13, alignment=TA_LEFT, wordWrap="CJK")
    return title, heading, body


def _write_pdf(path: Path, profile: dict[str, str], records: list[dict[str, Any]], attachments: list[dict[str, Any]], targets_by_ref: dict[str, list[str]], version: int) -> None:
    title, heading, body = _pdf_styles()
    person_total = sum((Decimal(r["claim_amount"]) for r in records if r["status"] in FORMAL_STATUSES and r["claim_amount"] is not None and not r["public_transfer"]), Decimal("0"))
    public_total = sum((Decimal(r["claim_amount"]) for r in records if r["status"] in FORMAL_STATUSES and r["claim_amount"] is not None and r["public_transfer"]), Decimal("0"))
    unmatched = [a for a in attachments if a["ref"] not in targets_by_ref]
    blocking = any(r["status"] in {"pending", "missing"} for r in records) or bool(unmatched)
    story: list[Any] = [Paragraph("个人集中报销补件与复核提醒", title)]
    summary = (
        f"姓名：{escape(profile['name'])}　截止日期：{escape(profile['cutoff_date'])}　版本：v{version}<br/>"
        f"状态：{'尚不可提交' if blocking else '可提交已确认部分'}　个人暂计：{_compact_money(person_total)} 元　"
        f"对公暂计：{_compact_money(public_total)} 元<br/>"
        "本包仅作离线材料核对，未进行税务平台验真；暂计金额含缺件和待核实票的已知金额，不等于最终获批金额；待核实原因见 Excel 备注。"
    )
    story.append(Paragraph(summary, body))

    def section(label: str, subset: list[dict[str, Any]], field: str) -> None:
        story.append(Paragraph(label + f"（{len(subset)}）", heading))
        if not subset:
            story.append(Paragraph("无。", body))
            return
        for r in subset:
            code = f"#{r['number']} " if r.get("number") else ""
            filename = Path(r["ref"].split("::")[-1]).name
            details = r.get(field) or []
            if isinstance(details, str):
                details = [details]
            content = f"{code}{escape(r['description'])}｜{escape(filename)}<br/>" + "；".join(escape(str(x)) for x in details)
            story.append(Paragraph(content, body))
            story.append(Spacer(1, 2))

    section("必须补充的材料", [r for r in records if r["status"] in FORMAL_STATUSES and r["missing_docs"]], "missing_docs")
    section("不可报销的发票", [r for r in records if r["status"] == "ineligible"], "issues")
    offline: dict[str, list[str]] = defaultdict(list)
    for r in records:
        label = f"#{r['number']}" if r.get("number") else r["description"]
        for note in r.get("manual_tasks", []) + r.get("warnings", []):
            offline[note].append(label)
    story.append(Paragraph(f"线下操作与其他提醒（{sum(len(v) for v in offline.values())} 项）", heading))
    if offline:
        for note, labels in offline.items():
            story.append(Paragraph(escape(note) + "：" + "、".join(escape(label) for label in labels), body))
            story.append(Spacer(1, 3))
    else:
        story.append(Paragraph("无。", body))
    story.append(Paragraph(f"尚未匹配的附件（{len(unmatched)}）", heading))
    if unmatched:
        for att in unmatched:
            note = "；".join(att.get("scan_errors", []))
            story.append(Paragraph(escape(att["ref"] + ("（" + note + "）" if note else "")), body))
    else:
        story.append(Paragraph("无。", body))
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=40, rightMargin=40, topMargin=32, bottomMargin=32)
    doc.build(story)


def _copy_invoice(root: Path, record: dict[str, Any], dest: Path) -> None:
    data = read_source(root, record["ref"])
    pages = record.get("pages")
    if pages is not None:
        if fitz is None:
            raise RuntimeError("多票 PDF 拆分需要 PyMuPDF")
        with fitz.open(stream=data, filetype="pdf") as source:
            result = fitz.open()
            for page_no in pages:
                result.insert_pdf(source, from_page=page_no, to_page=page_no)
            data = result.tobytes()
            result.close()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(data)


def _unique_path(parent: Path, filename: str) -> Path:
    candidate = parent / filename
    if not candidate.exists():
        return candidate
    p = Path(filename)
    for index in range(2, 10000):
        candidate = parent / f"{p.stem}-{index}{p.suffix}"
        if not candidate.exists():
            return candidate
    raise RuntimeError("输出文件名冲突过多")


def _write_mapping(path: Path, records: list[dict[str, Any]], attachments: list[dict[str, Any]], targets_by_ref: dict[str, list[str]], ignored_refs: set[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["原文件", "发票号码", "状态", "Excel编号", "上一版编号", "输出文件", "关联发票ID", "问题", "SHA256"])
        writer.writeheader()
        for record in records:
            writer.writerow({
                "原文件": record["ref"], "发票号码": record.get("invoice_number_final") or "",
                "状态": STATUS_LABELS[record["status"]], "Excel编号": record.get("number") or "",
                "上一版编号": record.get("previous_number") or "", "输出文件": record.get("output_file") or "",
                "关联发票ID": record["id"], "问题": "见 Excel 备注" if record["status"] == "pending" else "；".join(record.get("issues", [])), "SHA256": record["sha256"],
            })
            for alias in record.get("aliases", []):
                writer.writerow({"原文件": alias, "发票号码": record.get("invoice_number_final") or "", "状态": "重复副本", "Excel编号": record.get("number") or "", "上一版编号": "", "输出文件": record.get("output_file") or "", "关联发票ID": record["id"], "问题": "与主文件 SHA256 相同", "SHA256": record["sha256"]})
        for att in attachments:
            status = "已确认非报销材料" if att["ref"] in ignored_refs else "附件"
            issue = "" if att["ref"] in ignored_refs or att["ref"] in targets_by_ref else "未匹配到唯一发票"
            writer.writerow({"原文件": att["ref"], "发票号码": "", "状态": status, "Excel编号": "", "上一版编号": "", "输出文件": att.get("output_file") or "", "关联发票ID": ";".join(targets_by_ref.get(att["ref"], [])), "问题": issue, "SHA256": att["sha256"]})


def build(draft: dict[str, Any]) -> dict[str, Any]:
    records, targets_by_ref = evaluate(draft)
    root = Path(draft["input_dir"]).resolve()
    output_root = Path(draft["output_dir"]).resolve()
    run_id = str(draft.get("run_id") or "")
    if not run_id:
        raise ValueError("review.json 缺少本次运行标识，请重新 scan 后构建")
    if not root.is_dir():
        raise FileNotFoundError(root)
    output_root.mkdir(parents=True, exist_ok=True)
    state_path = output_root / "state.json"
    prior = json.loads(state_path.read_text(encoding="utf-8")) if state_path.exists() else {}
    same_run = bool(prior.get("version") and prior.get("run_id") == run_id)
    version = int(prior.get("version", 0)) + (0 if same_run else 1)
    final_dir = output_root / f"v{version}"
    zip_path = output_root / f"{_safe_name(draft['profile']['name'])}-集中报销-v{version}.zip"
    if not same_run and (final_dir.exists() or zip_path.exists()):
        raise FileExistsError(f"v{version} 已存在但不属于本次运行；请核对 state.json，不自动跳号")
    stage = output_root / f".v{version}.building"
    staged_zip = output_root / f".v{version}.building.zip"
    if stage.exists() or staged_zip.exists():
        raise FileExistsError(f"v{version} 的构建暂存文件已存在，请先核查，不自动覆盖")
    stage.mkdir()
    try:
        sorted_records = sorted(records, key=_sort_key)
        _number_repeated_names(sorted_records)
        formal = [r for r in sorted_records if r["status"] in FORMAL_STATUSES]
        for index, record in enumerate(formal, 1):
            record["number"] = index
            amount = money(record["claim_amount"])
            amount_label = f"{_compact_money(amount)}元" if amount is not None else "金额待核"
            name = f"{index}-{_safe_name(record['description'], 34)}-{amount_label}{record['ext']}"
            path = _unique_path(stage / "发票", name)
            _copy_invoice(root, record, path)
            record["output_file"] = path.relative_to(stage).as_posix()
        for record in sorted_records:
            if record["status"] != "ineligible":
                continue
            original_name = _safe_name(Path(record["ref"].split("::")[-1]).name)
            path = stage / "不可报销" / record["id"] / original_name
            _copy_invoice(root, record, path)
            record["output_file"] = path.relative_to(stage).as_posix()
        attachments = draft.get("attachments", [])
        ignored_refs = set(draft.get("ignored_attachment_refs", []))
        for att in attachments:
            if att["ref"] in ignored_refs:
                continue
            target_ids = targets_by_ref.get(att["ref"], [])
            number = next((r.get("number") for r in sorted_records if r["id"] in target_ids and r.get("number")), None)
            subdir = stage / "附件" / (str(number) if number else "未匹配")
            path = _unique_path(subdir, _safe_name(_attachment_name(att)))
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(read_source(root, att["ref"]))
            att["output_file"] = path.relative_to(stage).as_posix()
        profile = draft["profile"]
        workbook_path = stage / f"{_safe_name(profile['name'])}-报销.xlsx"
        _write_workbook(workbook_path, profile, sorted_records)
        active_attachments = [a for a in attachments if a["ref"] not in ignored_refs]
        _write_pdf(stage / "补件与复核提醒.pdf", profile, sorted_records, active_attachments, targets_by_ref, version)
        _write_mapping(stage / "文件对照表.csv", sorted_records, attachments, targets_by_ref, ignored_refs)
        changes = [r for r in sorted_records if r.get("previous_number") and r.get("number") and r["previous_number"] != r["number"]]
        old_by_id = {r["id"]: r for r in sorted_records}
        replacements = [
            (old_by_id.get(r.get("manual", {}).get("replacement_for")), r)
            for r in sorted_records if r.get("manual", {}).get("replacement_for") and r.get("number")
        ]
        with (stage / "编号变更.csv").open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(["变更类型", "发票ID", "原文件", "旧编号", "新编号"])
            for r in changes:
                writer.writerow(["重新编号", r["id"], r["ref"], r["previous_number"], r["number"]])
            for old, new in replacements:
                if old and old.get("previous_number"):
                    writer.writerow(["更正发票替换", new["id"], new["ref"], old["previous_number"], new["number"]])
        with ZipFile(staged_zip, "x", compression=ZIP_DEFLATED, compresslevel=6) as archive:
            for path in sorted(stage.rglob("*")):
                if path.is_file():
                    archive.write(path, (final_dir.name + "/" + path.relative_to(stage).as_posix()))
        archive_dir = None
        if same_run and (final_dir.exists() or zip_path.exists()):
            drafts = output_root / ".drafts" / run_id
            attempt = 1
            while (drafts / f"attempt-{attempt}").exists():
                attempt += 1
            archive_dir = drafts / f"attempt-{attempt}"
            archive_dir.mkdir(parents=True)
            if final_dir.exists():
                final_dir.replace(archive_dir / final_dir.name)
            if zip_path.exists():
                zip_path.replace(archive_dir / zip_path.name)
        try:
            stage.replace(final_dir)
            staged_zip.replace(zip_path)
        except OSError:
            if archive_dir is not None:
                if final_dir.exists():
                    final_dir.replace(archive_dir / f"incomplete-{final_dir.name}")
                if zip_path.exists():
                    zip_path.replace(archive_dir / f"incomplete-{zip_path.name}")
                if (archive_dir / final_dir.name).exists():
                    (archive_dir / final_dir.name).replace(final_dir)
                if (archive_dir / zip_path.name).exists():
                    (archive_dir / zip_path.name).replace(zip_path)
            raise
        state = {
            "schema_version": 1, "version": version, "run_id": run_id,
            "input_dir": str(root), "output_dir": str(output_root),
            "history_dir": draft.get("history_dir"), "history_invoice_numbers": draft.get("history_invoice_numbers", []),
            "profile": profile, "public_invoice_numbers": draft.get("public_invoice_numbers", []),
            "public_record_ids": draft.get("public_record_ids", []),
            "ignored_attachment_refs": draft.get("ignored_attachment_refs", []),
            "records": sorted_records, "attachments": attachments,
            "reference_files": draft.get("reference_files", []), "last_zip": str(zip_path),
        }
        write_json(output_root / "state.json", state)
        counts = {status: sum(r["status"] == status for r in sorted_records) for status in STATUS_LABELS}
        return {"version": version, "revised_same_run": same_run, "zip": str(zip_path), "directory": str(final_dir), "counts": counts, "formal_count": len(formal)}
    except Exception:
        # Keep an incomplete stage for diagnosis rather than deleting user data.
        raise
