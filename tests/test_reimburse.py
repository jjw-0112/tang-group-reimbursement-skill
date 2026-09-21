"""Behavioral checks for extraction, decisions, and versioned reruns."""

from __future__ import annotations

import csv
import hashlib
import io
import json
import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook
from PIL import Image
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from engine import _item_lines, _ofd_text, scan  # noqa: E402
from workflow import _item_thresholds, build, evaluate  # noqa: E402


BUYER = "西安交通大学"
TAX_ID = "12100000435230200R"
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))


def invoice_pdf(number: str, amount: str, product: str = "运费", buyer: str = BUYER) -> bytes:
    buf = io.BytesIO()
    pdf = canvas.Canvas(buf)
    pdf.setFont("STSong-Light", 11)
    lines = [
        "电子发票（普通发票） 发票号码：" + number,
        "开票日期：2026年06月01日",
        "购 名称：" + buyer + "       销 名称：某供应商",
        "统一社会信用代码/纳税人识别号：" + TAX_ID + "    统一社会信用代码/纳税人识别号：913000000000000001",
        "项目名称 规格型号 数量 单价 金额 税率 税额",
        f"*服务*{product} 1 {amount} {amount} 0% 0.00",
        "合计",
        "价税合计（小写）¥" + amount,
    ]
    for i, line in enumerate(lines):
        pdf.drawString(35, 790 - i * 24, line)
    pdf.save()
    return buf.getvalue()


class ReimbursementTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.input = self.root / "原件"
        self.output = self.root / "输出"
        self.input.mkdir()
        self.profile = {"name": "测试学生", "student_id": "123", "phone": "13800000000", "cutoff_date": "2026-06-15"}

    def draft(self) -> dict:
        result = scan(self.input, self.output)
        result["profile"] = self.profile
        return result

    def test_item_thresholds_apply_to_each_product(self) -> None:
        text = """项目名称 规格型号 数量 单价 金额 税率 税额
*材料*零件甲 2 150.00 300.00 0% 0.00
*材料*零件乙 1 300.00 300.00 0% 0.00
*材料*传感器 2 250.00 500.00 0% 0.00
*材料*探头 1 1000.00 1000.00 0% 0.00
合计"""
        items, verified = _item_lines(text, Decimal("2100.00"))
        self.assertTrue(verified)
        self.assertEqual(len(items), 4)
        record = {"items": items}
        inventory, high_value = _item_thresholds(record)
        self.assertEqual(len(inventory), 2)
        self.assertTrue(any("传感器" in item for item in inventory))
        self.assertFalse(any("零件甲" in item or "零件乙" in item for item in inventory))
        self.assertEqual(len(high_value), 1)
        self.assertIn("探头", high_value[0])

    def test_buyer_history_and_dates(self) -> None:
        (self.input / "1-材料费20.pdf").write_bytes(invoice_pdf("12345678901234567890", "20.00", "零件", "其他单位"))
        draft = self.draft()
        record = evaluate(draft)[0][0]
        self.assertEqual(record["status"], "ineligible")
        self.assertIn("抬头不符", "；".join(record["issues"]))
        draft["records"][0]["manual"]["buyer_name"] = BUYER
        draft["history_invoice_numbers"] = ["12345678901234567890"]
        record = evaluate(draft)[0][0]
        self.assertEqual(record["status"], "pending")
        self.assertIn("疑似重复", "；".join(record["issues"]))
        draft["history_invoice_numbers"] = []
        draft["records"][0]["manual"]["issue_date"] = "2025-12-31"
        record = evaluate(draft)[0][0]
        self.assertEqual(record["status"], "ready")
        self.assertIn("跨年度", "；".join(record["warnings"]))
        draft["records"][0]["manual"]["issue_date"] = "2026-06-16"
        self.assertEqual(evaluate(draft)[0][0]["status"], "pending")

    def test_public_payment_conflict_and_private_limit(self) -> None:
        (self.input / "1-材料费20.pdf").write_bytes(invoice_pdf("12345678901234567890", "20.00", "零件"))
        (self.input / "1-付款凭证.txt").write_text("支付金额：99.00 元", encoding="utf-8")
        draft = self.draft()
        draft["public_invoice_numbers"] = ["12345678901234567890"]
        record = evaluate(draft)[0][0]
        self.assertEqual(record["status"], "pending")
        self.assertIn("付款凭证金额", "；".join(record["issues"]))
        draft["records"][0]["manual"]["confirm_payment_amount"] = True
        self.assertEqual(evaluate(draft)[0][0]["status"], "ready")
        draft["public_invoice_numbers"] = []
        draft["records"][0]["manual"]["claim_amount"] = "1000.00"
        self.assertEqual(evaluate(draft)[0][0]["status"], "pending")
        (self.input / "1-材料费20.pdf").write_bytes(invoice_pdf("12345678901234567890", "1000.00", "零件"))
        draft = self.draft()
        self.assertEqual(evaluate(draft)[0][0]["status"], "ineligible")

    def test_missing_rerun_renumber_and_confirmed_replacement(self) -> None:
        first = self.input / "1-快递费50.pdf"
        first.write_bytes(invoice_pdf("12345678901234567890", "50.00"))
        original_hash = hashlib.sha256(first.read_bytes()).hexdigest()
        draft = self.draft()
        self.assertEqual(len(draft["records"]), 1)
        self.assertEqual(draft["records"][0]["amount"], "50.00")
        self.assertEqual(evaluate(draft)[0][0]["status"], "missing")
        package1 = build(draft)
        self.assertEqual(package1["version"], 1)
        with ZipFile(package1["zip"]) as archive:
            names = archive.namelist()
            self.assertTrue(any("1-快递费-50元.pdf" in name for name in names))
            self.assertTrue(any("补件与复核提醒.pdf" in name for name in names))
        wb = load_workbook(Path(package1["directory"]) / "测试学生-报销.xlsx")
        ws = wb.active
        self.assertEqual(ws["B1"].value, "测试学生")
        self.assertTrue(any("缺件" in str(row[3].value) for row in ws.iter_rows(min_row=2, max_col=4) if row[3].value))
        self.assertTrue(any(cell.data_type == "f" for row in ws for cell in row))

        (self.input / "1-运单.txt").write_text("运单：本票运费 50 元", encoding="utf-8")
        draft = self.draft()
        self.assertEqual(evaluate(draft)[0][0]["status"], "ready")
        package2 = build(draft)
        self.assertEqual(package2["version"], 2)
        self.assertTrue(Path(package1["zip"]).is_file())

        (self.input / "0-材料费20.pdf").write_bytes(invoice_pdf("22345678901234567890", "20.00", "零件"))
        draft = self.draft()
        package3 = build(draft)
        with (Path(package3["directory"]) / "编号变更.csv").open(encoding="utf-8-sig", newline="") as handle:
            changes = list(csv.DictReader(handle))
        self.assertTrue(any(row["变更类型"] == "重新编号" and row["旧编号"] == "1" and row["新编号"] == "2" for row in changes))

        (self.input / "2-快递费52.pdf").write_bytes(invoice_pdf("12345678901234567890", "52.00"))
        draft = self.draft()
        self.assertEqual(len(draft["replacement_candidates"]), 1)
        new_id = draft["replacement_candidates"][0]["new_id"]
        old_id = draft["replacement_candidates"][0]["old_id"]
        self.assertTrue(all(r["status"] == "pending" for r in evaluate(draft)[0] if r["invoice_number_final"] == "12345678901234567890"))
        next(r for r in draft["records"] if r["id"] == new_id)["manual"] = {
            "replacement_for": old_id, "attachment_refs": ["1-运单.txt"]
        }
        package4 = build(draft)
        state = json.loads((self.output / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(next(r for r in state["records"] if r["id"] == old_id)["status"], "replaced")
        self.assertEqual(next(r for r in state["records"] if r["id"] == new_id)["status"], "ready")
        formal = [r for r in state["records"] if r["status"] in {"ready", "missing"}]
        self.assertEqual(sorted(r["number"] for r in formal), list(range(1, len(formal) + 1)))
        with (Path(package4["directory"]) / "文件对照表.csv").open(encoding="utf-8-sig", newline="") as handle:
            mappings = list(csv.DictReader(handle))
        with ZipFile(package4["zip"]) as archive:
            names = set(archive.namelist())
            self.assertTrue(all("v4/" + row["输出文件"] in names for row in mappings if row["输出文件"]))
        with (Path(package4["directory"]) / "编号变更.csv").open(encoding="utf-8-sig", newline="") as handle:
            changes = list(csv.DictReader(handle))
        self.assertTrue(any(row["变更类型"] == "更正发票替换" and row["旧编号"] == "2" and row["新编号"] == "2" for row in changes))
        self.assertEqual(hashlib.sha256(first.read_bytes()).hexdigest(), original_hash)

    def test_zip_and_ofd(self) -> None:
        with ZipFile(self.input / "一批材料.zip", "w") as archive:
            archive.writestr("发票/1-材料费20.pdf", invoice_pdf("32345678901234567890", "20.00", "零件"))
        draft = self.draft()
        self.assertEqual(draft["records"][0]["amount"], "20.00")
        self.assertIn("::", draft["records"][0]["ref"])
        package = build(draft)
        self.assertEqual(package["formal_count"], 1)
        with ZipFile(package["zip"]) as archive:
            self.assertTrue(any(name.endswith(".pdf") and "/发票/" in name for name in archive.namelist()))
        ofd = io.BytesIO()
        with ZipFile(ofd, "w") as archive:
            archive.writestr("Doc_0/Pages/Page_0/Content.xml", "<Page><TextObject><TextCode>铁路电子客票 票价：289.00 发票号码：42345678901234567890 开票日期：2026年06月01日</TextCode></TextObject></Page>")
        text, error = _ofd_text(ofd.getvalue())
        self.assertIsNone(error)
        self.assertIn("289.00", text)

    def test_unreadable_image_remains_pending(self) -> None:
        Image.new("RGB", (220, 100), "white").save(self.input / "scan.png")
        draft = self.draft()
        self.assertEqual(len(draft["records"]), 1)
        self.assertEqual(len(draft["attachments"]), 0)
        self.assertEqual(evaluate(draft)[0][0]["status"], "pending")


if __name__ == "__main__":
    unittest.main()
