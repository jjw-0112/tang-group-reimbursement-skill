"""Read-only inventory and invoice extraction for the reimbursement skill.

The scanner deliberately separates observations from decisions.  Its JSON output
is reviewable, and all source files are reopened from their original location
when a package is built.
"""

from __future__ import annotations

import hashlib
import io
import json
import re
import shutil
import subprocess
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path, PurePosixPath
from typing import Any, Iterator
from zipfile import BadZipFile, ZipFile
from xml.etree import ElementTree as ET

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover - doctor reports this
    fitz = None

try:
    from pypdf import PdfReader
except ImportError:  # pragma: no cover - doctor reports this
    PdfReader = None

try:
    from PIL import Image
except ImportError:  # pragma: no cover - doctor reports this
    Image = None


SCHEMA_VERSION = 1
MAX_MEMBER_BYTES = 100 * 1024 * 1024
MAX_ARCHIVE_BYTES = 500 * 1024 * 1024
SUPPORTED = {".pdf", ".ofd", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".docx", ".txt", ".xlsx"}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff"}
ATTACHMENT_HINTS = (
    "行程单", "行程报销单", "运单", "运费清单", "打印清单", "打印明细", "支付记录",
    "付款记录", "付款凭证", "订单详情", "订单截图", "入库单", "高值材料", "高值耗材",
    "审批表", "录用通知", "审稿通知", "缴费通知", "收费通知", "论文首页", "测试报告",
    "测试清单", "合同", "会议邀请函", "会议通知", "地铁行程", "行程列表", "销货清单",
)
CATEGORY_ORDER = [
    "材料费", "快递费", "交通费", "书籍", "差旅费", "文印费", "测试化验加工",
    "版面费/审稿费", "查收查引/专利", "其他", "对公转账",
]
MONEY = r"\d[\d,]*(?:\.\d{1,2})?"
NUMERIC_ROW = re.compile(
    r"(?P<qty>-?\d+(?:\.\d+)?)\s+"
    r"(?P<unit>-?\d+(?:\.\d+)?)\s+"
    r"(?P<net>-?\d+(?:\.\d+)?)\s+"
    r"(?P<rate>\d+(?:\.\d+)?%|免税|不征税)\s+"
    r"(?P<tax>-?\d+(?:\.\d+)?)(?![\d.])"
)
SHORT_ROW = re.compile(
    r"(?P<net>-?\d+(?:\.\d+)?)\s+"
    r"(?P<rate>\d+(?:\.\d+)?%|免税|不征税)\s+"
    r"(?P<tax>-?\d+(?:\.\d+)?)(?![\d.])"
)


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def money(value: Any) -> Decimal | None:
    if value is None or value == "":
        return None
    try:
        return Decimal(str(value).replace(",", "")).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return None


def money_str(value: Decimal | None) -> str | None:
    return format(value, ".2f") if value is not None else None


def natural_key(value: str) -> tuple[Any, ...]:
    return tuple(int(part) if part.isdigit() else part.casefold() for part in re.split(r"(\d+)", value))


def _safe_member(member: str) -> bool:
    member = member.replace("\\", "/")
    p = PurePosixPath(member)
    return not p.is_absolute() and ".." not in p.parts and not any(":" in part for part in p.parts)


def _archive_members_7z(archive: Path) -> list[tuple[str, int]]:
    exe = shutil.which("7z") or shutil.which("7za") or shutil.which("7zz")
    if not exe:
        raise RuntimeError("未安装 7-Zip/7z，无法读取 RAR")
    proc = subprocess.run(
        [exe, "l", "-slt", "-sccUTF-8", str(archive)], capture_output=True,
        text=True, encoding="utf-8", errors="replace", timeout=90,
    )
    if proc.returncode != 0:
        raise RuntimeError("RAR 列表读取失败：" + (proc.stderr or proc.stdout)[-300:])
    blocks = proc.stdout.split("----------", 1)
    if len(blocks) != 2:
        raise RuntimeError("RAR 列表格式无法识别")
    result: list[tuple[str, int]] = []
    for block in re.split(r"\r?\n\s*\r?\n", blocks[1]):
        fields = dict(re.findall(r"^([^\r\n=]+) = (.*)$", block, re.MULTILINE))
        name = fields.get("Path")
        if name and fields.get("Folder") == "-":
            try:
                size = int(fields.get("Size", "0"))
            except ValueError:
                size = 0
            result.append((name, size))
    return result


def _read_rar_member(archive: Path, member: str) -> bytes:
    exe = shutil.which("7z") or shutil.which("7za") or shutil.which("7zz")
    if not exe:
        raise RuntimeError("未安装 7-Zip/7z，无法读取 RAR")
    proc = subprocess.run(
        [exe, "x", "-so", "-sccUTF-8", str(archive), member],
        capture_output=True, timeout=120,
    )
    if proc.returncode != 0:
        raise RuntimeError("RAR 成员提取失败：" + proc.stderr.decode("utf-8", "replace")[-300:])
    if len(proc.stdout) > MAX_MEMBER_BYTES:
        raise RuntimeError("RAR 成员超过单文件大小上限")
    return proc.stdout


def iter_sources(root: Path, output_dir: Path | None = None) -> Iterator[dict[str, Any]]:
    root = root.resolve()
    output_dir = output_dir.resolve() if output_dir else None
    for path in sorted(root.rglob("*"), key=lambda p: natural_key(str(p.relative_to(root)))):
        if not path.is_file() or path.name.startswith("~$") or path.name in {"Thumbs.db", ".DS_Store"}:
            continue
        if output_dir and (path == output_dir or output_dir in path.parents):
            continue
        rel = path.relative_to(root).as_posix()
        if path.stat().st_size > MAX_ARCHIVE_BYTES:
            yield {"ref": rel, "ext": path.suffix.lower(), "data": b"", "error": "文件超过大小上限"}
            continue
        ext = path.suffix.lower()
        if ext == ".zip":
            try:
                total = 0
                with ZipFile(path) as archive:
                    for info in archive.infolist():
                        if info.is_dir():
                            continue
                        member = info.filename
                        if not _safe_member(member):
                            yield {"ref": rel + "::" + member, "ext": Path(member).suffix.lower(), "data": b"", "error": "压缩包内路径不安全"}
                            continue
                        total += info.file_size
                        if info.file_size > MAX_MEMBER_BYTES or total > MAX_ARCHIVE_BYTES:
                            yield {"ref": rel + "::" + member, "ext": Path(member).suffix.lower(), "data": b"", "error": "压缩包内容超过大小上限"}
                            continue
                        yield {"ref": rel + "::" + member, "ext": Path(member).suffix.lower(), "data": archive.read(info)}
            except (BadZipFile, OSError, RuntimeError) as exc:
                yield {"ref": rel, "ext": ext, "data": b"", "error": f"ZIP 无法读取：{exc}"}
        elif ext == ".rar":
            try:
                total = 0
                for member, size in _archive_members_7z(path):
                    if not _safe_member(member):
                        yield {"ref": rel + "::" + member, "ext": Path(member).suffix.lower(), "data": b"", "error": "压缩包内路径不安全"}
                        continue
                    total += size
                    if size > MAX_MEMBER_BYTES or total > MAX_ARCHIVE_BYTES:
                        yield {"ref": rel + "::" + member, "ext": Path(member).suffix.lower(), "data": b"", "error": "压缩包内容超过大小上限"}
                        continue
                    yield {"ref": rel + "::" + member, "ext": Path(member).suffix.lower(), "data": _read_rar_member(path, member)}
            except (OSError, RuntimeError, subprocess.TimeoutExpired) as exc:
                yield {"ref": rel, "ext": ext, "data": b"", "error": f"RAR 无法读取：{exc}"}
        else:
            try:
                yield {"ref": rel, "ext": ext, "data": path.read_bytes()}
            except OSError as exc:
                yield {"ref": rel, "ext": ext, "data": b"", "error": f"文件无法读取：{exc}"}


def read_source(root: Path, ref: str) -> bytes:
    outer, sep, member = ref.partition("::")
    path = (root / PurePosixPath(outer)).resolve()
    if root.resolve() not in path.parents and path != root.resolve():
        raise ValueError("来源路径超出输入目录")
    if not sep:
        return path.read_bytes()
    if not _safe_member(member):
        raise ValueError("压缩包成员路径不安全")
    if path.suffix.lower() == ".zip":
        with ZipFile(path) as archive:
            info = archive.getinfo(member)
            if info.file_size > MAX_MEMBER_BYTES:
                raise ValueError("压缩包成员超过大小上限")
            return archive.read(info)
    if path.suffix.lower() == ".rar":
        raw = next((n for n, _ in _archive_members_7z(path) if n.replace("\\", "/") == member.replace("\\", "/")), None)
        if raw is None:
            raise FileNotFoundError(ref)
        return _read_rar_member(path, raw)
    raise ValueError("不支持的压缩包类型")


def _ocr_image_bytes(data: bytes) -> tuple[str, str | None]:
    exe = shutil.which("tesseract")
    if not exe:
        return "", "未安装 Tesseract OCR"
    if Image is None:
        return "", "缺少 Pillow，无法读取图片"
    try:
        image = Image.open(io.BytesIO(data)).convert("RGB")
        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        proc = subprocess.run(
            [exe, "stdin", "stdout", "-l", "chi_sim+eng"], input=buffer.getvalue(),
            capture_output=True, timeout=120,
        )
        if proc.returncode:
            return "", "OCR 失败或缺少 chi_sim/eng 语言包"
        return proc.stdout.decode("utf-8", "replace"), None
    except Exception as exc:
        return "", f"OCR 失败：{exc}"


def _pdf_pages(data: bytes) -> tuple[list[str], list[str]]:
    errors: list[str] = []
    texts: list[str] = []
    if fitz is not None:
        try:
            with fitz.open(stream=data, filetype="pdf") as doc:
                for page in doc:
                    text = page.get_text(sort=True)
                    if len(text.strip()) < 35:
                        png = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False).tobytes("png")
                        ocr, error = _ocr_image_bytes(png)
                        if ocr.strip():
                            text = ocr
                        if error:
                            errors.append(error)
                    texts.append(text)
            return texts, errors
        except Exception as exc:
            errors.append(f"PDF 读取失败：{exc}")
    if PdfReader is not None:
        try:
            reader = PdfReader(io.BytesIO(data))
            texts = [page.extract_text() or "" for page in reader.pages]
            return texts, errors
        except Exception as exc:
            errors.append(f"备用 PDF 提取失败：{exc}")
    return [""], errors or ["缺少 PDF 读取工具"]


def _ofd_text(data: bytes) -> tuple[str, str | None]:
    try:
        with ZipFile(io.BytesIO(data)) as archive:
            parts = []
            for name in sorted(archive.namelist()):
                if not name.lower().endswith(".xml") or archive.getinfo(name).file_size > 8 * 1024 * 1024:
                    continue
                try:
                    root = ET.fromstring(archive.read(name))
                except ET.ParseError:
                    continue
                for element in root.iter():
                    if element.tag.rsplit("}", 1)[-1] in {"TextCode", "Value"} and element.text:
                        parts.append(element.text)
            return "\n".join(parts), None
    except Exception as exc:
        return "", f"OFD 无法读取：{exc}"


def _docx_text(data: bytes) -> str:
    try:
        with ZipFile(io.BytesIO(data)) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
            return " ".join(el.text or "" for el in root.iter() if el.tag.endswith("}t"))
    except Exception:
        return ""


def _invoice_numbers(text: str) -> list[str]:
    values = re.findall(r"发票号码\s*[:：]?\s*([0-9][0-9\s]{7,27})", text)
    values += re.findall(r"票据号码\s*[:：]?\s*([0-9]{8,20})", text)
    result = []
    for value in values:
        digits = re.sub(r"\D", "", value)
        # Some PDFs concatenate a 20-digit number with OCR/overlay residue.
        if len(digits) >= 20:
            digits = digits[:20]
        if 8 <= len(digits) <= 20 and digits not in result:
            result.append(digits)
    return result


def _issue_date(text: str) -> str | None:
    match = re.search(r"开票日期\s*[:：]?\s*(20\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})", text)
    if not match:
        # In some layouts the invoice number is interposed after 开票日期.
        match = re.search(r"(20\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})", text[:500])
    if not match:
        return None
    try:
        return date(*(int(value) for value in match.groups())).isoformat()
    except ValueError:
        return None


def _invoice_amount(text: str) -> Decimal | None:
    matches = re.findall(r"[（(]?小\s*写[）)]?\s*[¥￥]?\s*(" + MONEY + r")", text)
    if matches:
        return money(matches[-1])
    tail = re.search(r"价\s*税\s*合\s*计[\s\S]{0,160}", text)
    if tail:
        matches = re.findall(r"[¥￥]\s*(" + MONEY + r")", tail.group(0))
        if matches:
            return money(matches[-1])
    for pattern in (r"票价\s*[:：]?\s*[¥￥]?\s*(" + MONEY + r")", r"金额\s*[:：]?[¥￥]\s*(" + MONEY + r")"):
        match = re.search(pattern, text)
        if match:
            return money(match.group(1))
    return None


def _buyer(text: str) -> tuple[str | None, str | None, bool]:
    """Read the buyer column, without treating a seller name/ID as the buyer."""
    head = text[:2000]
    buyer_label = r"(?:购\s*买\s*方?|购|买)\s*名\s*称\s*[:：]"
    buyer_line = next((line for line in head.splitlines() if re.search(buyer_label, line)), "")
    buyer_match = re.search(buyer_label + r"\s*(.*?)(?=\s+(?:销|售)\s*名\s*称|$)", buyer_line)
    buyer_name = buyer_match.group(1).strip() if buyer_match else None
    if not buyer_name and "购" in head[:500] and "销" in head[:500]:
        # Some PDF text layers place 购/销 on adjacent lines. In that layout
        # the first of two horizontally separated 名称 fields is the buyer.
        for line in head.splitlines()[:18]:
            pair = re.search(r"名\s*称\s*[:：]\s*(\S+?)\s{2,}名\s*称\s*[:：]", line)
            if pair:
                buyer_name = pair.group(1)
                break
    buyer_name = buyer_name or None
    # In the standard two-column VAT layout, the first tax ID on this row is
    # the buyer's. Do not search the entire invoice for the expected school ID.
    tax_line = next((line for line in head.splitlines() if "识别号" in line or "信用代码" in line), "")
    ids = re.findall(r"[0-9A-Z]{18}", tax_line)
    tax_id = ids[0] if ids else None
    return buyer_name, tax_id, bool(buyer_name)


def _item_lines(text: str, amount: Decimal | None) -> tuple[list[dict[str, str]], bool]:
    start = re.search(r"项\s*目\s*名\s*称", text)
    if not start:
        return [], False
    end = re.search(r"合\s*计", text[start.end():])
    body = text[start.end():start.end() + end.start()] if end else text[start.end():]
    items: list[dict[str, str]] = []
    for block in re.split(r"(?=^\s*\*)", body, flags=re.MULTILINE):
        block = block.strip()
        if not block.startswith("*"):
            continue
        match = NUMERIC_ROW.search(block)
        short = SHORT_ROW.search(block) if not match else None
        if not match and not short:
            continue
        try:
            qty = Decimal(match.group("qty")) if match else None
            unit_net = Decimal(match.group("unit")) if match else None
            source = match or short
            assert source is not None
            net = Decimal(source.group("net"))
            tax = Decimal(source.group("tax"))
            gross = (net + tax).quantize(Decimal("0.01"))
            unit_gross = (gross / qty).quantize(Decimal("0.01")) if qty else None
        except (InvalidOperation, ZeroDivisionError):
            continue
        title = re.sub(r"\s+", " ", block[:source.start()]).strip(" *")[:80]
        if short and net < 0 and items:
            prior = items[-1]
            prior_net = Decimal(prior["line_net"]) + net
            prior_tax = Decimal(prior["line_tax"]) + tax
            prior_gross = prior_net + prior_tax
            prior["line_net"] = money_str(prior_net)
            prior["line_tax"] = money_str(prior_tax)
            prior["line_gross"] = money_str(prior_gross)
            if prior.get("quantity"):
                prior["unit_gross"] = money_str(prior_gross / Decimal(prior["quantity"]))
            continue
        items.append({
            "name": title or "商品明细", "quantity": str(qty) if qty is not None else None,
            "unit_net": money_str(unit_net) if unit_net is not None else None,
            "unit_gross": money_str(unit_gross) if unit_gross is not None else None,
            "line_net": money_str(net), "line_tax": money_str(tax), "line_gross": money_str(gross),
        })
    verified = bool(items and amount is not None and abs(sum(Decimal(i["line_gross"]) for i in items) - amount) <= Decimal("0.03"))
    return items, verified


def _filename_amount(ref: str) -> Decimal | None:
    stem = Path(ref.split("::")[-1]).stem
    match = re.search(r"(\d+(?:\.\d{1,2})?)\s*元?$", stem)
    if not match or len(match.group(1).split(".")[0]) > 6:
        return None
    return money(match.group(1))


def _attachment_type(name: str, text: str) -> str | None:
    sample = (name + " " + text[:200]).lower()
    kinds = [
        ("高值材料", "high_value"), ("高值耗材", "high_value"), ("入库", "inventory"),
        ("运单", "waybill"), ("运费清单", "waybill"),
        ("地铁行程", "metro_trip"), ("行程列表", "metro_trip"),
        ("行程单", "itinerary"), ("行程报销单", "itinerary"),
        ("打印清单", "print_list"), ("打印明细", "print_list"),
        ("打印内容证明", "print_proof"),
        ("支付记录", "payment"), ("付款记录", "payment"), ("付款凭证", "payment"),
        ("缴费通知", "fee_notice"), ("收费通知", "fee_notice"),
        ("录用通知", "acceptance"), ("审稿通知", "acceptance"),
        ("论文首页", "paper_first_page"),
        ("测试报告", "test_report"), ("测试清单", "test_report"),
        ("销货清单", "sales_list"), ("会议通知", "meeting_notice"), ("会议邀请函", "meeting_notice"),
        ("订单", "order"), ("合同", "contract"),
    ]
    for marker, kind in kinds:
        if marker.lower() in sample:
            return kind
    return None


def _payment_amount(text: str) -> Decimal | None:
    values = re.findall(r"(?:实付金额|支付金额|付款金额|支付总额|付款总额)\s*[:：]?\s*[¥￥]?\s*(" + MONEY + r")", text)
    parsed = {money(value) for value in values}
    return next(iter(parsed)) if len(parsed) == 1 else None


def _is_invoice(name: str, text: str, ext: str) -> bool:
    if any(word in name for word in ATTACHMENT_HINTS) and "航空运输电子客票行程单" not in text:
        return False
    if "电子发票" in text or "增值税发票" in text or "发票号码" in text:
        return True
    if any(word in text for word in ("出租车票", "铁路电子客票", "航空运输电子客票")):
        return True
    if "发票" in name and ext in {".pdf", ".ofd", *IMAGE_EXTS}:
        return True
    if re.match(r"^\d+[-_－]", name) and ext in {".pdf", ".ofd", *IMAGE_EXTS}:
        return True
    return False


def _invoice_kind(name: str, text: str) -> str:
    if any(marker in text + name for marker in ("铁路电子客票", "火车票", "航空运输电子客票", "机票", "出租车票")) and "电子发票（普通发票）" not in text:
        return "ticket"
    if "电子发票" in text or "增值税" in text or "价税合计" in text:
        return "vat"
    if "票据号码" in text or "统一票据" in text:
        return "receipt"
    return "unknown"


def _category(name: str, text: str, items: list[dict[str, str]]) -> tuple[str, list[str], bool]:
    sample = (name + " " + text[:2500]).lower()
    if "学会会员" in sample or "会费" in sample:
        return "其他", [], False
    if "测试" in name and "测试架" not in name:
        return "测试化验加工", [], False
    explicit_names = [
        ("快递费", ("快递", "顺丰", "邮寄")),
        ("交通费", ("网约车", "滴滴", "地铁", "出租车")),
        ("差旅费", ("火车票", "高铁", "机票", "住宿", "酒店")),
        ("文印费", ("打印费", "文印费", "复印费")),
        ("版面费/审稿费", ("版面费", "审稿费")),
    ]
    for category, markers in explicit_names:
        if any(marker in name for marker in markers):
            return category, [], False
    if items:
        totals: dict[str, Decimal] = {}
        for item in items:
            product = item.get("name", "").split("*")[-1]
            if any(x in product for x in ("运费", "收派服务", "邮寄费", "物流服务")):
                category = "快递费"
            elif any(x in product for x in ("客运服务", "网约车", "地铁", "出租车")):
                category = "交通费"
            elif any(x in product for x in ("图书", "书籍")):
                category = "书籍"
            elif any(x in product for x in ("打印", "复印", "文印")):
                category = "文印费"
            elif any(x in product for x in ("测试费", "检测费", "化验费")):
                category = "测试化验加工"
            elif any(x in product for x in ("版面费", "审稿费")):
                category = "版面费/审稿费"
            else:
                category = "材料费"
            totals[category] = totals.get(category, Decimal("0")) + (money(item.get("line_gross")) or Decimal("0"))
        positive = {key: value for key, value in totals.items() if value > 0}
        if positive:
            main = max(positive, key=positive.get)
            total = sum(positive.values())
            mixed = sorted(positive) if len(positive) > 1 else []
            clear = positive[main] / total >= Decimal("0.80") if total else True
            return main, mixed, not clear
    rules = [
        ("版面费/审稿费", ("版面费", "审稿费", "article processing", "publication fee", "open access fee")),
        ("查收查引/专利", ("查收查引", "查引", "专利费", "专利年费")),
        ("快递费", ("快递", "邮寄", "顺丰", "收派服务", "运费")),
        ("交通费", ("网约车", "滴滴", "出租车", "地铁", "客运服务", "高德打车")),
        ("差旅费", ("火车票", "高铁", "铁路电子客票", "机票", "航空运输", "酒店", "住宿费", "会议注册费")),
        ("书籍", ("图书", "书籍", "印刷品")),
        ("文印费", ("打印费", "复印费", "文印费")),
        ("测试化验加工", ("测试费", "化验费", "检测费", "检验费")),
    ]
    for category, markers in rules:
        if any(marker in sample for marker in markers):
            return category, [], False
    return ("材料费" if "项目名称" in text or "电子发票" in text else "其他"), [], False


def _description(ref: str, category: str, items: list[dict[str, str]]) -> str:
    stem = Path(ref.split("::")[-1]).stem
    stem = re.sub(r"^\d+[-_－ ]+", "", stem)
    stem = re.sub(r"[-_－ ]?\d+(?:\.\d{1,2})?\s*元?$", "", stem)
    stem = re.sub(r"电子发票|普通发票|发票", "", stem).strip(" -_－")
    if stem and len(stem) <= 24:
        return stem
    if items:
        title = items[0]["name"].split("*")[-1].strip()
        if title:
            return title[:20]
    return category


def _extract_one(ref: str, ext: str, data: bytes, text: str, pages: list[int] | None, errors: list[str]) -> dict[str, Any]:
    name = Path(ref.split("::")[-1]).name
    numbers = _invoice_numbers(text)
    amount = _invoice_amount(text)
    items, verified = _item_lines(text, amount)
    category, mixed_categories, mixed_ambiguous = _category(name, text, items)
    buyer_name, buyer_tax_id, buyer_explicit = _buyer(text)
    return {
        "id": sha256((ref + "#" + str(pages) + "#" + sha256(data)).encode("utf-8"))[:16],
        "ref": ref, "aliases": [], "sha256": sha256(data), "ext": ext,
        "pages": pages, "kind": "invoice", "invoice_kind": _invoice_kind(name, text),
        "invoice_number": numbers[0] if len(numbers) == 1 else None,
        "multiple_numbers": numbers if len(numbers) > 1 else [],
        "issue_date": _issue_date(text), "buyer_name": buyer_name, "buyer_tax_id": buyer_tax_id,
        "buyer_explicit": buyer_explicit,
        "amount": money_str(amount), "filename_amount": money_str(_filename_amount(ref)),
        "items": items, "items_verified": verified,
        "category_guess": category, "mixed_categories": mixed_categories,
        "mixed_ambiguous": mixed_ambiguous,
        "description_guess": _description(ref, category, items),
        "text_excerpt": text[:1800], "scan_errors": sorted(set(errors)), "manual": {},
    }


def analyze_source(source: dict[str, Any]) -> list[dict[str, Any]]:
    ref, ext, data = source["ref"], source["ext"], source["data"]
    errors = [source["error"]] if source.get("error") else []
    if not data:
        return [_extract_one(ref, ext, data, "", None, errors or ["空文件"])]
    name = Path(ref.split("::")[-1]).name
    if ext in {".xlsx", ".docx"} and any(marker in name for marker in ("报销清单", "报销明细", "个人清单", "总发票统计", "-报销.xlsx", "报销.xlsx")):
        return [{"id": sha256((ref + sha256(data)).encode("utf-8"))[:16], "ref": ref,
                 "aliases": [], "sha256": sha256(data), "ext": ext, "kind": "reference",
                 "scan_errors": errors}]
    if ext == ".pdf":
        pages, page_errors = _pdf_pages(data)
        errors.extend(page_errors)
        ids_by_page = [_invoice_numbers(t) for t in pages]
        if any(len(ids) > 1 for ids in ids_by_page):
            return [_extract_one(ref, ext, data, "\n".join(pages), None, errors + ["单页含多张发票，无法可靠拆分"])]
        unique = list(dict.fromkeys(ids[0] for ids in ids_by_page if ids))
        if len(unique) > 1:
            groups: list[tuple[list[int], list[str]]] = []
            current_id: str | None = None
            for index, page_text in enumerate(pages):
                page_id = ids_by_page[index][0] if ids_by_page[index] else current_id
                if page_id is not None and page_id != current_id:
                    groups.append(([], []))
                    current_id = page_id
                if not groups:
                    groups.append(([], []))
                groups[-1][0].append(index)
                groups[-1][1].append(page_text)
            return [_extract_one(ref, ext, data, "\n".join(texts), indices, errors) for indices, texts in groups]
        text = "\n".join(pages)
    elif ext == ".ofd":
        text, error = _ofd_text(data)
        if error:
            errors.append(error)
    elif ext in IMAGE_EXTS:
        text, error = _ocr_image_bytes(data)
        if error:
            errors.append(error)
        elif not text.strip():
            errors.append("图片文字无法识别")
    elif ext == ".docx":
        text = _docx_text(data)
    elif ext == ".txt":
        text = data.decode("utf-8", "replace")
    else:
        text = ""
        if ext not in SUPPORTED:
            errors.append("暂不支持此文件格式")
    attachment_named = any(marker in name for marker in ATTACHMENT_HINTS)
    if _is_invoice(name, text, ext) or errors and ext in {".pdf", ".ofd", ".zip", ".rar", *IMAGE_EXTS} and not attachment_named:
        return [_extract_one(ref, ext, data, text, None, errors)]
    attachment_type = _attachment_type(name, text)
    return [{
        "id": sha256((ref + "#" + sha256(data)).encode("utf-8"))[:16],
        "ref": ref, "aliases": [], "sha256": sha256(data), "ext": ext,
        "kind": "attachment", "attachment_type": attachment_type,
        "payment_amount": money_str(_payment_amount(text)) if attachment_type == "payment" else None,
        "text_excerpt": text[:1200], "scan_errors": sorted(set(errors)),
    }]


def _history_numbers(root: Path | None) -> list[str]:
    if root is None or not root.exists():
        return []
    numbers: set[str] = set()
    for source in iter_sources(root):
        if source["ext"] not in {".pdf", ".ofd"} or not source["data"]:
            continue
        if source["ext"] == ".pdf":
            pages, _ = _pdf_pages(source["data"])
            text = "\n".join(pages)
        else:
            text, _ = _ofd_text(source["data"])
        numbers.update(_invoice_numbers(text))
    return sorted(numbers)


def scan(input_dir: Path, output_dir: Path, history_dir: Path | None = None) -> dict[str, Any]:
    input_dir = input_dir.resolve()
    output_dir = output_dir.resolve()
    if not input_dir.is_dir():
        raise FileNotFoundError(input_dir)
    if input_dir == output_dir or output_dir in input_dir.parents:
        raise ValueError("输出目录不能等于或包含原材料目录")
    if history_dir and history_dir.resolve() == input_dir:
        raise ValueError("历史目录不能等于本次原材料目录")
    if history_dir and not history_dir.is_dir():
        raise FileNotFoundError(history_dir)
    prior_path = output_dir / "state.json"
    prior = json.loads(prior_path.read_text(encoding="utf-8")) if prior_path.exists() else {}
    review_path = output_dir / "review.json"
    previous_review = json.loads(review_path.read_text(encoding="utf-8")) if review_path.exists() else {}
    decision_source = previous_review or prior
    def record_key(record: dict[str, Any]) -> tuple[str | None, tuple[int, ...]]:
        return record.get("sha256"), tuple(record.get("pages") or [])
    prior_records = {record_key(r): r for r in decision_source.get("records", []) if r.get("sha256")}
    built_records = {record_key(r): r for r in prior.get("records", []) if r.get("sha256")}
    entries: list[dict[str, Any]] = []
    seen_hashes: dict[str, dict[str, Any]] = {}
    for source in iter_sources(input_dir, output_dir):
        for entry in analyze_source(source):
            if entry["sha256"] and entry["sha256"] in seen_hashes and entry.get("pages") is None:
                seen_hashes[entry["sha256"]]["aliases"].append(entry["ref"])
                continue
            old = prior_records.get(record_key(entry))
            if old and entry["kind"] == "invoice":
                entry["manual"] = old.get("manual", {})
            built_old = built_records.get(record_key(entry))
            if built_old and entry["kind"] == "invoice":
                entry["previous_number"] = built_old.get("number")
            entries.append(entry)
            if entry["sha256"] and entry.get("pages") is None:
                seen_hashes[entry["sha256"]] = entry
    records = [x for x in entries if x["kind"] == "invoice"]
    attachments = [x for x in entries if x["kind"] == "attachment"]
    references = [x for x in entries if x["kind"] == "reference"]
    old_by_number = {r.get("invoice_number"): r for r in prior.get("records", []) if r.get("invoice_number")}
    candidates = []
    for record in records:
        old = old_by_number.get(record.get("invoice_number"))
        if old and old.get("sha256") != record["sha256"]:
            candidates.append({"new_id": record["id"], "old_id": old["id"], "invoice_number": record["invoice_number"]})
    draft = {
        "schema_version": SCHEMA_VERSION, "input_dir": str(input_dir), "output_dir": str(output_dir),
        "history_dir": str(history_dir.resolve()) if history_dir else decision_source.get("history_dir"),
        "profile": decision_source.get("profile", {"name": "", "student_id": "", "phone": "", "cutoff_date": ""}),
        "public_invoice_numbers": decision_source.get("public_invoice_numbers", []),
        "public_record_ids": decision_source.get("public_record_ids", []),
        "ignored_attachment_refs": decision_source.get("ignored_attachment_refs", []),
        "records": records, "attachments": attachments, "reference_files": references,
        "history_invoice_numbers": _history_numbers(history_dir) if history_dir else decision_source.get("history_invoice_numbers", []),
        "replacement_candidates": candidates,
        "needs_review": [r["id"] for r in records if r["scan_errors"] or not r["invoice_number"] or not r["amount"] or not r["issue_date"] or r["multiple_numbers"]],
        "unreviewed_files": [
            {"ref": a["ref"], "reason": a["scan_errors"] or ["附件类型未识别，须确认是否属于报销材料"]}
            for a in attachments
            if a["ref"] not in set(decision_source.get("ignored_attachment_refs", []))
            and (a["scan_errors"] or not a.get("attachment_type"))
        ],
        "previous_version": prior.get("version", 0),
    }
    return draft


def write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temp.replace(path)
