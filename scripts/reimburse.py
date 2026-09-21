#!/usr/bin/env python3
"""Portable CLI for a reviewable, versioned personal reimbursement package."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

from engine import scan, write_json
from workflow import build


def doctor() -> dict[str, object]:
    modules = {}
    for name in ("fitz", "openpyxl", "PIL", "pypdf", "reportlab"):
        try:
            __import__(name)
            modules[name] = True
        except ImportError:
            modules[name] = False
    tesseract = shutil.which("tesseract")
    languages: list[str] = []
    if tesseract:
        try:
            result = subprocess.run([tesseract, "--list-langs"], capture_output=True, text=True, timeout=15)
            languages = [line.strip() for line in result.stdout.splitlines()[1:] if line.strip()]
        except (OSError, subprocess.TimeoutExpired):
            pass
    return {
        "python": sys.version.split()[0], "modules": modules,
        "tesseract": tesseract, "ocr_languages": languages,
        "ocr_ready": bool(tesseract and "chi_sim" in languages and "eng" in languages),
        "rar_extractor": shutil.which("7z") or shutil.which("7za") or shutil.which("7zz"),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="课题组个人集中报销文件整理")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("doctor", help="检查本地处理依赖")
    scan_parser = sub.add_parser("scan", help="只读扫描材料并生成可复核 review.json")
    scan_parser.add_argument("--input", required=True, type=Path, help="学生原材料目录")
    scan_parser.add_argument("--output", type=Path, help="输出目录，默认在原材料目录旁")
    scan_parser.add_argument("--history", type=Path, help="可选历史报销目录")
    build_parser = sub.add_parser("build", help="依据已复核 review.json 生成新版本交件包")
    build_parser.add_argument("--review", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        if args.command == "doctor":
            result = doctor()
        elif args.command == "scan":
            input_dir = args.input.resolve()
            output_dir = args.output.resolve() if args.output else input_dir.with_name(input_dir.name + "_报销输出")
            draft = scan(input_dir, output_dir, args.history)
            review_path = output_dir / "review.json"
            write_json(review_path, draft)
            result = {
                "review": str(review_path), "invoices": len(draft["records"]),
                "attachments": len(draft["attachments"]), "needs_review": draft["needs_review"],
                "unreviewed_files": draft["unreviewed_files"],
                "replacement_candidates": draft["replacement_candidates"],
            }
        else:
            draft = json.loads(args.review.read_text(encoding="utf-8"))
            result = build(draft)
    except (OSError, ValueError, RuntimeError, KeyError, json.JSONDecodeError) as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
