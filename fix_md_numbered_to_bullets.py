#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Optional


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Đổi các dòng '1. ...' lặp lại thành bullet '-' trong Markdown."
    )
    parser.add_argument(
        "input_path",
        type=Path,
        help="File .md hoặc thư mục chứa .md",
    )
    parser.add_argument(
        "output_path",
        type=Path,
        help="File .md xuất ra hoặc thư mục đích khi xử lý hàng loạt",
    )
    return parser.parse_args(argv)


def fix_content(text: str) -> str:
    # Chỉ đổi các dòng bắt đầu bằng '1.' thành '-'
    return re.sub(r"^(\s*)1\.\s+", r"\1- ", text, flags=re.MULTILINE)


def process_file(input_file: Path, output_file: Path) -> None:
    content = input_file.read_text(encoding="utf-8")
    fixed = fix_content(content)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    output_file.write_text(fixed, encoding="utf-8")


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    if args.input_path.is_dir():
        if args.output_path.exists() and not args.output_path.is_dir():
            print("Output phải là thư mục khi input là thư mục", file=sys.stderr)
            return 2
        inputs = sorted(p for p in args.input_path.iterdir() if p.suffix.lower() == ".md")
        if not inputs:
            print("Không tìm thấy file .md trong thư mục", file=sys.stderr)
            return 2
        args.output_path.mkdir(parents=True, exist_ok=True)
        for input_file in inputs:
            out_file = args.output_path / input_file.name
            process_file(input_file, out_file)
    else:
        if args.output_path.is_dir():
            args.output_path.mkdir(parents=True, exist_ok=True)
            out_file = args.output_path / args.input_path.name
        else:
            out_file = args.output_path
        process_file(args.input_path, out_file)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
