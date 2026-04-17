#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from pathlib import Path
from typing import List, Optional

LINE_EQ = "=========================="
LINE_DASH = "-----------------------------------"


def slugify(title: str) -> str:
    text = unicodedata.normalize("NFKD", title)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.strip()
    text = re.sub(r"[\\/:*?\"<>|]+", "-", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or "untitled"


def split_blocks(lines: List[str]) -> List[List[str]]:
    blocks: List[List[str]] = []
    current: List[str] = []
    for line in lines:
        if line.strip() == LINE_EQ and current:
            current.append(line)
            blocks.append(current)
            current = []
        else:
            current.append(line)
    if current:
        blocks.append(current)
    return blocks


def extract_title(block: List[str]) -> Optional[str]:
    # Expected pattern:
    # ==========================
    # -----------------------------------
    # <Title>
    # -----------------------------------
    # ==========================
    for i in range(len(block) - 4):
        if (
            block[i].strip() == LINE_EQ
            and block[i + 1].strip() == LINE_DASH
            and block[i + 3].strip() == LINE_DASH
            and block[i + 4].strip() == LINE_EQ
        ):
            title = block[i + 2].strip()
            return title if title else None
    return None


def write_blocks(input_md: Path, output_dir: Path) -> None:
    lines = input_md.read_text(encoding="utf-8").splitlines()
    blocks = split_blocks(lines)
    output_dir.mkdir(parents=True, exist_ok=True)

    files: dict[str, List[str]] = {}
    current_title: Optional[str] = None

    for block in blocks:
        title = extract_title(block)
        if title:
            current_title = title
            files.setdefault(current_title, []).extend(block)
            continue
        if current_title:
            files.setdefault(current_title, []).append("")
            files[current_title].extend(block)

    for title, content_lines in files.items():
        filename = slugify(title) + ".md"
        (output_dir / filename).write_text("\n".join(content_lines).rstrip() + "\n", encoding="utf-8")


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Cắt file Markdown theo từng tiêu đề và lưu thành từng file riêng."
    )
    parser.add_argument("input_md", type=Path, help="File md nguồn.")
    parser.add_argument("output_dir", type=Path, help="Thư mục đích.")
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> None:
    args = parse_args(argv)
    if not args.input_md.exists():
        print(f"Không tìm thấy file: {args.input_md}", file=sys.stderr)
        raise SystemExit(2)
    write_blocks(args.input_md, args.output_dir)


if __name__ == "__main__":
    main()
