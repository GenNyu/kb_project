#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    import openpyxl
except ImportError as exc:  # pragma: no cover - runtime dependency check
    print("Thiếu thư viện openpyxl. Cài bằng: pip install openpyxl", file=sys.stderr)
    raise SystemExit(2) from exc

LINE_EQ = "=========================="
LINE_DASH = "-----------------------------------"


def normalize_header(value: Any) -> str:
    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    text = re.sub(r"\s+", " ", text)
    return text


def clean_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def find_header_row(rows: List[Tuple[Any, ...]]) -> Tuple[int, Dict[str, int]]:
    key_map = {
        "ten tai lieu": "topic",
        "tên tài liệu": "topic",
        "topic": "topic",
        "question": "question",
        "answer": "answer",
        "comment": "comment",
        "evidence": "evidence",
    }
    for idx, row in enumerate(rows):
        mapping: Dict[str, int] = {}
        for col_idx, cell in enumerate(row):
            if cell is None:
                continue
            key = normalize_header(cell)
            mapped = key_map.get(key)
            if mapped:
                mapping[mapped] = col_idx
        if mapping:
            return idx, mapping
    raise ValueError("Không tìm thấy dòng tiêu đề phù hợp trong file Excel.")


def build_entry(
    topic: str,
    question: str,
    answer: str,
    comment: str,
    evidence: str,
    include_header: bool,
) -> str:
    lines: List[str] = []
    if include_header:
        title = topic.strip()
        if not title:
            title = "General Question"
        lines.extend([LINE_EQ, LINE_DASH, title, LINE_DASH, LINE_EQ])
    section_lines: List[str] = []
    if question.strip():
        section_lines.extend(["###Question:", question, ""])
    if answer.strip():
        section_lines.extend(["###Answer:", answer, ""])
    if comment.strip():
        section_lines.extend(["###Comment:", comment, ""])
    if evidence.strip():
        section_lines.extend(["###Evidence:", evidence, ""])
    if not section_lines:
        return ""
    lines.extend(section_lines)
    lines.append(LINE_EQ)
    return "\n".join(lines)


def convert_sheet_to_md(path: Path, sheet_name: Optional[str] = None) -> str:
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = wb[sheet_name] if sheet_name else wb.active

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return ""

    header_row_idx, mapping = find_header_row(rows)

    required = {"topic", "question", "answer"}
    missing = required - set(mapping.keys())
    if missing:
        raise ValueError(f"Thiếu cột bắt buộc: {', '.join(sorted(missing))}")

    entries: List[str] = []
    last_topic = ""
    for row in rows[header_row_idx + 1 :]:
        if all(cell is None or str(cell).strip() == "" for cell in row):
            continue
        topic = clean_cell(row[mapping["topic"]]) if "topic" in mapping else ""
        question = clean_cell(row[mapping["question"]]) if "question" in mapping else ""
        answer = clean_cell(row[mapping["answer"]]) if "answer" in mapping else ""
        comment = clean_cell(row[mapping["comment"]]) if "comment" in mapping else ""
        evidence = clean_cell(row[mapping["evidence"]]) if "evidence" in mapping else ""
        include_header = bool(topic) and topic != last_topic
        entry = build_entry(topic, question, answer, comment, evidence, include_header)
        if entry:
            entries.append(entry)
        if topic:
            last_topic = topic

    return "\n".join(entries)


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Chuyển Excel (xlsx) sang Markdown theo mẫu Question/Answer/Comment/Evidence."
    )
    parser.add_argument("input_xlsx", type=Path, help="File xlsx hoặc thư mục chứa các file xlsx.")
    parser.add_argument("output_md", type=Path, help="File md đích hoặc thư mục đích.")
    parser.add_argument("--sheet", dest="sheet_name", help="Tên sheet cần đọc (mặc định sheet active).")
    return parser.parse_args(argv)


def ensure_output_path(input_path: Path, output_path: Path) -> Path:
    if output_path.is_dir() or str(output_path).endswith("/"):
        output_path.mkdir(parents=True, exist_ok=True)
        return output_path / f"{input_path.stem}.md"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    return output_path


def convert_file(input_path: Path, output_path: Path, sheet_name: Optional[str]) -> None:
    md_text = convert_sheet_to_md(input_path, sheet_name)
    output_path.write_text(md_text, encoding="utf-8")


def main(argv: Optional[List[str]] = None) -> None:
    args = parse_args(argv)
    input_path: Path = args.input_xlsx
    output_path: Path = args.output_md
    sheet_name: Optional[str] = args.sheet_name

    if not input_path.exists():
        print(f"Không tìm thấy file hoặc thư mục: {input_path}", file=sys.stderr)
        raise SystemExit(2)

    if input_path.is_dir():
        xlsx_files = sorted(input_path.glob("*.xlsx"))
        if not xlsx_files:
            print("Không tìm thấy file .xlsx trong thư mục.", file=sys.stderr)
            raise SystemExit(2)
        if not output_path.exists():
            output_path.mkdir(parents=True, exist_ok=True)
        for xlsx_file in xlsx_files:
            out_file = ensure_output_path(xlsx_file, output_path)
            convert_file(xlsx_file, out_file, sheet_name)
    else:
        out_file = ensure_output_path(input_path, output_path)
        convert_file(input_path, out_file, sheet_name)


if __name__ == "__main__":
    main()
