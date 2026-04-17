#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Iterable, Iterator, Optional

try:
    from docx import Document
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError as exc:  # pragma: no cover - runtime dependency check
    print("Thiếu thư viện python-docx. Cài bằng: pip install -r requirements.txt", file=sys.stderr)
    raise SystemExit(2) from exc


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Chuyển DOCX sang Markdown (.md).")
    parser.add_argument(
        "input_docx",
        type=Path,
        help="File DOCX hoặc thư mục chứa DOCX.",
    )
    parser.add_argument(
        "output_md",
        type=Path,
        help="File .md xuất ra hoặc thư mục đích khi xử lý hàng loạt.",
    )
    return parser.parse_args(argv)


def iter_block_items(doc: Document) -> Iterator[Paragraph | Table]:
    for child in doc.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def clean_text(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def render_run_text(run_text: str, bold: bool, italic: bool) -> str:
    if not run_text:
        return ""
    text = run_text.replace("\n", " ")
    if bold and italic:
        return f"***{text}***"
    if bold:
        return f"**{text}**"
    if italic:
        return f"*{text}*"
    return text


def paragraph_to_markdown(paragraph: Paragraph) -> str:
    style_name = (paragraph.style.name or "").strip()
    raw_text = paragraph.text or ""
    if not raw_text.strip():
        return ""

    heading_match = re.match(r"Heading (\d+)", style_name, re.IGNORECASE)
    if heading_match:
        level = min(max(int(heading_match.group(1)), 1), 6)
        return f"{'#' * level} {clean_text(raw_text)}"
    if style_name.lower() == "title":
        return f"# {clean_text(raw_text)}"

    num_pr = paragraph._p.pPr.numPr if paragraph._p.pPr is not None else None
    is_bullet = "bullet" in style_name.lower()
    is_number = "number" in style_name.lower()
    is_list = num_pr is not None or is_bullet or is_number

    if is_list:
        ilvl_val = 0
        if num_pr is not None and num_pr.ilvl is not None and num_pr.ilvl.val is not None:
            ilvl_val = int(num_pr.ilvl.val)
        indent = "  " * ilvl_val
        marker = "- " if is_bullet and not is_number else "1. "
        parts: list[str] = []
        for run in paragraph.runs:
            parts.append(render_run_text(run.text, bool(run.bold), bool(run.italic)))
        content = clean_text("".join(parts))
        return f"{indent}{marker}{content}"

    parts = []
    for run in paragraph.runs:
        parts.append(render_run_text(run.text, bool(run.bold), bool(run.italic)))
    return clean_text("".join(parts))


def table_to_markdown(table: Table) -> list[str]:
    rows = table.rows
    if not rows:
        return []
    matrix: list[list[str]] = []
    for row in rows:
        cells = [clean_text(cell.text) for cell in row.cells]
        matrix.append(cells)

    max_cols = max((len(row) for row in matrix), default=0)
    if max_cols == 0:
        return []

    normalized = [row + [""] * (max_cols - len(row)) for row in matrix]
    header = normalized[0]
    lines = []
    lines.append("| " + " | ".join(header) + " |")
    lines.append("| " + " | ".join(["---"] * max_cols) + " |")
    for row in normalized[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return lines


def docx_to_markdown(input_docx: Path) -> list[str]:
    doc = Document(input_docx)
    lines: list[str] = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            rendered = paragraph_to_markdown(block)
            if rendered:
                if lines and lines[-1] != "":
                    lines.append("")
                lines.append(rendered)
        else:
            table_lines = table_to_markdown(block)
            if table_lines:
                if lines and lines[-1] != "":
                    lines.append("")
                lines.extend(table_lines)
                lines.append("")
    while lines and lines[-1] == "":
        lines.pop()
    return lines


def write_markdown(lines: Iterable[str], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    content = "\n".join(lines).rstrip() + "\n"
    output_path.write_text(content, encoding="utf-8")


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    if args.input_docx.is_dir():
        if args.output_md.exists() and not args.output_md.is_dir():
            print("Output phải là thư mục khi input là thư mục", file=sys.stderr)
            return 2
        inputs = sorted(p for p in args.input_docx.iterdir() if p.suffix.lower() == ".docx")
        if not inputs:
            print("Không tìm thấy file .docx trong thư mục", file=sys.stderr)
            return 2
        args.output_md.mkdir(parents=True, exist_ok=True)
        for input_path in inputs:
            out_path = args.output_md / f"{input_path.stem}.md"
            lines = docx_to_markdown(input_path)
            write_markdown(lines, out_path)
    else:
        if args.output_md.is_dir():
            args.output_md.mkdir(parents=True, exist_ok=True)
            out_path = args.output_md / f"{args.input_docx.stem}.md"
        else:
            out_path = args.output_md
        lines = docx_to_markdown(args.input_docx)
        write_markdown(lines, out_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
