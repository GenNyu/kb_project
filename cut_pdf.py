#!/usr/bin/env python3
"""Cut a PDF by selecting pages.

- Source PDF is read from ./input
- Output PDF is written to ./output
- Pages are provided by user input, 1-based (e.g. 1,3,5-7)
"""

from __future__ import annotations

from pathlib import Path


def _load_pdf_module():
    try:
        from pypdf import PdfReader, PdfWriter  # type: ignore
        return PdfReader, PdfWriter
    except Exception:
        from PyPDF2 import PdfReader, PdfWriter  # type: ignore
        return PdfReader, PdfWriter


def parse_pages(spec: str, max_pages: int) -> list[int]:
    spec = spec.strip()
    if not spec:
        raise ValueError("Empty page specification")

    pages: list[int] = []
    parts = [p.strip() for p in spec.split(",") if p.strip()]
    for part in parts:
        if "-" in part:
            start_s, end_s = [x.strip() for x in part.split("-", 1)]
            if not start_s or not end_s:
                raise ValueError(f"Invalid range: '{part}'")
            start = int(start_s)
            end = int(end_s)
            if start <= 0 or end <= 0:
                raise ValueError("Page numbers must be >= 1")
            if start > end:
                raise ValueError(f"Range start > end: '{part}'")
            pages.extend(range(start, end + 1))
        else:
            n = int(part)
            if n <= 0:
                raise ValueError("Page numbers must be >= 1")
            pages.append(n)

    # De-duplicate while preserving order
    seen = set()
    ordered = []
    for p in pages:
        if p not in seen:
            seen.add(p)
            ordered.append(p)

    # Validate against max pages
    for p in ordered:
        if p > max_pages:
            raise ValueError(f"Page {p} exceeds total pages {max_pages}")

    return ordered


def list_pdfs(folder: Path) -> list[Path]:
    if not folder.exists():
        return []
    return sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"])


def main() -> None:
    input_dir = Path("input")
    output_dir = Path("output")
    output_dir.mkdir(parents=True, exist_ok=True)

    pdfs = list_pdfs(input_dir)
    if not pdfs:
        print("No PDF files found in ./input")
        return

    print("PDFs in ./input:")
    for i, p in enumerate(pdfs, start=1):
        print(f"  {i}. {p.name}")

    choice = input("Select file by number (or type filename): ").strip()
    if choice.isdigit():
        idx = int(choice)
        if idx < 1 or idx > len(pdfs):
            raise SystemExit("Invalid selection")
        in_path = pdfs[idx - 1]
    else:
        in_path = input_dir / choice
        if not in_path.exists():
            raise SystemExit(f"File not found: {in_path}")

    PdfReader, PdfWriter = _load_pdf_module()
    reader = PdfReader(str(in_path))
    total_pages = len(reader.pages)
    print(f"Total pages: {total_pages}")

    spec = input("Pages to keep (e.g. 1,3,5-7): ").strip()
    pages = parse_pages(spec, total_pages)

    out_name = input("Output filename (e.g. result.pdf): ").strip()
    if not out_name.lower().endswith(".pdf"):
        out_name += ".pdf"
    out_path = output_dir / out_name

    writer = PdfWriter()
    for p in pages:
        writer.add_page(reader.pages[p - 1])

    with out_path.open("wb") as f:
        writer.write(f)

    print(f"Saved: {out_path}")


if __name__ == "__main__":
    main()
