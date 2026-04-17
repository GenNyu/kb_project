from __future__ import annotations

import argparse
import re
import subprocess
from pathlib import Path

import pandas as pd
from docling.document_converter import DocumentConverter


_PROC_ID_RE = re.compile(r"\b\d+\.\d+\.\d+\.[a-z]\b")


def _extract_procedure_text(input_path: Path) -> dict[str, str]:
    """Extract testing-procedure lines from PDF text via pdftotext.

    Returns mapping id -> full text (single line).
    """
    try:
        res = subprocess.run(
            ["pdftotext", str(input_path), "-"],
            check=True,
            capture_output=True,
            text=True,
        )
    except Exception as exc:
        print(f"[warn] pdftotext failed: {exc}")
        return {}

    lines = [ln.strip() for ln in res.stdout.splitlines()]
    out: dict[str, str] = {}
    i = 0
    while i < len(lines):
        line = lines[i]
        m = _PROC_ID_RE.match(line)
        if not m:
            i += 1
            continue
        proc_id = m.group(0)
        parts = [line]
        i += 1
        while i < len(lines) and lines[i]:
            parts.append(lines[i])
            i += 1
        text = " ".join(parts)
        text = re.sub(r"\s+", " ", text).strip()
        out[proc_id] = text
        i += 1
    return out


def _append_missing_procedures(
    merged: pd.DataFrame, input_path: Path
) -> tuple[pd.DataFrame, int]:
    if merged.empty:
        return merged, 0

    existing_ids: set[str] = set()
    for col in merged.columns:
        if col == "__table_index__":
            continue
        s = merged[col].astype(str)
        for ids in s.str.findall(_PROC_ID_RE):
            existing_ids.update(ids)

    extracted = _extract_procedure_text(input_path)
    missing = [pid for pid in extracted.keys() if pid not in existing_ids]
    if not missing:
        return merged, 0

    # Append missing procedures to c1 by default (testing-procedures column)
    rows = []
    for pid in missing:
        row = {col: "" for col in merged.columns}
        row["__table_index__"] = 0
        if "c1" in merged.columns:
            row["c1"] = extracted[pid]
        else:
            # Fallback: store in last column
            row[merged.columns[-1]] = extracted[pid]
        rows.append(row)

    merged = pd.concat([merged, pd.DataFrame(rows)], ignore_index=True)
    return merged, len(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract Docling tables and merge into one table.")
    parser.add_argument("input", nargs="?", default="input/PCI_243_245.pdf")
    parser.add_argument(
        "-o",
        "--output",
        default="output/docling/combined.csv",
        help="Output file (.csv) to store the merged table.",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    converter = DocumentConverter()
    result = converter.convert(str(input_path))

    tables = result.document.tables
    if not tables:
        print("[warn] no tables found")
        return

    frames = []
    for i, table in enumerate(tables, start=1):
        try:
            df = table.export_to_dataframe(doc=result.document)
        except Exception as exc:
            print(f"[warn] table {i} export failed: {exc}")
            continue
        if df is None or df.empty:
            print(f"[warn] table {i} empty, skipping")
            continue
        df = df.copy()
        # Ensure unique, stable column labels for concat
        df.columns = [f"c{idx}" for idx in range(df.shape[1])]
        df.insert(0, "__table_index__", i)
        frames.append(df)

    if not frames:
        print("[warn] no usable tables after parsing")
        return

    merged = pd.concat(frames, ignore_index=True)
    merged, added = _append_missing_procedures(merged, input_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    merged.to_csv(output_path, index=False)
    if added:
        print(f"[warn] appended {added} missing procedure rows from text")
    print(f"[ok] wrote {output_path} (rows={len(merged)})")


if __name__ == "__main__":
    main()
