#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Convert a CSV file into grouped TXT blocks with the format:

* Mã Yêu cầu: "[code]"
* Defined Approach Requirements: ...
* Customized Approach Objective: ...
* Applicability Notes: ...
* Defined Approach Testing Procedures:
              "[test a]" "[test b]"
* Guidance - Purpose: ...
* Guidance - Good Practice: ...
* Guidance - Further Information: ...
* Guidance - Definitions: ...
* Guidance - Examples: ...

Rules:
- Group rows by requirement code.
- Only keep codes with 3 levels or more, including sub-items like 1.2.5.a.
- Columns are mapped as:
  - c0 -> Defined Approach Requirements, Customized Approach Objective, Applicability Notes
  - c1 -> Defined Approach Testing Procedures
  - c2 -> Guidance - Purpose, Guidance - Good Practice, Guidance - Further Information,
          Guidance - Definitions, Guidance - Examples
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import List, Dict, Any

import pandas as pd


def is_valid_requirement_code(value: Any) -> bool:
    """Return True for codes with at least 3 levels, e.g. 1.2.1 or 1.2.5.a."""
    if pd.isna(value):
        return False
    code = str(value).strip()
    if not code:
        return False
    parts = code.split(".")
    return len(parts) >= 3 and all(part.strip() for part in parts)


def clean_cell(value: Any) -> str | None:
    """Normalize a CSV cell value."""
    if pd.isna(value):
        return None
    text = str(value).strip()
    return text if text else None


def get_or_na(items: List[str], index: int) -> str:
    """Return the item at index or N/A."""
    if index < len(items) and items[index]:
        return items[index]
    return "N/A"


def convert_csv_to_txt(input_csv: str, output_txt: str) -> None:
    df = pd.read_csv(input_csv)

    if len(df.columns) < 4:
        raise ValueError(
            "CSV must have at least 4 columns: code, c0, c1, c2."
        )

    code_col = df.columns[0]
    c0_col = df.columns[1]
    c1_col = df.columns[2]
    c2_col = df.columns[3]

    grouped: List[Dict[str, Any]] = []
    current: Dict[str, Any] | None = None

    for _, row in df.iterrows():
        code = clean_cell(row[code_col])

        # Start a new block whenever a valid code appears
        if code and is_valid_requirement_code(code):
            if current is not None:
                grouped.append(current)
            current = {
                "code": code,
                "c0": [],
                "c1": [],
                "c2": [],
            }

        # Append row contents to the current block
        if current is not None:
            c0 = clean_cell(row[c0_col])
            c1 = clean_cell(row[c1_col])
            c2 = clean_cell(row[c2_col])

            if c0:
                current["c0"].append(c0)
            if c1:
                current["c1"].append(c1)
            if c2:
                current["c2"].append(c2)

    if current is not None:
        grouped.append(current)

    lines: List[str] = []

    for item in grouped:
        lines.append(f'* Mã Yêu cầu: "{item["code"]}"')
        lines.append(
            f'* Defined Approach Requirements: {get_or_na(item["c0"], 0)}'
        )
        lines.append(
            f'* Customized Approach Objective: {get_or_na(item["c0"], 1)}'
        )
        lines.append(
            f'* Applicability Notes: {get_or_na(item["c0"], 2)}'
        )

        tests = (
            " ".join(f'"{x}"' for x in item["c1"])
            if item["c1"]
            else '"N/A"'
        )
        lines.append("* Defined Approach Testing Procedures:")
        lines.append(f"              {tests}")

        lines.append(f'* Guidance - Purpose: {get_or_na(item["c2"], 0)}')
        lines.append(f'* Guidance - Good Practice: {get_or_na(item["c2"], 1)}')
        lines.append(
            f'* Guidance - Further Information: {get_or_na(item["c2"], 2)}'
        )
        lines.append(f'* Guidance - Definitions: {get_or_na(item["c2"], 3)}')
        lines.append(f'* Guidance - Examples: {get_or_na(item["c2"], 4)}')
        lines.append("")

    Path(output_txt).write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert grouped CSV requirement data into TXT format."
    )
    parser.add_argument(
        "input_csv",
        nargs="?",
        default="combined.csv",
        help="Path to input CSV file (default: combined.csv)",
    )
    parser.add_argument(
        "output_txt",
        nargs="?",
        default="combined_converted.txt",
        help="Path to output TXT file (default: combined_converted.txt)",
    )
    args = parser.parse_args()

    convert_csv_to_txt(args.input_csv, args.output_txt)
    print(f"Done. Output written to: {args.output_txt}")


if __name__ == "__main__":
    main()
