import argparse
import os
import re
import sys
from pathlib import Path
from collections import OrderedDict


DROP_SUBSTRINGS = [
    "Security e",
    "Security Standards Council",
    "Standards Council",
    "Requirements and Testing Procedures",
    "Payment Card Industry Data Security Standard",
    "All Rights Reserved.",
]

DROP_LINE_REGEX = [
    re.compile(r"^Guidance\s*$", re.IGNORECASE),
    re.compile(r"^June\s+\d{4}\s*$", re.IGNORECASE),
    re.compile(r"^Page\s+\d+\s*$", re.IGNORECASE),
    re.compile(r"^©\s*\d{4}", re.IGNORECASE),
    re.compile(r"^.*\|.*$"),  # stray table separators
    re.compile(r".*Customized Approach Objective.*\bExamine\b.*", re.IGNORECASE),
    re.compile(r"^Requirements and Testing Procedures Guidance$", re.IGNORECASE),
    re.compile(r".*data to verify that invalid logical access attempts.*", re.IGNORECASE),
    re.compile(r".*invalid logical access attempts unauthorized user's attempts.*", re.IGNORECASE),
]


SECTION_LABELS = {
    "Defined Approach Requirements": "dar",
    "Customized Approach Objective": "cao",
    "Applicability Notes": "applicability",
    "Defined Approach Testing Procedures": "test",
    "Purpose": "purpose",
    "Good Practice": "good_practice",
    "Further Information": "further_information",
    "Definitions": "definitions",
    "Examples": "examples",
}


REQ_CODE_RE = re.compile(r"^(\d+\.\d+\.\d+(?:\.\d+)?)\s+(.*)$")
TEST_CODE_RE = re.compile(
    r"^(\d+\.\d+\.\d+(?:\.\d+)?)(?:\.([a-z]))\s+(.*)$", re.IGNORECASE
)
TEST_CODE_NO_LETTER_RE = re.compile(r"^(\d+\.\d+\.\d+(?:\.\d+)?)\s+(.*)$")
TEST_VERB_RE = re.compile(
    r"^(Examine|Interview|Observe|Review|Inspect|Verify|Test|Compare|Check|Obtain|Evaluate|Reperform)\b",
    re.IGNORECASE,
)


def _clean_line(line: str) -> str:
    line = line.replace("\u00ad", "")
    lower = line.lower()
    if "multiple" in lower and "brute force" in lower and "guess a password" in lower:
        return (
            "Multiple invalid login attempts may be an indication of an unauthorized "
            "user's attempts to “brute force” or guess a password."
        )
    line = re.sub(r"^>?\s*Security\s+e\s*", "", line)
    line = re.sub(r"\(continued on next page\)", "", line, flags=re.IGNORECASE)
    line = re.sub(r"\d+\.\d+\.\d+\s*\(continued\)", "", line, flags=re.IGNORECASE)
    line = re.sub(r"\(continued\)", "", line, flags=re.IGNORECASE)
    line = re.sub(r"\s+", " ", line).strip()
    line = re.sub(r"\bInuse\b", "In use", line)
    line = re.sub(r"\bAllin-", "All in-", line)
    line = re.sub(r"\be-\s+commerce\b", "e-commerce", line, flags=re.IGNORECASE)
    return line


def _is_bullet(line: str) -> bool:
    return bool(re.match(r"^(?:[•\u2022\-\*«]|[e¢])\s+", line))


def _strip_bullet(line: str) -> str:
    return re.sub(r"^(?:[•\u2022\-\*«]|[e¢])\s+", "", line).strip()


def _should_drop(line: str) -> bool:
    if not line:
        return False
    for s in DROP_SUBSTRINGS:
        if s in line:
            return True
    for rx in DROP_LINE_REGEX:
        if rx.match(line):
            return True
    return False


def _preprocess_lines(text: str) -> list[str]:
    raw_lines = [l.rstrip() for l in text.splitlines()]
    out = []
    i = 0
    while i < len(raw_lines):
        line = raw_lines[i].strip()
        if line == "Defi":
            line = "Definitions"
        elif line == "itions":
            i += 1
            continue
        # Stitch split OCR sentence about brute-force attempts.
        if "brute force" in line.lower() and i + 1 < len(raw_lines):
            nxt = raw_lines[i + 1].strip()
            if "guess a password" in nxt.lower():
                line = (
                    "Multiple invalid login attempts may be an indication of an unauthorized "
                    "user's attempts to “brute force” or guess a password."
                )
                i += 1
        if line == "Guid" and i + 1 < len(raw_lines) and raw_lines[i + 1].strip() == "ance":
            line = "Guidance"
            i += 1
        line = _clean_line(line)
        if _should_drop(line):
            i += 1
            continue
        out.append(line)
        i += 1
    return out


def _add_line(lines: list[tuple[str, bool]], line: str, is_bullet: bool) -> None:
    if not line:
        return
    if not lines:
        lines.append((line, is_bullet))
        return
    prev_text, prev_is_bullet = lines[-1]
    if line.startswith("Multiple ") and prev_text.endswith("Multiple"):
        prev_text = prev_text[:-8].rstrip()
        lines[-1] = (prev_text, prev_is_bullet)
    if not is_bullet and not prev_is_bullet and not prev_text.endswith((".", ":", ";")):
        lines[-1] = (prev_text + " " + line, prev_is_bullet)
        return
    if prev_is_bullet and not is_bullet and not prev_text.endswith((".", ":", ";")):
        lines[-1] = (prev_text + " " + line, prev_is_bullet)
        return
    lines.append((line, is_bullet))


def _lines_to_text(lines: list[tuple[str, bool]]) -> str:
    if not lines:
        return "N/A"
    out = []
    for t, is_bullet in lines:
        if is_bullet:
            txt = t.rstrip().rstrip(".")
            out.append(f"  - {txt}")
        else:
            out.append(t)
    return "\n".join(out).strip() or "N/A"


def _ensure_req(reqs: OrderedDict, code: str) -> dict:
    if code not in reqs:
        reqs[code] = {
            "dar": [],
            "cao": [],
            "applicability": [],
            "tests": OrderedDict(),
            "purpose": [],
            "good_practice": [],
            "further_information": [],
            "definitions": [],
            "examples": [],
        }
    return reqs[code]




def parse(text: str) -> OrderedDict:
    lines = _preprocess_lines(text)
    reqs: OrderedDict[str, dict] = OrderedDict()
    current_section = None
    current_req_code = None
    current_test_code = None
    expect_req_code = False

    for raw in lines:
        line = raw
        if not line:
            continue

        # Skip top-level requirement headers like "1.2 ..."
        if re.match(r"^\d+\.\d+(?:\.\d+)?\s+.+", line) and not REQ_CODE_RE.match(line):
            current_section = None
            current_req_code = None
            current_test_code = None
            expect_req_code = False
            continue

        # Section labels
        if line in SECTION_LABELS:
            current_section = SECTION_LABELS[line]
            current_test_code = None
            expect_req_code = current_section == "dar"
            continue

        # Test lines with letter (e.g., 1.1.2.a)
        m = TEST_CODE_RE.match(line)
        if m:
            code, letter, text_part = m.group(1), m.group(2), m.group(3)
            if current_section == "test" or TEST_VERB_RE.match(text_part):
                current_req_code = code
                req = _ensure_req(reqs, code)
                test_code = f"{code}.{letter}"
                req["tests"][test_code] = text_part.strip()
                current_test_code = test_code
                continue

        # Requirement or test line without letter
        m = REQ_CODE_RE.match(line)
        if m:
            code, text_part = m.group(1), m.group(2)
            is_subreq = code.count(".") >= 3
            if is_subreq and not TEST_VERB_RE.match(text_part):
                current_req_code = code
                req = _ensure_req(reqs, code)
                _add_line(req["dar"], text_part, is_bullet=False)
                expect_req_code = False
                continue
            if expect_req_code:
                current_req_code = code
                req = _ensure_req(reqs, code)
                _add_line(req["dar"], text_part, is_bullet=False)
                expect_req_code = False
                continue
            if current_section == "test" or TEST_VERB_RE.match(text_part):
                current_req_code = code
                req = _ensure_req(reqs, code)
                test_code = code
                req["tests"][test_code] = text_part.strip()
                current_test_code = test_code
                continue
            if current_section in (None, "dar"):
                current_req_code = code
                req = _ensure_req(reqs, code)
                _add_line(req["dar"], text_part, is_bullet=False)
                continue

        # Continuation of a test line
        if current_section == "test" and current_req_code and current_test_code:
            req = _ensure_req(reqs, current_req_code)
            req["tests"][current_test_code] = (req["tests"][current_test_code] + " " + line).strip()
            continue

        # Content lines for other sections
        if current_req_code:
            req = _ensure_req(reqs, current_req_code)
            is_bullet = _is_bullet(line)
            text_part = _strip_bullet(line) if is_bullet else line

            if current_section in req:
                _add_line(req[current_section], text_part, is_bullet=is_bullet)
                continue

            # Fallback: if in DAR but line doesn't start with code (bullets)
            if current_section == "dar":
                _add_line(req["dar"], text_part, is_bullet=is_bullet)
                continue

    return reqs


def format_output_md(reqs: OrderedDict) -> str:
    blocks = []
    for code, data in reqs.items():
        lines = []
        lines.append(f'## Mã Yêu cầu: "{code}"')
        lines.append(f'**Defined Approach Requirements:** {_lines_to_text(data["dar"])}')
        lines.append(f'**Customized Approach Objective:** {_lines_to_text(data["cao"])}')
        lines.append(f'**Applicability Notes:** {_lines_to_text(data["applicability"])}')
        lines.append("**Defined Approach Testing Procedures:**")
        if data["tests"]:
            for t_code, t_text in data["tests"].items():
                t_text = t_text.strip() or "N/A"
                lines.append(f'- "{t_code}": "{t_text}"')
        else:
            lines.append('- "N/A": "N/A"')
        lines.append(f'**Guidance - Purpose:** {_lines_to_text(data["purpose"])}')
        lines.append(f'**Guidance - Good Practice:** {_lines_to_text(data["good_practice"])}')
        lines.append(f'**Guidance - Further Information:** {_lines_to_text(data["further_information"])}')
        lines.append(f'**Guidance - Definitions:** {_lines_to_text(data["definitions"])}')
        lines.append(f'**Guidance - Examples:** {_lines_to_text(data["examples"])}')
        blocks.append("\n".join(lines))
    return "\n\n---\n\n".join(blocks)


def _derive_output_path(input_path: Path, output_path: Path | None) -> Path:
    if output_path and output_path.is_dir():
        out_dir = output_path
    elif output_path and not output_path.suffix:
        out_dir = output_path
    elif output_path:
        return output_path
    else:
        out_dir = Path("output")
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir / f"{input_path.stem}_structured_rule.md"


def _process_file(input_path: Path, output_path: Path | None) -> None:
    with input_path.open("r", encoding="utf-8") as f:
        text = f.read()
    reqs = parse(text)
    out_text = format_output_md(reqs)
    out_file = _derive_output_path(input_path, output_path)
    out_file.parent.mkdir(parents=True, exist_ok=True)
    out_file.write_text(out_text, encoding="utf-8")
    print(f"Wrote MD output to {out_file}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert PCI TXT to structured MD.")
    parser.add_argument(
        "input",
        nargs="?",
        default="PCI_44_54.txt",
        help="Input TXT file or folder (default: PCI_44_54.txt)",
    )
    parser.add_argument(
        "output",
        nargs="?",
        default=None,
        help="Output MD file or output folder (default: output/)",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else None

    if input_path.is_dir():
        txt_files = sorted(p for p in input_path.iterdir() if p.is_file() and p.suffix.lower() == ".txt")
        if not txt_files:
            raise SystemExit(f"No .txt files found in folder: {input_path}")
        for fpath in txt_files:
            _process_file(fpath, output_path or input_path / "output")
        return

    if not input_path.exists():
        raise SystemExit(f"Input not found: {input_path}")

    _process_file(input_path, output_path)


if __name__ == "__main__":
    main()
