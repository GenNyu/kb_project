
import csv
import re
import sys
from collections import OrderedDict

SECTION_LABELS = [
    "Customized Approach Objective",
    "Applicability Notes",
    "Purpose",
    "Good Practice",
    "Further Information",
    "Definitions",
    "Examples",
]

def clean_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\r", " ").replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def extract_requirement_code(text: str):
    m = re.match(r"^(\d+\.\d+\.\d+)\b", clean_text(text))
    return m.group(1) if m else None

def extract_test_code(text: str):
    m = re.match(r"^(\d+\.\d+\.\d+\.[a-z])\b", clean_text(text), flags=re.I)
    return m.group(1) if m else None

def split_guidance(text: str):
    text = clean_text(text)
    result = {
        "Purpose": "N/A",
        "Good Practice": "N/A",
        "Further Information": "N/A",
        "Definitions": "N/A",
        "Examples": "N/A",
    }
    if not text:
        return result

    labels = ["Purpose", "Good Practice", "Further Information", "Definitions", "Examples"]
    matches = list(re.finditer(r"(Purpose|Good Practice|Further Information|Definitions|Examples)\b", text))
    if not matches:
        return result

    for i, m in enumerate(matches):
        label = m.group(1)
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        content = clean_text(text[start:end])
        result[label] = content if content else "N/A"
    return result

def split_labeled_sections(text: str):
    """
    For cells that may contain:
    Customized Approach Objective ...
    Applicability Notes ...
    Purpose ...
    Good Practice ...
    ...
    """
    text = clean_text(text)
    out = {label: "" for label in SECTION_LABELS}
    if not text:
        return out

    pattern = r"(Customized Approach Objective|Applicability Notes|Purpose|Good Practice|Further Information|Definitions|Examples)\b"
    matches = list(re.finditer(pattern, text))
    if not matches:
        return out

    for i, m in enumerate(matches):
        label = m.group(1)
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        out[label] = clean_text(text[start:end])
    return out

def parse_test_procedure(text: str):
    text = clean_text(text)
    code = extract_test_code(text)
    if not code:
        return None

    body = text[len(code):].strip(" .")
    # Keep original text as fallback
    original_body = body

    subject = "Assessor"

    # Typical PCI phrasing:
    # Examine configuration settings for NSCs to verify that ...
    m = re.match(r"^(Examine|Observe|Inspect|Interview)\s+(.+?)\s+to\s+(verify\b.*)$", body, flags=re.I)
    if m:
        action = m.group(1)
        obj = clean_text(m.group(2))
        verify = clean_text(m.group(3))
        verify = verify[:1].upper() + verify[1:] if verify else "Verify."
        return f'- "{code}": [Chủ thể: {subject}] thực hiện [Hành động: {action}] [Đối tượng: {obj}] để [Hành động chốt chặn: {verify}].'

    return f'- "{code}": {original_body}.'

def main():
    if len(sys.argv) < 3:
        print("Usage: python reformat_combined_v2.py input.csv output.txt")
        sys.exit(1)

    input_csv = sys.argv[1]
    output_txt = sys.argv[2]

    groups = OrderedDict()

    with open(input_csv, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            cells = [clean_text(row.get("c0", "")), clean_text(row.get("c1", "")), clean_text(row.get("c2", ""))]
            req_code = None
            for cell in cells:
                req_code = extract_requirement_code(cell)
                if req_code:
                    break
            if not req_code:
                continue

            item = groups.setdefault(req_code, {
                "code": req_code,
                "requirement": "N/A",
                "tests": [],
                "customized_objective": "N/A",
                "applicability_notes": "N/A",
                "guidance": {
                    "Purpose": "N/A",
                    "Good Practice": "N/A",
                    "Further Information": "N/A",
                    "Definitions": "N/A",
                    "Examples": "N/A",
                }
            })

            # c0 normally holds requirement text
            c0 = cells[0]
            if extract_requirement_code(c0):
                req_text = clean_text(re.sub(rf"^{re.escape(req_code)}\s*", "", c0))
                if req_text:
                    item["requirement"] = req_text

            # c1 normally holds one test procedure
            c1 = cells[1]
            test_line = parse_test_procedure(c1)
            if test_line and test_line not in item["tests"]:
                item["tests"].append(test_line)

            # c2 may hold objective/applicability/guidance, or guidance only
            c2 = cells[2]
            parts = split_labeled_sections(c2)
            if parts.get("Customized Approach Objective"):
                item["customized_objective"] = parts["Customized Approach Objective"]
            if parts.get("Applicability Notes"):
                item["applicability_notes"] = parts["Applicability Notes"]

            guidance_source = c2
            if any(parts.get(k) for k in ["Purpose", "Good Practice", "Further Information", "Definitions", "Examples"]):
                guidance = {
                    "Purpose": parts.get("Purpose") or "N/A",
                    "Good Practice": parts.get("Good Practice") or "N/A",
                    "Further Information": parts.get("Further Information") or "N/A",
                    "Definitions": parts.get("Definitions") or "N/A",
                    "Examples": parts.get("Examples") or "N/A",
                }
            else:
                guidance = split_guidance(guidance_source)

            for k, v in guidance.items():
                if v and v != "N/A":
                    item["guidance"][k] = v

    with open(output_txt, "w", encoding="utf-8") as out:
        first = True
        for code, item in groups.items():
            if not first:
                out.write("---\n")
            first = False

            out.write(f'Mã Yêu cầu: "{item["code"]}"\n')
            out.write(f'Defined Approach Requirements: {item["requirement"]}\n')
            out.write('Defined Approach Testing Procedures:\n')
            if item["tests"]:
                for t in item["tests"]:
                    out.write(f"{t}\n")
            else:
                out.write("- N/A\n")
            out.write(f'Customized Approach Objective: {item["customized_objective"]}\n')
            out.write(f'Applicability Notes: {item["applicability_notes"]}\n')
            out.write(f'Guidance - Purpose: {item["guidance"]["Purpose"]}\n')
            out.write(f'Guidance - Good Practice: {item["guidance"]["Good Practice"]}\n')
            out.write(f'Guidance - Further Information: {item["guidance"]["Further Information"]}\n')
            out.write(f'Guidance - Definitions: {item["guidance"]["Definitions"]}\n')
            out.write(f'Guidance - Examples: {item["guidance"]["Examples"]}\n')

if __name__ == "__main__":
    main()
