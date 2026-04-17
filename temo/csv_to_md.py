import argparse
import csv
import re
from collections import OrderedDict
from pathlib import Path
from typing import Optional

REQ_CODE_RE = re.compile(r'\b([A-Za-z]?\d+(?:\.\d+)+(?:\s*\(continued\))?)\b', re.I)
PURE_CONTINUED_RE = re.compile(r'^\s*([A-Za-z]?\d+(?:\.\d+)+)\s*\(continued\)\s*$', re.I)

C0_LABELS = [
    'Defined Approach Requirements',
    'Customized Approach Objective',
    'Applicability Notes',
]
C2_LABELS = [
    'Purpose',
    'Good Practice',
    'Further Information',
    'Definitions',
    'Examples',
    'Example',
]

TEST_VERBS = [
    'Examine',
    'Interview',
    'Observe',
    'Review',
    'Verify',
    'Inspect',
    'Confirm',
    'Obtain',
    'Test',
    'Evaluate',
    'Determine',
    'Check',
    'Compare',
]
TEST_VERBS_RE = '|'.join(re.escape(v) for v in TEST_VERBS)
TEST_CODE_PATTERN = r'(?:[A-Za-z](?:\.)?)?\d+(?:\.\d+)+(?:\.[a-z])?'
TEST_CODE_RE = re.compile(rf'\b({TEST_CODE_PATTERN})\b', re.I)
TEST_START_RE = re.compile(rf'\b({TEST_CODE_PATTERN})\b\s+({TEST_VERBS_RE})\b', re.I)


def clean_text(text: str) -> str:
    if text is None:
        return ''
    text = str(text).replace('\ufeff', ' ').strip()
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def strip_continued(code: str) -> str:
    return re.sub(r'\s*\(continued\)\s*', '', code or '', flags=re.I).strip()


def find_label_positions(text: str, labels):
    found = []
    for label in labels:
        pattern = rf'(^|[.!?]\s+)({re.escape(label)})\b'
        for m in re.finditer(pattern, text, flags=re.I | re.M):
            found.append((m.start(2), m.end(2), label))
            break
    return sorted(found, key=lambda x: x[0])


def split_by_labels(text: str, labels):
    positions = find_label_positions(text, labels)
    if not positions:
        return []
    parts = []
    for i, (start, end, label) in enumerate(positions):
        next_start = positions[i + 1][0] if i + 1 < len(positions) else len(text)
        value = text[end:next_start].strip(' :.-')
        parts.append((label, value.strip()))
    return parts


def get_requirement(store: OrderedDict, code: str):
    code = strip_continued(code)
    if code not in store:
        store[code] = {
            'code': code,
            'defined_requirements': '',
            'customized_objective': '',
            'applicability_notes': '',
            'tests': OrderedDict(),
            'guidance_purpose': '',
            'guidance_good_practice': '',
            'guidance_further_information': '',
            'guidance_definitions': '',
            'guidance_examples': '',
            '_last_c0_field': None,
            '_last_guidance_field': None,
            '_last_test_code': None,
        }
    return store[code]


def append_field(req: dict, field: str, text: str):
    text = clean_text(text)
    if not text:
        return
    existing = req.get(field, '')
    if existing:
        if existing == text or existing.endswith(text):
            return
        if text in existing:
            return
        req[field] = f"{existing} {text}".strip()
    else:
        req[field] = text


def append_test(req: dict, test_code: str, text: str):
    test_code = strip_continued(test_code)
    text = clean_text(text)
    if not test_code or not text:
        return
    if test_code in req['tests']:
        existing = req['tests'][test_code]
        if existing == text or existing.endswith(text):
            return
        if text in existing:
            return
        req['tests'][test_code] = f"{existing} {text}".strip()
    else:
        req['tests'][test_code] = text
    req['_last_test_code'] = test_code


def parse_c0(text: str, store: OrderedDict, current_code: str):
    text = clean_text(text)
    if not text:
        return current_code

    parts = split_by_labels(text, C0_LABELS)
    if parts:
        active_req = None
        active_code = current_code
        for label, value in parts:
            if label == 'Defined Approach Requirements':
                m = REQ_CODE_RE.search(value)
                if m:
                    active_code = strip_continued(m.group(1))
                    value = value[m.end():].strip(' .:-')
                if not active_code:
                    continue
                active_req = get_requirement(store, active_code)
                append_field(active_req, 'defined_requirements', value)
                active_req['_last_c0_field'] = 'defined_requirements'
            elif label == 'Customized Approach Objective':
                if not active_code:
                    continue
                active_req = get_requirement(store, active_code)
                append_field(active_req, 'customized_objective', value)
                active_req['_last_c0_field'] = 'customized_objective'
            elif label == 'Applicability Notes':
                if not active_code:
                    continue
                active_req = get_requirement(store, active_code)
                append_field(active_req, 'applicability_notes', value)
                active_req['_last_c0_field'] = 'applicability_notes'
        return active_code

    m_cont = PURE_CONTINUED_RE.match(text)
    if m_cont:
        return strip_continued(m_cont.group(1))

    m_bare = re.match(r'^\s*([A-Za-z]?\d+(?:\.\d+)+)\s+(.+)$', text)
    if m_bare:
        code = strip_continued(m_bare.group(1))
        rest = m_bare.group(2).strip(' .:-')
        if code.count('.') >= 2:
            req = get_requirement(store, code)
            append_field(req, 'defined_requirements', rest)
            req['_last_c0_field'] = 'defined_requirements'
            return code
        return current_code

    if current_code and current_code in store:
        req = store[current_code]
        if req.get('_last_c0_field'):
            append_field(req, req['_last_c0_field'], text)
    return current_code


def parse_c1(text: str, store: OrderedDict, current_code: str):
    text = clean_text(text)
    if not text or not current_code:
        return current_code

    req = get_requirement(store, current_code)
    text = re.sub(r'^Defined Approach Testing Procedures\s*', '', text, flags=re.I).strip()
    if not text:
        return current_code

    text = re.sub(r'^\s*\d+\.\s+(?=\d+\.)', '', text)

    m0 = re.match(rf'^\s*({TEST_CODE_PATTERN})\s+([A-Za-z]+)', text)
    if m0 and m0.group(2).capitalize() not in TEST_VERBS:
        return current_code

    matches = list(TEST_START_RE.finditer(text))
    if matches:
        for i, m in enumerate(matches):
            next_start = matches[i + 1].start() if i + 1 < len(matches) else len(text)
            code = m.group(1)
            desc = text[m.end(1):next_start].strip()
            append_test(req, code, desc)
    else:
        last_code = req.get('_last_test_code')
        if last_code:
            append_test(req, last_code, text)
    return current_code


def parse_c2(text: str, store: OrderedDict, current_code: str):
    text = clean_text(text)
    if not text or not current_code:
        return current_code

    req = get_requirement(store, current_code)
    parts = split_by_labels(text, C2_LABELS)
    if parts:
        for label, value in parts:
            normalized = label.lower()
            if normalized == 'purpose':
                field = 'guidance_purpose'
            elif normalized == 'good practice':
                field = 'guidance_good_practice'
            elif normalized == 'further information':
                field = 'guidance_further_information'
            elif normalized == 'definitions':
                field = 'guidance_definitions'
            elif normalized in ('example', 'examples'):
                field = 'guidance_examples'
            else:
                continue
            append_field(req, field, value)
            req['_last_guidance_field'] = field
    else:
        last_field = req.get('_last_guidance_field')
        if last_field:
            append_field(req, last_field, text)
    return current_code


def finalize_text(value: str) -> str:
    return clean_text(value)


def format_requirement_md(req: dict) -> str:
    lines = [f'## Mã Yêu cầu: `{req["code"]}`']

    defined = finalize_text(req["defined_requirements"])
    customized = finalize_text(req["customized_objective"])
    applicability = finalize_text(req["applicability_notes"])

    if defined:
        lines.append(f'**Defined Approach Requirements:** {defined}')
    if customized:
        lines.append(f'**Customized Approach Objective:** {customized}')
    if applicability:
        lines.append(f'**Applicability Notes:** {applicability}')

    if req['tests']:
        lines.append('**Defined Approach Testing Procedures:**')
        for test_code, test_text in req['tests'].items():
            test_text = finalize_text(test_text)
            if test_text:
                lines.append(f'- `{test_code}`: {test_text}')
            else:
                lines.append(f'- `{test_code}`')
    else:
        # Không ghi phần Testing Procedures nếu không có dữ liệu
        pass

    guidance_fields = [
        ('Purpose', req["guidance_purpose"]),
        ('Good Practice', req["guidance_good_practice"]),
        ('Further Information', req["guidance_further_information"]),
        ('Definitions', req["guidance_definitions"]),
        ('Examples', req["guidance_examples"]),
    ]
    for label, value in guidance_fields:
        value = finalize_text(value)
        if value:
            lines.append(f'**Guidance - {label}:** {value}')
    return '\n'.join(lines)


def convert_csv_to_md(input_path: str, output_path: str):
    store = OrderedDict()
    current_code = None

    with open(input_path, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            c0 = row.get('c0', '')
            c1 = row.get('c1', '')
            c2 = row.get('c2', '')

            current_code = parse_c0(c0, store, current_code)
            current_code = parse_c1(c1, store, current_code)
            current_code = parse_c2(c2, store, current_code)

    records = [req for req in store.values() if req['defined_requirements'] or req['tests']]

    with open(output_path, 'w', encoding='utf-8') as out:
        for idx, req in enumerate(records):
            if idx:
                out.write('\n\n---\n')
            out.write(format_requirement_md(req))


def escape_md_table(text: str) -> str:
    return clean_text(text).replace('|', '\\|')


def convert_appendix_g_to_md(input_path: str, output_path: str):
    rows = []
    with open(input_path, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            term = clean_text(row.get('c0', ''))
            definition = clean_text(row.get('c1', ''))
            if not term and not definition:
                continue
            rows.append((term, definition))

    lines = [
        '# Appendix G - PCI DSS Glossary of Terms, Abbreviations, and Acronyms',
        '',
        '| Term | Definition |',
        '| --- | --- |',
    ]
    for term, definition in rows:
        term = escape_md_table(term)
        definition = escape_md_table(definition)
        if term:
            term = f'**{term}**'
        lines.append(f'| {term} | {definition} |')

    with open(output_path, 'w', encoding='utf-8') as out:
        out.write('\n'.join(lines))


def process_folder(input_dir: Path, output_path: Optional[Path]):
    if output_path is None:
        output_dir = input_dir
    else:
        output_dir = output_path
        if output_dir.suffix:
            raise ValueError('Khi input là folder, -o phải là folder (không phải file).')
        output_dir.mkdir(parents=True, exist_ok=True)

    csv_files = sorted(p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() == '.csv')
    if not csv_files:
        print(f'Không tìm thấy file CSV trong: {input_dir}')
        return

    for csv_file in csv_files:
        out_file = output_dir / f'{csv_file.stem}.md'
        convert_csv_to_md(str(csv_file), str(out_file))
        print(f'Đã tạo: {out_file}')


def main():
    parser = argparse.ArgumentParser(description='Chuyển CSV đã tách cột c0/c1/c2 sang Markdown theo form yêu cầu.')
    parser.add_argument('input_csv', help='Đường dẫn file CSV đầu vào')
    parser.add_argument('-o', '--output', help='Đường dẫn file Markdown đầu ra')
    parser.add_argument('--appendixg', action='store_true', help='Xuất Markdown theo format Appendix G (Term/Definition).')
    args = parser.parse_args()

    input_path = Path(args.input_csv)
    output_path = Path(args.output) if args.output else None

    if input_path.is_dir():
        if args.appendixg:
            if output_path is not None and output_path.suffix:
                raise ValueError('Khi input là folder, -o phải là folder (không phải file).')
            output_dir = output_path or input_path
            output_dir.mkdir(parents=True, exist_ok=True)
            csv_files = sorted(p for p in input_path.iterdir() if p.is_file() and p.suffix.lower() == '.csv')
            if not csv_files:
                print(f'Không tìm thấy file CSV trong: {input_path}')
                return
            for csv_file in csv_files:
                out_file = output_dir / f'{csv_file.stem}.md'
                convert_appendix_g_to_md(str(csv_file), str(out_file))
                print(f'Đã tạo: {out_file}')
            return
        process_folder(input_path, output_path)
        return

    if output_path is None:
        output_path = input_path.with_suffix('.md')

    if args.appendixg:
        convert_appendix_g_to_md(str(input_path), str(output_path))
    else:
        convert_csv_to_md(str(input_path), str(output_path))
    print(f'Đã tạo: {output_path}')


if __name__ == '__main__':
    main()
