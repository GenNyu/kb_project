import argparse
import csv
import re
from collections import OrderedDict
from pathlib import Path

REQ_CODE_RE = re.compile(r'\b(\d+(?:\.\d+)+(?:\s*\(continued\))?)\b', re.I)
PURE_CONTINUED_RE = re.compile(r'^\s*(\d+(?:\.\d+)+)\s*\(continued\)\s*$', re.I)

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
TEST_CODE_RE = re.compile(r'\b(\d+(?:\.\d+)+(?:\.[a-z])?)\b', re.I)
TEST_START_RE = re.compile(rf'\b(\d+(?:\.\d+)+(?:\.[a-z])?)\b\s+({TEST_VERBS_RE})\b', re.I)

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
        for m in re.finditer(re.escape(label), text, flags=re.I):
            found.append((m.start(), m.end(), label))
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
        # Tránh lặp đoạn y hệt do CSV bị nhân bản
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

    # Dòng continuation chỉ có mã + (continued)
    m_cont = PURE_CONTINUED_RE.match(text)
    if m_cont:
        return strip_continued(m_cont.group(1))

    # Dòng requirement dạng "1.2.6 ...": tạo requirement mới dù không có nhãn
    m_bare = re.match(r'^\s*(\d+(?:\.\d+)+)\s+(.+)$', text)
    if m_bare:
        code = strip_continued(m_bare.group(1))
        rest = m_bare.group(2).strip(' .:-')
        # Chỉ coi là requirement nếu có ít nhất 2 dấu chấm (vd 1.2.6)
        if code.count('.') >= 2:
            req = get_requirement(store, code)
            append_field(req, 'defined_requirements', rest)
            req['_last_c0_field'] = 'defined_requirements'
            return code
        # Nếu chỉ là tiêu đề mục (vd 1.2 ...), bỏ qua để tránh dính vào requirement trước đó
        return current_code

    # Không có nhãn: nối vào field c0 trước đó của requirement hiện tại.
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

    # Loại bỏ tiền tố đánh số "1." nếu ngay sau đó là mã test (vd: "1. 1.1 Examine ...")
    text = re.sub(r'^\s*\d+\.\s+(?=\d+\.)', '', text)

    # Nếu dòng bắt đầu bằng mã nhưng không theo verb kiểm tra, coi là không phải test
    m0 = re.match(r'^\s*(\d+(?:\.\d+)+(?:\.[a-z])?)\s+([A-Za-z]+)', text)
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


def finalize_text(value: str, default='N/A') -> str:
    value = clean_text(value)
    return value if value else default


def format_requirement(req: dict) -> str:
    lines = [
        f'* Mã Yêu cầu: "{req["code"]}"',
        f'* Defined Approach Requirements: {finalize_text(req["defined_requirements"])}',
        f'* Customized Approach Objective: {finalize_text(req["customized_objective"])}',
        f'* Applicability Notes: {finalize_text(req["applicability_notes"])}',
        '* Defined Approach Testing Procedures:',
    ]

    if req['tests']:
        for test_code, test_text in req['tests'].items():
            lines.append(f'    "{test_code}": {finalize_text(test_text)}')
    else:
        lines.append('    "N/A": N/A')

    lines.extend([
        f'* Guidance - Purpose: {finalize_text(req["guidance_purpose"])}',
        f'* Guidance - Good Practice: {finalize_text(req["guidance_good_practice"])}',
        f'* Guidance - Further Information: {finalize_text(req["guidance_further_information"])}',
        f'* Guidance - Definitions: {finalize_text(req["guidance_definitions"])}',
        f'* Guidance - Examples: {finalize_text(req["guidance_examples"])}',
    ])
    return '\n'.join(lines)


def convert_csv_to_txt(input_path: str, output_path: str):
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

    # Bỏ các metadata/header không parse được nội dung chính
    records = [req for req in store.values() if req['defined_requirements'] or req['tests']]

    with open(output_path, 'w', encoding='utf-8') as out:
        for idx, req in enumerate(records):
            if idx:
                out.write('\n\n')
            out.write(format_requirement(req))


def main():
    parser = argparse.ArgumentParser(description='Chuyển CSV đã tách cột c0/c1/c2 sang TXT theo form yêu cầu.')
    parser.add_argument('input_csv', help='Đường dẫn file CSV đầu vào')
    parser.add_argument('-o', '--output', help='Đường dẫn file TXT đầu ra')
    args = parser.parse_args()

    input_path = Path(args.input_csv)
    output_path = Path(args.output) if args.output else input_path.with_suffix('.txt')

    convert_csv_to_txt(str(input_path), str(output_path))
    print(f'Đã tạo: {output_path}')


if __name__ == '__main__':
    main()
