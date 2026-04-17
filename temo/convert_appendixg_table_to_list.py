import argparse
import re
from pathlib import Path


def clean_cell(text: str) -> str:
    text = text.strip()
    if text.startswith('**') and text.endswith('**') and len(text) >= 4:
        text = text[2:-2].strip()
    return text


def parse_md_table(lines):
    rows = []
    in_table = False
    for line in lines:
        if not in_table:
            if line.strip().startswith('|') and '|' in line:
                in_table = True
            else:
                continue
        if not line.strip().startswith('|'):
            if in_table:
                break
            continue
        parts = [p.strip() for p in line.strip().strip('|').split('|')]
        if len(parts) < 2:
            continue
        if parts[0].lower() == 'term' and parts[1].lower() == 'definition':
            continue
        if re.fullmatch(r'-{3,}', parts[0]) and re.fullmatch(r'-{3,}', parts[1]):
            continue
        rows.append((clean_cell(parts[0]), clean_cell(parts[1])))
    return rows


def main():
    parser = argparse.ArgumentParser(description='Chuyển bảng Appendix G sang dạng danh sách bullet.')
    parser.add_argument('input_md', help='Đường dẫn file Markdown đầu vào (bảng)')
    parser.add_argument('-o', '--output', help='Đường dẫn file Markdown đầu ra')
    args = parser.parse_args()

    input_path = Path(args.input_md)
    output_path = Path(args.output) if args.output else input_path.with_name(f'{input_path.stem}_list.md')

    lines = input_path.read_text(encoding='utf-8').splitlines()
    rows = parse_md_table(lines)

    out_lines = []
    for term, definition in rows:
        if term and definition:
            out_lines.append(f'- **{term}:** {definition}')
        elif term:
            out_lines.append(f'- **{term}:**')
        elif definition:
            out_lines.append(f'- {definition}')

    output_path.write_text('\n'.join(out_lines), encoding='utf-8')
    print(f'Đã tạo: {output_path}')


if __name__ == '__main__':
    main()
