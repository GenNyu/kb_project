#!/usr/bin/env python3
from __future__ import annotations

import argparse
import copy
import json
import math
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}


@dataclass
class BodyItem:
    element: etree._Element
    text: str
    char_count: int
    kind: str


@dataclass
class Cluster:
    items: List[BodyItem]
    text: str
    char_count: int
    start_idx: int
    end_idx: int


@dataclass
class Chunk:
    clusters: List[Cluster]
    char_count: int


def iter_body_items(document_xml: bytes) -> tuple[list[BodyItem], etree._ElementTree, etree._Element, Optional[etree._Element]]:
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.fromstring(document_xml, parser=parser)
    body = tree.find("w:body", namespaces=NSMAP)
    if body is None:
        raise ValueError("Không tìm thấy w:body trong word/document.xml")

    sect_pr: Optional[etree._Element] = None
    items: list[BodyItem] = []

    for child in body:
        local_name = etree.QName(child).localname
        if local_name == "sectPr":
            sect_pr = copy.deepcopy(child)
            continue
        text = extract_text(child)
        items.append(
            BodyItem(
                element=copy.deepcopy(child),
                text=text,
                char_count=len(normalize_ws(text)),
                kind=local_name,
            )
        )
    return items, etree.ElementTree(tree), body, sect_pr


def extract_text(element: etree._Element) -> str:
    texts = element.xpath(".//w:t/text()", namespaces=NSMAP)
    return " ".join(t for t in texts if t is not None)


def normalize_ws(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def find_separator_template(
    items: Sequence[BodyItem],
    separator_regex: re.Pattern[str],
) -> tuple[Optional[BodyItem], int]:
    for item in items:
        clean = normalize_ws(item.text)
        if clean and separator_regex.fullmatch(clean):
            return item, item.char_count
    return None, 0


def build_clusters(
    items: Sequence[BodyItem],
    separator_regex: re.Pattern[str],
    start_regex: Optional[re.Pattern[str]],
    keep_separator_with_previous: bool = False,
) -> list[Cluster]:
    clusters: list[Cluster] = []
    current: list[BodyItem] = []
    current_start = 0

    def flush() -> None:
        nonlocal current, current_start
        if not current:
            return
        text = "\n".join(i.text for i in current if i.text)
        clusters.append(
            Cluster(
                items=current,
                text=text,
                char_count=sum(i.char_count for i in current),
                start_idx=current_start,
                end_idx=current_start + len(current) - 1,
            )
        )
        current = []

    for idx, item in enumerate(items):
        clean = normalize_ws(item.text)
        is_separator = bool(clean) and bool(separator_regex.fullmatch(clean))
        is_start = bool(clean) and start_regex is not None and bool(start_regex.search(clean))

        if is_separator:
            if keep_separator_with_previous and current:
                current.append(item)
                flush()
            else:
                flush()
            current_start = idx + 1
            continue

        if is_start and current:
            flush()
            current_start = idx

        if not current:
            current_start = idx
        current.append(item)

    flush()
    return [c for c in clusters if c.items]


def group_clusters_by_objective(
    clusters: Sequence[Cluster],
    objective_id_regex: re.Pattern[str],
) -> list[Chunk]:
    order: list[str] = []
    grouped: dict[str, list[Cluster]] = {}
    ungrouped: list[Chunk] = []

    for cluster in clusters:
        match = objective_id_regex.search(cluster.text)
        if not match:
            ungrouped.append(Chunk(clusters=[cluster], char_count=cluster.char_count))
            continue
        obj_id = match.group(0)
        if obj_id not in grouped:
            grouped[obj_id] = []
            order.append(obj_id)
        grouped[obj_id].append(cluster)

    chunks: list[Chunk] = []
    for obj_id in order:
        items = grouped[obj_id]
        chunks.append(
            Chunk(
                clusters=items,
                char_count=sum(c.char_count for c in items),
            )
        )
    chunks.extend(ungrouped)
    return chunks


def pack_clusters(
    clusters: Sequence[Cluster],
    max_chars: int,
    tolerance: int,
    separator_char_count: int = 0,
) -> list[Chunk]:
    if max_chars <= 0:
        raise ValueError("max_chars phải > 0")
    if tolerance < 0:
        raise ValueError("tolerance phải >= 0")

    lower = max(1, max_chars - tolerance)
    upper = max_chars + tolerance

    chunks: list[Chunk] = []
    current: list[Cluster] = []
    current_chars = 0

    def flush() -> None:
        nonlocal current, current_chars
        if current:
            chunks.append(Chunk(clusters=current, char_count=current_chars))
        current = []
        current_chars = 0

    for cluster in clusters:
        c = cluster.char_count

        if c > upper:
            flush()
            chunks.append(Chunk(clusters=[cluster], char_count=c))
            continue

        if not current:
            current = [cluster]
            current_chars = c
            continue

        proposed = current_chars + separator_char_count + c
        if proposed <= upper:
            current.append(cluster)
            current_chars = proposed
            continue

        # proposed > upper: so sánh giữ chunk hiện tại hay thêm cluster để gần max_chars hơn
        distance_if_stop = abs(max_chars - current_chars)
        distance_if_add = abs(max_chars - proposed)

        if current_chars < lower and distance_if_add <= distance_if_stop:
            current.append(cluster)
            current_chars = proposed
            flush()
        else:
            flush()
            current = [cluster]
            current_chars = c

    flush()
    return chunks


def replace_document_body(
    src_docx: Path,
    dest_docx: Path,
    chunk: Chunk,
    sect_pr: Optional[etree._Element],
    separator_item: Optional[BodyItem] = None,
    separator_text: Optional[str] = None,
) -> None:
    def build_separator_element() -> Optional[etree._Element]:
        if not separator_text:
            return copy.deepcopy(separator_item.element) if separator_item is not None else None
        # Create or reuse a paragraph element and overwrite its text runs.
        if separator_item is not None:
            elem = copy.deepcopy(separator_item.element)
        else:
            elem = etree.Element(f"{{{W_NS}}}p")
        # Remove existing text runs
        for t in elem.xpath(".//w:t", namespaces=NSMAP):
            parent = t.getparent()
            if parent is not None:
                parent.remove(t)
        # Ensure at least one run with text
        r = elem.find("w:r", namespaces=NSMAP)
        if r is None:
            r = etree.SubElement(elem, f"{{{W_NS}}}r")
        t = etree.SubElement(r, f"{{{W_NS}}}t")
        t.text = separator_text
        return elem

    separator_element = build_separator_element()
    with zipfile.ZipFile(src_docx, "r") as zin:
        document_xml = zin.read("word/document.xml")
        parser = etree.XMLParser(remove_blank_text=False)
        root = etree.fromstring(document_xml, parser=parser)
        body = root.find("w:body", namespaces=NSMAP)
        if body is None:
            raise ValueError("Không tìm thấy body trong tài liệu nguồn")

        for child in list(body):
            body.remove(child)

        for cluster_idx, cluster in enumerate(chunk.clusters):
            if cluster_idx > 0 and separator_element is not None:
                body.append(copy.deepcopy(separator_element))
            for item in cluster.items:
                body.append(copy.deepcopy(item.element))

        if sect_pr is not None:
            body.append(copy.deepcopy(sect_pr))

        new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        with zipfile.ZipFile(dest_docx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = new_xml if info.filename == "word/document.xml" else zin.read(info.filename)
                zout.writestr(info, data)


def summarize(chunks: Sequence[Chunk]) -> list[dict]:
    out: list[dict] = []
    for idx, chunk in enumerate(chunks, start=1):
        out.append(
            {
                "part": idx,
                "char_count": chunk.char_count,
                "cluster_count": len(chunk.clusters),
                "cluster_ranges": [
                    {"start_item": c.start_idx, "end_item": c.end_idx, "char_count": c.char_count}
                    for c in chunk.clusters
                ],
            }
        )
    return out


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Cắt file DOCX theo hạn mức ký tự +/- sai số nhưng không tách lẻ cụm dữ liệu. "
            "Cụm mặc định được tách bởi dòng phân cách như --- hoặc bởi regex bắt đầu cụm. "
            "Có thể chọn chế độ cắt theo Control objective dạng 1.1, 2.4, 3.1, v.v."
        )
    )
    parser.add_argument(
        "input_docx",
        type=Path,
        help="Đường dẫn file .docx nguồn hoặc thư mục chứa nhiều .docx",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=Path("out_docx_parts"),
        help="Thư mục chứa các file DOCX đã cắt",
    )
    parser.add_argument(
        "--max-chars",
        type=int,
        default=None,
        help="Hạn mức ký tự mục tiêu cho mỗi phần. Bắt buộc với chế độ start-regex; bỏ trống để objective không giới hạn ký tự.",
    )
    parser.add_argument("--tolerance", type=int, default=0, help="Sai số cho phép quanh max-chars")
    parser.add_argument(
        "--separator-regex",
        default=r"^(?:[-=*_]{3,})$",
        help="Regex nhận diện dòng phân cách giữa các cụm",
    )
    parser.add_argument(
        "--separator-text",
        default="========================================",
        help="Nếu đặt, sẽ thay dòng phân cách chèn giữa các cụm bằng chuỗi này (ví dụ ===...).",
    )
    parser.add_argument(
        "--cut-mode",
        choices=("start-regex", "objective"),
        default="start-regex",
        help="Chế độ nhận diện bắt đầu cụm: start-regex (mặc định) hoặc objective (theo số 1.1, 2.4, ...)",
    )
    parser.add_argument(
        "--start-regex",
        default=r"^Control objectives:\s*",
        help=(
            "Regex nhận diện dòng bắt đầu cụm mới. Đặt chuỗi rỗng để tắt. "
            "Mặc định phù hợp với mẫu PCI trong file hiện tại."
        ),
    )
    parser.add_argument(
        "--objective-regex",
        default=r"^\s*\d+\.\d+\b",
        help="Regex nhận diện dòng bắt đầu Control objective (ví dụ 1.1, 2.4, 3.1, ...)",
    )
    parser.add_argument(
        "--objective-id-regex",
        default=r"\b\d+\.\d+\b",
        help="Regex nhận diện mã Control objective để gom (ví dụ 2.3).",
    )
    parser.add_argument(
        "--no-group-objective",
        action="store_true",
        help="Tắt chế độ gộp các mục có cùng mã Control objective.",
    )
    parser.add_argument(
        "--keep-separator-with-previous",
        action="store_true",
        help="Nếu bật, dòng phân cách gốc sẽ được giữ ở cuối cụm trước đó. Mặc định separator chỉ được chèn giữa các cụm, không ở đầu/cuối phần.",
    )
    parser.add_argument(
        "--prefix",
        default=None,
        help="Tiền tố tên file xuất ra, ví dụ part_001.docx",
    )
    parser.add_argument(
        "--report-json",
        type=Path,
        default=None,
        help="Ghi thêm file JSON thống kê các phần",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Chỉ phân tích và in thống kê, không tạo file DOCX",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    if not args.input_docx.exists():
        print(f"Không tìm thấy file hoặc thư mục: {args.input_docx}", file=sys.stderr)
        return 2

    if args.cut_mode != "objective" and args.max_chars is None:
        print("Thiếu --max-chars cho chế độ start-regex", file=sys.stderr)
        return 2

    inputs = (
        sorted(p for p in args.input_docx.iterdir() if p.suffix.lower() == ".docx")
        if args.input_docx.is_dir()
        else [args.input_docx]
    )
    if not inputs:
        print("Không tìm thấy file .docx nào trong thư mục", file=sys.stderr)
        return 2

    separator_regex = re.compile(args.separator_regex)
    if args.cut_mode == "objective":
        start_regex = re.compile(args.objective_regex)
    else:
        start_regex = re.compile(args.start_regex) if args.start_regex else None

    overall_reports = []

    for input_docx in inputs:
        with zipfile.ZipFile(input_docx, "r") as zf:
            try:
                document_xml = zf.read("word/document.xml")
            except KeyError:
                print(f"File DOCX không hợp lệ: {input_docx}", file=sys.stderr)
                continue

        items, _, _, sect_pr = iter_body_items(document_xml)
        separator_item, separator_char_count = find_separator_template(items, separator_regex)
        clusters = build_clusters(
            items,
            separator_regex=separator_regex,
            start_regex=start_regex,
            keep_separator_with_previous=args.keep_separator_with_previous,
        )

        if not clusters:
            print(f"Không tạo được cụm dữ liệu nào: {input_docx}", file=sys.stderr)
            continue

        if args.cut_mode == "objective":
            if not args.no_group_objective:
                chunks = group_clusters_by_objective(
                    clusters,
                    objective_id_regex=re.compile(args.objective_id_regex),
                )
            else:
                chunks = [Chunk(clusters=[c], char_count=c.char_count) for c in clusters]
            if args.max_chars is not None and args.max_chars > 0:
                chunks = pack_clusters(
                    [c for chunk in chunks for c in chunk.clusters],
                    args.max_chars,
                    args.tolerance,
                    separator_char_count=0 if args.keep_separator_with_previous else separator_char_count,
                )
        else:
            chunks = pack_clusters(
                clusters,
                args.max_chars,
                args.tolerance,
                separator_char_count=0 if args.keep_separator_with_previous else separator_char_count,
            )

        report = summarize(chunks)
        overall_reports.append(
            {
                "input": str(input_docx),
                "items": len(items),
                "clusters": len(clusters),
                "parts": len(chunks),
                "max_chars": args.max_chars,
                "tolerance": args.tolerance,
                "report": report,
            }
        )

        if args.dry_run:
            continue

        args.output_dir.mkdir(parents=True, exist_ok=True)
        prefix = args.prefix if args.prefix else input_docx.stem
        width = max(3, len(str(len(chunks))))
        for i, chunk in enumerate(chunks, start=1):
            out_file = args.output_dir / f"{prefix}_{i:0{width}d}.docx"
            replace_document_body(
                input_docx,
                out_file,
                chunk,
                sect_pr,
                separator_item=None if args.keep_separator_with_previous else separator_item,
                separator_text=args.separator_text,
            )

    print(json.dumps(
        {"results": overall_reports},
        ensure_ascii=False,
        indent=2,
    ))

    if args.report_json:
        args.report_json.parent.mkdir(parents=True, exist_ok=True)
        args.report_json.write_text(
            json.dumps(overall_reports, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
