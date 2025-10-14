#!/usr/bin/env python3
"""Fetch profile bios from LinkedIn URLs stored in an Excel workbook.

The script reads the workbook located at the repository root (``LinkedIN.xlsx``),
requests the public profile pages and extracts the bio/description meta tag, and
writes a new workbook with the collected information to ``Results``.

The implementation relies solely on Python's standard library so that it works
in restricted execution environments where additional packages cannot be
installed.
"""
from __future__ import annotations

import argparse
import os
import sys
import time
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from html import unescape
from html.parser import HTMLParser
from typing import Dict, Iterable, List, Sequence
from zipfile import ZipFile

NAMESPACE = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


class MetaTagParser(HTMLParser):
    """Simple HTML parser that collects meta tag content by name/property."""

    def __init__(self) -> None:
        super().__init__()
        self.meta: Dict[str, List[str]] = {}

    def handle_starttag(self, tag: str, attrs: List[tuple[str, str | None]]) -> None:  # type: ignore[override]
        if tag.lower() != "meta":
            return
        attr_dict = {k.lower(): (v or "") for k, v in attrs}
        key = attr_dict.get("property") or attr_dict.get("name")
        if not key:
            return
        content = attr_dict.get("content", "")
        self.meta.setdefault(key.lower(), []).append(content)


def column_letters_to_index(column: str) -> int:
    """Convert Excel column letters (e.g. ``"A"``) to a zero-based index."""
    index = 0
    for char in column:
        if not char.isalpha():
            break
        index = index * 26 + (ord(char.upper()) - ord("A") + 1)
    return index - 1


def column_index_to_letters(index: int) -> str:
    """Convert a zero-based column index to Excel column letters."""
    if index < 0:
        raise ValueError("Column index must be non-negative")
    result = ""
    current = index + 1
    while current:
        current, remainder = divmod(current - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def read_shared_strings(zf: ZipFile) -> List[str]:
    """Load the shared strings table if present."""
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    strings: List[str] = []
    for si in root.findall("main:si", NAMESPACE):
        text_fragments: List[str] = []
        t_node = si.find("main:t", NAMESPACE)
        if t_node is not None and t_node.text is not None:
            text_fragments.append(t_node.text)
        else:
            for run in si.findall("main:r", NAMESPACE):
                run_text = run.find("main:t", NAMESPACE)
                if run_text is not None and run_text.text is not None:
                    text_fragments.append(run_text.text)
        strings.append("".join(text_fragments))
    return strings


def extract_cell_text(cell: ET.Element, shared: Sequence[str]) -> str:
    """Return the string value for a cell element."""
    cell_type = cell.get("t")
    if cell_type == "s":
        value = cell.findtext("main:v", default="", namespaces=NAMESPACE)
        if not value:
            return ""
        try:
            return shared[int(value)]
        except (IndexError, ValueError):
            return ""
    if cell_type == "inlineStr":
        texts = [
            node.text or ""
            for node in cell.findall("main:is/main:t", NAMESPACE)
        ]
        return "".join(texts)
    value = cell.findtext("main:v", default="", namespaces=NAMESPACE)
    return value or ""


def read_rows_from_workbook(path: str) -> List[List[str]]:
    """Return rows of the first worksheet as lists of strings."""
    with ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        sheet_data = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        sheet_data_element = sheet_data.find("main:sheetData", NAMESPACE)
        if sheet_data_element is None:
            return []
        rows_raw: List[Dict[int, str]] = []
        max_col = 0
        for row in sheet_data_element:
            cell_values: Dict[int, str] = {}
            for cell in row.findall("main:c", NAMESPACE):
                ref = cell.get("r", "")
                letters = "".join(ch for ch in ref if ch.isalpha())
                if not letters:
                    continue
                column_index = column_letters_to_index(letters)
                value = extract_cell_text(cell, shared_strings)
                cell_values[column_index] = value
                max_col = max(max_col, column_index + 1)
            rows_raw.append(cell_values)
        rows: List[List[str]] = []
        for mapping in rows_raw:
            row_values = [mapping.get(idx, "") for idx in range(max_col)]
            rows.append(row_values)
        return rows


def fetch_profile_bio(url: str, timeout: float = 15.0) -> str:
    """Request the LinkedIn profile page and extract the description meta tag."""
    if not url:
        return ""
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/123.0 Safari/537.36",
            "Accept-Language": "en",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            html_bytes = response.read()
    except urllib.error.URLError as exc:  # pragma: no cover - network errors depend on environment
        return f"ERROR: {exc.reason if hasattr(exc, 'reason') else exc}"
    except Exception as exc:  # pragma: no cover - safety net
        return f"ERROR: {exc}"

    try:
        html_text = html_bytes.decode("utf-8", errors="ignore")
    except Exception:  # pragma: no cover
        html_text = html_bytes.decode("latin-1", errors="ignore")

    parser = MetaTagParser()
    parser.feed(html_text)
    for key in ("og:description", "description"):
        values = parser.meta.get(key)
        if not values:
            continue
        for value in values:
            cleaned = unescape((value or "").strip())
            if cleaned:
                return cleaned
    return "Bio not found"


def sanitize_rows(rows: Iterable[Sequence[str]]) -> List[List[str]]:
    """Ensure rows are stored as lists of strings."""
    normalized: List[List[str]] = []
    for row in rows:
        normalized.append([str(cell) if cell is not None else "" for cell in row])
    return normalized


def build_sheet_xml(rows: Sequence[Sequence[str]]) -> bytes:
    from xml.sax.saxutils import escape

    lines = [
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
        "  <sheetData>",
    ]
    for row_idx, row in enumerate(rows, start=1):
        cells_xml: List[str] = []
        for col_idx, value in enumerate(row, start=1):
            if value == "":
                continue
            cell_ref = f"{column_index_to_letters(col_idx - 1)}{row_idx}"
            text = escape(value)
            cell_xml = (
                f"      <c r=\"{cell_ref}\" t=\"inlineStr\">"
                f"<is><t xml:space=\"preserve\">{text}</t></is></c>"
            )
            cells_xml.append(cell_xml)
        if cells_xml:
            lines.append(f"    <row r=\"{row_idx}\">")
            lines.extend(cells_xml)
            lines.append("    </row>")
        else:
            lines.append(f"    <row r=\"{row_idx}\"/>")
    lines.append("  </sheetData>")
    lines.append("</worksheet>")
    xml_text = "\n".join(lines)
    return xml_text.encode("utf-8")


def write_rows_to_workbook(path: str, rows: Sequence[Sequence[str]]) -> None:
    rows = sanitize_rows(rows)
    sheet_xml = build_sheet_xml(rows)
    content_types_xml = b"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
</Types>
"""
    rels_xml = b"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
</Relationships>
"""
    workbook_xml = b"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <sheets>
    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>
  </sheets>
</workbook>
"""
    workbook_rels_xml = b"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
</Relationships>
"""

    with ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


@dataclass
class ExecutionConfig:
    input_path: str
    output_path: str
    delay: float


def parse_args(argv: Sequence[str]) -> ExecutionConfig:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        default=os.path.join(os.path.dirname(__file__), "..", "LinkedIN.xlsx"),
        help="Path to the input Excel workbook (default: LinkedIN.xlsx at repository root)",
    )
    parser.add_argument(
        "--output",
        default=os.path.join(os.path.dirname(__file__), "..", "Results", "LinkedIn_Bios.xlsx"),
        help="Path for the generated workbook (default: Results/LinkedIn_Bios.xlsx)",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.0,
        help="Optional delay in seconds between HTTP requests",
    )
    args = parser.parse_args(argv)
    input_path = os.path.abspath(args.input)
    output_path = os.path.abspath(args.output)
    return ExecutionConfig(input_path=input_path, output_path=output_path, delay=max(0.0, args.delay))


def main(argv: Sequence[str] | None = None) -> int:
    config = parse_args(argv or sys.argv[1:])

    if not os.path.exists(config.input_path):
        print(f"Input workbook not found: {config.input_path}", file=sys.stderr)
        return 1

    rows = read_rows_from_workbook(config.input_path)
    if not rows:
        print("Input workbook does not contain any rows", file=sys.stderr)
        return 1

    header = rows[0]
    rows_data = rows[1:]

    try:
        url_index = header.index("LinkedIn Page")
    except ValueError:
        print("Could not find 'LinkedIn Page' column in the input workbook", file=sys.stderr)
        return 1

    try:
        bio_index = header.index("Bio")
    except ValueError:
        bio_index = len(header)
        header = header + ["Bio"]

    output_rows: List[List[str]] = [list(header)]

    for row_number, row in enumerate(rows_data, start=2):
        row = list(row) + [""] * max(0, bio_index + 1 - len(row))
        url = row[url_index].strip()
        if url:
            print(f"Fetching bio for row {row_number}: {url}")
            bio = fetch_profile_bio(url)
            if config.delay:
                time.sleep(config.delay)
        else:
            bio = ""
        row[bio_index] = bio
        output_rows.append(row[: len(header)])

    os.makedirs(os.path.dirname(config.output_path), exist_ok=True)
    write_rows_to_workbook(config.output_path, output_rows)
    print(f"Saved results to {config.output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
