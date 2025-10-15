#!/usr/bin/env python3
"""Fetch profile bios from LinkedIn URLs stored in an Excel workbook.

The script reads the workbook located at the repository root (``LinkedIN.xlsx``),
requests the public profile pages and extracts the bio/description meta tag, and
writes a new workbook with the collected information to ``Results``.

Optional Features:
- Background Check (--bg): Uses OpenAI API to perform AI-powered background checks
  on LinkedIn profiles. Requires OpenAI API key in environment or .env file.
- Company Lookup (--company): Uses OpenAI API to find company information for the
  person's current employer, including email, phone, company type, industry, and revenue.
- Verbose Mode (--verbose): Prints detailed results to terminal as well as saving to Excel.

Output files are automatically timestamped to avoid conflicts.

The implementation relies solely on Python's standard library so that it works
in restricted execution environments where additional packages cannot be
installed. The OpenAI package is only required when using background checks or company lookup.
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
from datetime import datetime
from html import unescape
from html.parser import HTMLParser
from typing import Dict, Iterable, List, Optional, Sequence
from zipfile import ZipFile

NAMESPACE = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def load_env_file(env_path: str = None) -> Dict[str, str]:
    """Load environment variables from a .env file.
    
    Args:
        env_path: Path to the .env file. If None, looks for .env in the repository root.
        
    Returns:
        Dictionary of environment variables.
    """
    if env_path is None:
        env_path = os.path.join(os.path.dirname(__file__), "..", ".env")
    
    env_vars: Dict[str, str] = {}
    
    if not os.path.exists(env_path):
        return env_vars
    
    try:
        with open(env_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    env_vars[key.strip()] = value.strip()
    except Exception:
        # Silently ignore errors reading .env file
        pass
    
    return env_vars


def get_openai_api_key() -> str:
    """Get the OpenAI API key from environment variables or .env file.
    
    Returns:
        The OpenAI API key, or empty string if not found.
    """
    # First check environment variables
    api_key = os.environ.get('OPENAI_API_KEY', '')
    if api_key:
        return api_key
    
    # Then check .env file
    env_vars = load_env_file()
    return env_vars.get('OPENAI_API_KEY', '')


def perform_background_check(url: str, api_key: str) -> str:
    """Perform a background check on a LinkedIn profile using OpenAI.
    
    Args:
        url: LinkedIn profile URL
        api_key: OpenAI API key
        
    Returns:
        Background check results or error message
    """
    if not url or not api_key:
        return "Background check skipped: Missing URL or API key"
    
    try:
        # Import OpenAI here to avoid dependency issues if not using background checks
        try:
            from openai import OpenAI
        except ImportError:
            return "ERROR: OpenAI package not installed. Run: pip install openai"
        
        client = OpenAI(api_key=api_key)
        
        response = client.responses.create(
            model="gpt-4o",  # Using gpt-4o instead of gpt-5 as it's more widely available
            input=[
                {
                    "role": "developer",
                    "content": [
                        {
                            "type": "input_text",
                            "text": "What are the main achievements of this entrepreneur:"
                        }
                    ]
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": url
                        }
                    ]
                }
            ],
            text={
                "format": {
                    "type": "text"
                },
                "verbosity": "medium"
            },
            reasoning={
                "effort": "medium",
                "summary": "auto"
            },
            tools=[
                {
                    "type": "web_search",
                    "user_location": {
                        "type": "approximate",
                        "country": "US",
                        "city": "NYC"
                    },
                    "search_context_size": "medium"
                }
            ],
            store=False,
            include=[
                "reasoning.encrypted_content",
                "web_search_call.action.sources"
            ]
        )
        
        # Extract the response content
        if hasattr(response, 'content') and response.content:
            return response.content
        elif hasattr(response, 'text') and response.text:
            return response.text
        else:
            return "Background check completed but no content returned"
            
    except Exception as exc:
        return f"ERROR during background check: {exc}"


def lookup_company_info(url: str, api_key: str) -> str:
    """Look up company information for the person's current employer using OpenAI.
    
    Args:
        url: LinkedIn profile URL
        api_key: OpenAI API key
        
    Returns:
        Company information or error message
    """
    if not url or not api_key:
        return "Company lookup skipped: Missing URL or API key"
    
    try:
        # Import OpenAI here to avoid dependency issues if not using company lookup
        try:
            from openai import OpenAI
        except ImportError:
            return "ERROR: OpenAI package not installed. Run: pip install openai"
        
        client = OpenAI(api_key=api_key)
        
        response = client.responses.create(
            model="gpt-4o",
            input=[
                {
                    "role": "developer",
                    "content": [
                        {
                            "type": "input_text",
                            "text": "Find the current company information for this person. Provide the following details in English: Email, Phone number, Company type, Industry, Latest revenue. If information is not available, write 'Not available' for that field."
                        }
                    ]
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": url
                        }
                    ]
                }
            ],
            text={
                "format": {
                    "type": "text"
                },
                "verbosity": "medium"
            },
            reasoning={
                "effort": "medium",
                "summary": "auto"
            },
            tools=[
                {
                    "type": "web_search",
                    "user_location": {
                        "type": "approximate",
                        "country": "US",
                        "city": "NYC"
                    },
                    "search_context_size": "medium"
                }
            ],
            store=False,
            include=[
                "reasoning.encrypted_content",
                "web_search_call.action.sources"
            ]
        )
        
        # Extract the response content
        if hasattr(response, 'content') and response.content:
            return response.content
        elif hasattr(response, 'text') and response.text:
            return response.text
        else:
            return "Company lookup completed but no content returned"
            
    except Exception as exc:
        return f"ERROR during company lookup: {exc}"


def print_verbose_results(row_number: int, url: str, bio: str, bg_check: str = None, company_info: str = None) -> None:
    """Print results to terminal in verbose mode.
    
    Args:
        row_number: Row number being processed
        url: LinkedIn URL
        bio: Extracted bio
        bg_check: Background check results (optional)
        company_info: Company information (optional)
    """
    print(f"\n{'='*80}")
    print(f"ROW {row_number} RESULTS")
    print(f"{'='*80}")
    print(f"URL: {url}")
    print(f"\nBIO:")
    print(f"{'-'*40}")
    print(bio if bio else "No bio found")
    
    if bg_check is not None:
        print(f"\nBACKGROUND CHECK:")
        print(f"{'-'*40}")
        print(bg_check)
    
    if company_info is not None:
        print(f"\nCOMPANY INFORMATION:")
        print(f"{'-'*40}")
        print(company_info)
    
    print(f"{'='*80}\n")


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
    openai_api_key: str
    background_check: bool
    company_lookup: bool
    verbose: bool


def parse_args(argv: Sequence[str]) -> ExecutionConfig:
    parser = argparse.ArgumentParser(description=__doc__)
    
    # Generate timestamped default output filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    default_output = os.path.join(os.path.dirname(__file__), "..", "Results", f"LinkedIn_Bios_{timestamp}.xlsx")
    
    parser.add_argument(
        "--input",
        default=os.path.join(os.path.dirname(__file__), "..", "LinkedIN.xlsx"),
        help="Path to the input Excel workbook (default: LinkedIN.xlsx at repository root)",
    )
    parser.add_argument(
        "--output",
        default=default_output,
        help=f"Path for the generated workbook (default: Results/LinkedIn_Bios_{{timestamp}}.xlsx)",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.0,
        help="Optional delay in seconds between HTTP requests",
    )
    parser.add_argument(
        "--bg",
        action="store_true",
        default=False,
        help="Enable background check using OpenAI (requires OpenAI API key)",
    )
    parser.add_argument(
        "--company",
        action="store_true",
        default=False,
        help="Enable company information lookup using OpenAI (requires OpenAI API key)",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        default=False,
        help="Enable verbose mode - print results to terminal as well as saving to Excel",
    )
    args = parser.parse_args(argv)
    input_path = os.path.abspath(args.input)
    output_path = os.path.abspath(args.output)
    openai_api_key = get_openai_api_key()
    return ExecutionConfig(
        input_path=input_path, 
        output_path=output_path, 
        delay=max(0.0, args.delay),
        openai_api_key=openai_api_key,
        background_check=args.bg,
        company_lookup=args.company,
        verbose=args.verbose
    )


def main(argv: Sequence[str] | None = None) -> int:
    config = parse_args(argv or sys.argv[1:])

    # Display OpenAI API key status (without exposing the actual key)
    if config.openai_api_key:
        print(f"OpenAI API key loaded (length: {len(config.openai_api_key)} characters)")
    else:
        print("No OpenAI API key found in environment or .env file")
    
    # Display background check status
    if config.background_check:
        if config.openai_api_key:
            print("Background check enabled - will perform AI-powered background checks")
        else:
            print("WARNING: Background check requested but no OpenAI API key available")
    else:
        print("Background check disabled")
    
    # Display company lookup status
    if config.company_lookup:
        if config.openai_api_key:
            print("Company lookup enabled - will perform AI-powered company information lookup")
        else:
            print("WARNING: Company lookup requested but no OpenAI API key available")
    else:
        print("Company lookup disabled")
    
    # Display verbose mode status
    if config.verbose:
        print("Verbose mode enabled - results will be printed to terminal")
    else:
        print("Verbose mode disabled")

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

    # Add background check column if background check is enabled
    if config.background_check:
        try:
            bg_check_index = header.index("Background Check")
        except ValueError:
            bg_check_index = len(header)
            header = header + ["Background Check"]
    else:
        bg_check_index = None

    # Add company information column if company lookup is enabled
    if config.company_lookup:
        try:
            company_info_index = header.index("Company Info")
        except ValueError:
            company_info_index = len(header)
            header = header + ["Company Info"]
    else:
        company_info_index = None

    output_rows: List[List[str]] = [list(header)]

    for row_number, row in enumerate(rows_data, start=2):
        # Ensure row has enough columns for all operations
        max_needed_indices = [bio_index]
        if bg_check_index is not None:
            max_needed_indices.append(bg_check_index)
        if company_info_index is not None:
            max_needed_indices.append(company_info_index)
        max_needed_index = max(max_needed_indices)
        
        row = list(row) + [""] * max(0, max_needed_index + 1 - len(row))
        
        url = row[url_index].strip()
        if url:
            print(f"Processing row {row_number}: {url}")
            
            # Fetch bio
            print(f"  - Fetching bio...")
            bio = fetch_profile_bio(url)
            row[bio_index] = bio
            
            # Initialize variables for verbose output
            bg_check_result = None
            company_info_result = None
            
            # Perform background check if enabled
            if config.background_check and bg_check_index is not None:
                print(f"  - Performing background check...")
                bg_check_result = perform_background_check(url, config.openai_api_key)
                row[bg_check_index] = bg_check_result
            
            # Perform company lookup if enabled
            if config.company_lookup and company_info_index is not None:
                print(f"  - Looking up company information...")
                company_info_result = lookup_company_info(url, config.openai_api_key)
                row[company_info_index] = company_info_result
            
            # Print verbose results if enabled
            if config.verbose:
                print_verbose_results(
                    row_number=row_number,
                    url=url,
                    bio=bio,
                    bg_check=bg_check_result,
                    company_info=company_info_result
                )
            
            if config.delay:
                time.sleep(config.delay)
        else:
            row[bio_index] = ""
            if config.background_check and bg_check_index is not None:
                row[bg_check_index] = ""
            if config.company_lookup and company_info_index is not None:
                row[company_info_index] = ""
            
            if config.verbose:
                print(f"\nRow {row_number}: No URL provided - skipping\n")
        
        output_rows.append(row[: len(header)])

    os.makedirs(os.path.dirname(config.output_path), exist_ok=True)
    write_rows_to_workbook(config.output_path, output_rows)
    print(f"Saved results to {config.output_path}")
    
    # Print summary in verbose mode
    if config.verbose:
        total_processed = len([row for row in rows_data if row[url_index].strip()])
        print(f"\n{'='*80}")
        print(f"PROCESSING SUMMARY")
        print(f"{'='*80}")
        print(f"Total profiles processed: {total_processed}")
        print(f"Output file: {config.output_path}")
        if config.background_check:
            print("✓ Background checks performed")
        if config.company_lookup:
            print("✓ Company information lookups performed")
        print(f"{'='*80}")
    
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
