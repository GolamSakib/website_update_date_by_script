#!/usr/bin/env python3
"""
Scrape last-update dates from https://mowr.gov.bd/ and update the Word document table.

For each row in the docx that contains a numbered topic (e.g. ১.১, ১.২, ১০.১.১),
the script:
  1. Extracts the topic title after the number prefix
  2. Finds the matching <a title="..."> link on the MOWR homepage
  3. Fetches that page and reads the date from div.content-update-block > p
  4. Writes the date (dd/mm/yyyy in Bengali numerals) into column
     "সর্বশেষ হালনাগাদের তারিখ" using Nikosh 14pt, centered
  5. Applies the same font, size, and alignment to every data cell in that
     column, including rows that keep their existing date
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from urllib.parse import urljoin

import requests
import urllib3
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://mowr.gov.bd/"
DATE_COLUMN_INDEX = 4
TOPIC_COLUMN_INDEX = 2
DATE_FONT_NAME = "Nikosh"
DATE_FONT_SIZE_PT = 14

BN_DIGITS = "০১২৩৪৫৬৭৮৯"
EN_DIGITS = "0123456789"

# Matches numbered topics like ১.১, ১.২, ১০.১.১ at the start of a cell.
NUMBERED_TOPIC_RE = re.compile(
    r"^([\u09E6-\u09EF0-9]+(?:\.[\u09E6-\u09EF0-9]+)+)\s+(.+)$"
)

# Extract day, month name, year from update paragraph text.
UPDATE_DATE_RE = re.compile(
    r"(\d{1,2}|[\u09E6-\u09EF]{1,2})\s+"
    r"([\u0980-\u09FFa-zA-Z\.]+)\s*,?\s*"
    r"(\d{4}|[\u09E6-\u09EF]{4})"
)

MONTHS = {
    "january": 1,
    "jan": 1,
    "february": 2,
    "feb": 2,
    "march": 3,
    "mar": 3,
    "april": 4,
    "apr": 4,
    "may": 5,
    "june": 6,
    "jun": 6,
    "july": 7,
    "jul": 7,
    "august": 8,
    "aug": 8,
    "september": 9,
    "sep": 9,
    "sept": 9,
    "october": 10,
    "oct": 10,
    "november": 11,
    "nov": 11,
    "december": 12,
    "dec": 12,
    "জানুয়ারি": 1,
    "জানু": 1,
    "ফেব্রুয়ারি": 2,
    "ফেব্রু": 2,
    "মার্চ": 3,
    "এপ্রিল": 4,
    "মে": 5,
    "জুন": 6,
    "জুলাই": 7,
    "আগস্ট": 8,
    "সেপ্টেম্বর": 9,
    "অক্টোবর": 10,
    "নভেম্বর": 11,
    "ডিসেম্বর": 12,
}


@dataclass
class DocItem:
    row_index: int
    number_prefix: str
    title: str
    raw_cell_text: str


@dataclass
class SiteLink:
    title: str
    href: str
    absolute_url: str


def en_to_bn(num: int, width: int = 0) -> str:
    text = str(num)
    if width:
        text = text.zfill(width)
    return "".join(BN_DIGITS[EN_DIGITS.index(ch)] for ch in text)


def bn_to_en(text: str) -> str:
    table = str.maketrans(BN_DIGITS, EN_DIGITS)
    return text.translate(table)


def normalize_title(text: str) -> str:
    text = text.strip()
    text = re.sub(r"\s+", " ", text)
    return text


def title_words(text: str) -> set[str]:
    text = normalize_title(text)
    text = text.replace(" ও ", " ")
    return {word for word in re.split(r"\s+", text) if word}


def titles_match(doc_title: str, site_title: str) -> bool:
    doc = normalize_title(doc_title)
    site = normalize_title(site_title)
    if doc == site:
        return True
    if doc in site or site in doc:
        return True

    doc_words = title_words(doc)
    site_words = title_words(site)
    if doc_words and doc_words == site_words:
        return True

    if not doc_words or not site_words:
        return False

    overlap = len(doc_words & site_words)
    threshold = min(len(doc_words), len(site_words))
    return overlap >= threshold and overlap >= max(1, threshold - 1)


def parse_topic_cell(cell_text: str) -> Optional[tuple[str, str]]:
    match = NUMBERED_TOPIC_RE.match(normalize_title(cell_text))
    if not match:
        return None
    return match.group(1), normalize_title(match.group(2))


def extract_items_from_doc(doc: Document) -> list[DocItem]:
    table = doc.tables[0]
    items: list[DocItem] = []

    for row_index, row in enumerate(table.rows):
        if row_index < 2:
            continue

        topic_cell = row.cells[TOPIC_COLUMN_INDEX].text
        parsed = parse_topic_cell(topic_cell)
        if not parsed:
            continue

        number_prefix, title = parsed
        items.append(
            DocItem(
                row_index=row_index,
                number_prefix=number_prefix,
                title=title,
                raw_cell_text=normalize_title(topic_cell),
            )
        )
    return items


def fetch_homepage_links(session: requests.Session) -> list[SiteLink]:
    response = session.get(BASE_URL, timeout=45, verify=False)
    response.raise_for_status()
    response.encoding = "utf-8"
    soup = BeautifulSoup(response.text, "html.parser")

    links: list[SiteLink] = []
    seen: set[tuple[str, str]] = set()

    for anchor in soup.find_all("a", title=True):
        title = normalize_title(anchor.get("title", ""))
        href = (anchor.get("href") or "").strip()
        if not title or not href or href == "#":
            continue

        absolute_url = urljoin(BASE_URL, href)
        key = (title, absolute_url)
        if key in seen:
            continue
        seen.add(key)
        links.append(SiteLink(title=title, href=href, absolute_url=absolute_url))

    return links


def find_matching_link(title: str, links: list[SiteLink]) -> Optional[SiteLink]:
    exact = [link for link in links if normalize_title(link.title) == normalize_title(title)]
    if exact:
        return prefer_mowr_link(exact)

    fuzzy = [link for link in links if titles_match(title, link.title)]
    if fuzzy:
        return prefer_mowr_link(fuzzy)

    return None


def prefer_mowr_link(links: list[SiteLink]) -> SiteLink:
    internal = [
        link
        for link in links
        if "mowr.gov.bd" in link.absolute_url and not link.absolute_url.rstrip("/").endswith("mowr.gov.bd")
    ]
    return internal[0] if internal else links[0]


def parse_update_paragraph(text: str) -> Optional[str]:
    match = UPDATE_DATE_RE.search(text)
    if not match:
        return None

    day_raw, month_raw, year_raw = match.groups()
    day = int(bn_to_en(day_raw))
    year = int(bn_to_en(year_raw))

    month_key = normalize_title(month_raw).lower()
    month = MONTHS.get(month_key)
    if month is None:
        month_key_bn = normalize_title(month_raw)
        month = MONTHS.get(month_key_bn)
    if month is None:
        return None

    return f"{en_to_bn(day, 2)}/{en_to_bn(month, 2)}/{en_to_bn(year, 4)}"


def fetch_update_date(session: requests.Session, url: str) -> Optional[str]:
    response = session.get(url, timeout=45, verify=False)
    response.raise_for_status()
    response.encoding = "utf-8"
    soup = BeautifulSoup(response.text, "html.parser")

    block = soup.find("div", class_="content-update-block")
    if not block:
        return None

    paragraph = block.find("p")
    if not paragraph:
        return None

    return parse_update_paragraph(paragraph.get_text(" ", strip=True))


def _set_run_font(run, font_name: str = DATE_FONT_NAME, font_size: int = DATE_FONT_SIZE_PT) -> None:
    run.font.name = font_name
    run.font.size = Pt(font_size)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    r_fonts.set(qn("w:ascii"), font_name)
    r_fonts.set(qn("w:hAnsi"), font_name)
    r_fonts.set(qn("w:eastAsia"), font_name)
    r_fonts.set(qn("w:cs"), font_name)


def _set_cell_vertical_center(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    v_align = tc_pr.find(qn("w:vAlign"))
    if v_align is None:
        v_align = OxmlElement("w:vAlign")
        tc_pr.append(v_align)
    v_align.set(qn("w:val"), "center")


def set_cell_text(cell, value: str) -> None:
    """Write cell text as Nikosh 14pt, horizontally and vertically centered."""
    cell.text = value
    _set_cell_vertical_center(cell)
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if not paragraph.runs:
            if value:
                run = paragraph.add_run(value)
                _set_run_font(run)
            continue
        for run in paragraph.runs:
            _set_run_font(run)


def is_section_banner_row(row) -> bool:
    texts = [cell.text.strip() for cell in row.cells]
    return bool(texts) and len(set(texts)) == 1


def format_date_column(table) -> None:
    """Apply Nikosh 14pt and center alignment to every data cell in the date column."""
    for row_index, row in enumerate(table.rows):
        if row_index < 2 or is_section_banner_row(row):
            continue
        cell = row.cells[DATE_COLUMN_INDEX]
        set_cell_text(cell, cell.text.strip())


def process_document(
    doc_path: Path,
    *,
    output_path: Optional[Path] = None,
    dry_run: bool = False,
    delay_seconds: float = 0.4,
    limit: Optional[int] = None,
) -> int:
    if not doc_path.exists():
        raise FileNotFoundError(f"Document not found: {doc_path}")

    backup_path = doc_path.with_name(doc_path.stem + "_backup.docx")
    if not dry_run and not backup_path.exists():
        shutil.copy2(doc_path, backup_path)
        print(f"Backup created: {backup_path}")

    doc = Document(str(doc_path))
    items = extract_items_from_doc(doc)
    if limit is not None:
        items = items[:limit]

    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
        }
    )

    print(f"Loaded {len(items)} numbered topics from document.")
    print("Fetching homepage links...")
    homepage_links = fetch_homepage_links(session)
    print(f"Found {len(homepage_links)} titled links on homepage.\n")

    table = doc.tables[0]
    updated_count = 0
    failures: list[str] = []

    for index, item in enumerate(items, start=1):
        print(f"[{index}/{len(items)}] {item.raw_cell_text}")

        link = find_matching_link(item.title, homepage_links)
        date_cell = table.rows[item.row_index].cells[DATE_COLUMN_INDEX]
        old_date = normalize_title(date_cell.text)

        if not link:
            message = f"  -> No matching link for title: {item.title}"
            print(message)
            failures.append(message)
            if not dry_run:
                set_cell_text(date_cell, old_date)
            continue

        print(f"  -> Matched: {link.title!r}")
        print(f"  -> URL: {link.absolute_url}")

        try:
            new_date = fetch_update_date(session, link.absolute_url)
        except requests.RequestException as exc:
            message = f"  -> Request failed: {exc}"
            print(message)
            failures.append(message)
            if not dry_run:
                set_cell_text(date_cell, old_date)
            time.sleep(delay_seconds)
            continue

        if not new_date:
            message = f"  -> No update date found on page for: {item.title}"
            print(message)
            failures.append(message)
            if not dry_run:
                set_cell_text(date_cell, old_date)
            time.sleep(delay_seconds)
            continue

        print(f"  -> Old date: {old_date or '(empty)'}")
        print(f"  -> New date: {new_date}")

        if not dry_run:
            set_cell_text(date_cell, new_date)
            updated_count += 1

        time.sleep(delay_seconds)

    if not dry_run:
        format_date_column(table)
        target_path = output_path or doc_path
        try:
            doc.save(str(target_path))
        except PermissionError:
            target_path = doc_path.with_name(doc_path.stem + "_updated.docx")
            doc.save(str(target_path))
            print(
                f"\nOriginal file is open/locked. Saved to: {target_path}",
                file=sys.stderr,
            )
        print(f"\nSaved updated document: {target_path}")
        print(f"Updated rows: {updated_count}")
    else:
        print("\nDry run complete. Document was not modified.")

    if failures:
        print(f"\nIssues ({len(failures)}):")
        for failure in failures:
            print(failure)

    return updated_count


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Update MOWR website last-modified dates in the Word report."
    )
    parser.add_argument(
        "--doc",
        default="ওয়েবসাইট_হালনাগাদ_তথ্য_আপডেটেড.docx",
        help="Path to the Word document (.docx)",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output .docx path (defaults to overwriting --doc)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Fetch and print results without modifying the document",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Process only the first N numbered topics (for testing)",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.4,
        help="Delay in seconds between page requests",
    )
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
    doc_path = Path(args.doc)
    if not doc_path.is_absolute():
        doc_path = script_dir / doc_path

    output_path = Path(args.output) if args.output else None
    if output_path and not output_path.is_absolute():
        output_path = script_dir / output_path

    try:
        process_document(
            doc_path,
            output_path=output_path,
            dry_run=args.dry_run,
            delay_seconds=args.delay,
            limit=args.limit,
        )
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
