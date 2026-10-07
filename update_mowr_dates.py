#!/usr/bin/env python3
"""
Scrape last-update dates from https://mowr.gov.bd/ and update the Word document table.

For each row in the docx that contains a numbered topic (e.g. ১.১, ১.২, ১০.১.১),
the script:
  1. Extracts the topic title after the number prefix
  2. Finds the matching <a title="..."> link on the MOWR homepage
  3. Fetches that page and reads the date from div.content-update-block > p

     SPECIAL CASE: if the topic title is "ই-সেবাসমূহের তালিকা" or
     "মন্ত্রণালয়/বিভাগ সম্পর্কিত কমিটিসমূহ", step 3 is different:
       3a. Open the matched page
       3b. Find every <td> that has BOTH classes "table-td" and "centered"
       3c. Open every link found inside those cells, one by one
       3d. Read each sub-page's div.content-update-block date
       3e. Keep the MOST RECENT date among all of them

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
from datetime import date
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
    "জানুয়ারী": 1,
    "জানু": 1,
    "ফেব্রুয়ারি": 2,
    "ফেব্রুয়ারী": 2,
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
    import unicodedata
    text = text.strip()
    text = re.sub(r"\s+", " ", text)
    text = unicodedata.normalize("NFC", text)
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


def _has_real_path(url: str) -> bool:
    """Return True if a URL has a meaningful path beyond just the domain root."""
    from urllib.parse import urlparse
    path = urlparse(url).path.strip("/")
    return bool(path)


def prefer_mowr_link(links: list[SiteLink]) -> SiteLink:
    # 1st priority: mowr.gov.bd links with an actual path
    internal = [
        link
        for link in links
        if "mowr.gov.bd" in link.absolute_url and _has_real_path(link.absolute_url)
    ]
    if internal:
        return internal[0]

    # 2nd priority: any link with an actual path (not just a bare domain root)
    with_path = [link for link in links if _has_real_path(link.absolute_url)]
    if with_path:
        return with_path[0]

    return links[0]


# ---------------------------------------------------------------------------
# Date parsing / formatting
# ---------------------------------------------------------------------------

def parse_update_date_value(text: str) -> Optional[date]:
    """Parse a date out of an update paragraph and return a real date object."""
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

    try:
        return date(year, month, day)
    except ValueError:
        return None


def format_bengali_date(value: date) -> str:
    return f"{en_to_bn(value.day, 2)}/{en_to_bn(value.month, 2)}/{en_to_bn(value.year, 4)}"


def parse_update_paragraph(text: str) -> Optional[str]:
    """Kept for backward compatibility: returns the Bengali-formatted string."""
    value = parse_update_date_value(text)
    return format_bengali_date(value) if value else None


# ---------------------------------------------------------------------------
# Fetching
# ---------------------------------------------------------------------------

def fetch_update_date_value(session: requests.Session, url: str) -> Optional[date]:
    """Fetch a page and return its content-update-block date as a date object."""
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

    return parse_update_date_value(paragraph.get_text(" ", strip=True))


def fetch_update_date(session: requests.Session, url: str) -> Optional[str]:
    """Fetch a page and return its content-update-block date, Bengali-formatted."""
    value = fetch_update_date_value(session, url)
    return format_bengali_date(value) if value else None


# Topics that require the "open every linked sub-page, take the latest date"
# logic instead of reading a single content-update-block directly.
# Both the correct spelling and the "সম্পর্কত" variant are accepted.
SPECIAL_MULTI_LINK_TITLES = {
    normalize_title("ই-সেবাসমূহের তালিকা"),
    normalize_title("মন্ত্রণালয়/বিভাগ সম্পর্কিত কমিটিসমূহ"),
    normalize_title("প্রজ্ঞাপন/ অফিস আদেশ/পরিপত্র/যোগদানপত্র"),
    normalize_title("আইন/বিধি/নীতিমালা/পরিপত্র/প্রজ্ঞাপন"),
    normalize_title("ফোকাল পয়েন্ট কর্মকর্তা"),
    normalize_title("মন্ত্রণালয়/বিভাগ সম্পর্কত কমিটিসমূহ"),
    normalize_title("প্রকাশনা ও বার্ষিক প্রতিবেদন"),
    normalize_title("আইন ও বিধি"),
    normalize_title("সকল ফোকাল পয়েন্ট"),
    normalize_title("উদ্ভাবন কার্যক্রম"),
    normalize_title("সভার কার্যবিবরণী"),
    normalize_title("সভার বিজ্ঞপ্তি"),
    normalize_title("সাংগঠনিক কাঠামো"),
    normalize_title("বিদেশ ভ্রমণের জি.ও"),
    normalize_title("বহিঃ বাংলাদেশ ছুটি"),
    normalize_title("পাসপোর্ট অনাপত্তি পত্র"),
    normalize_title("কর্মপরিকল্পনা, পরিবীক্ষণ ও মূল্যায়ন প্রতিবেদন"),
    normalize_title("নীতিমালা"),
    normalize_title("খসড়া"),
    normalize_title("অন্যান্য"),
    normalize_title("ত্রৈমাসিক/বার্ষিক পরিবীক্ষণ/মূল্যায়ন প্রতিবেদন"),
    normalize_title("কর্মপরিকল্পনা/পরিবীক্ষণ ও মূল্যায়ন প্রতিবেদন"),
    normalize_title("কর্মপরিকল্পনা /নির্দেশিকা/প্রকল্পসমূহ"),
    normalize_title("বার্ষিক ক্রয় পরিকল্পনা"),
    normalize_title("বাজেট ও এমটিবিএফ বাজেট"),
    normalize_title("টেন্ডার / কোটেশন / নিলাম"),
    normalize_title("নিয়োগ সংক্রান্ত সকল তথ্য"),
    normalize_title("বিশ্ব পানি দিবস"),
    normalize_title("বার্ষিক প্রতিবেদন"),
    normalize_title("কমিটিসমূহ"),
    normalize_title("ফোকাল পয়েন্ট ও অন্যান্য")
}

# ---------------------------------------------------------------------------
# URL overrides for entries that share the same title but point to different
# pages. Keyed by the Bengali number prefix (e.g. "১০.৭.৪") as it appears
# in the Word table.
# ---------------------------------------------------------------------------
URL_OVERRIDES: dict[str, str] = {
    "১০.৭.৪": "https://cabinet.gov.bd/pages/static-pages/6940329735ce18e1c055ef63",
    "১০.৮.৪": "https://mowr.gov.bd/pages/static-pages/694032d235ce18e1c0562f95",
}


def fetch_most_recent_date_from_subpages(
    session: requests.Session,
    page_url: str,
    delay_seconds: float = 0.4,
) -> Optional[date]:
    """
    Open page_url, collect every link inside <td class="table-td centered">
    cells, visit each linked page, and return the most recent
    content-update-block date found among them.
    """
    response = session.get(page_url, timeout=45, verify=False)
    response.raise_for_status()
    response.encoding = "utf-8"
    soup = BeautifulSoup(response.text, "html.parser")

    # Elements that have BOTH classes "table-td" and "centered".
    table_cells = soup.select("td.table-td.centered")

    sub_urls: list[str] = []
    seen_urls: set[str] = set()
    for cell in table_cells:
        for anchor in cell.find_all("a", href=True):
            href = anchor["href"].strip()
            if not href or href == "#":
                continue
            absolute_url = urljoin(page_url, href)
            if absolute_url in seen_urls:
                continue
            seen_urls.add(absolute_url)
            sub_urls.append(absolute_url)

    print(f"    Found {len(sub_urls)} link(s) inside td.table-td.centered cells.")

    latest_date: Optional[date] = None
    for index, sub_url in enumerate(sub_urls, start=1):
        try:
            value = fetch_update_date_value(session, sub_url)
        except requests.RequestException as exc:
            print(f"    [{index}/{len(sub_urls)}] Request failed for {sub_url}: {exc}")
            time.sleep(delay_seconds)
            continue

        if value:
            print(f"    [{index}/{len(sub_urls)}] {sub_url} -> {value.isoformat()}")
            if latest_date is None or value > latest_date:
                latest_date = value
        else:
            print(f"    [{index}/{len(sub_urls)}] {sub_url} -> no date found")

        time.sleep(delay_seconds)

    return latest_date


# ---------------------------------------------------------------------------
# Word document helpers
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Main processing
# ---------------------------------------------------------------------------

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

        date_cell = table.rows[item.row_index].cells[DATE_COLUMN_INDEX]
        old_date = normalize_title(date_cell.text)

        # Check for a hardcoded URL override first (handles duplicate titles).
        override_url = URL_OVERRIDES.get(item.number_prefix)
        if override_url:
            link = SiteLink(
                title=item.title,
                href=override_url,
                absolute_url=override_url,
            )
            print(f"  -> Using URL override: {override_url}")
        else:
            link = find_matching_link(item.title, homepage_links)

        if not link:
            message = f"  -> No matching link for title: {item.title}"
            print(message)
            failures.append(message)
            if not dry_run:
                set_cell_text(date_cell, old_date)
            continue

        if not override_url:
            print(f"  -> Matched: {link.title!r}")
        print(f"  -> URL: {link.absolute_url}")

        is_special = normalize_title(item.title) in SPECIAL_MULTI_LINK_TITLES

        try:
            if is_special:
                print("  -> Special topic: scanning td.table-td.centered links for latest date...")
                latest_value = fetch_most_recent_date_from_subpages(
                    session, link.absolute_url, delay_seconds=delay_seconds
                )
                new_date = format_bengali_date(latest_value) if latest_value else None
            else:
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