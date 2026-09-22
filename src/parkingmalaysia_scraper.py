#!/usr/bin/env python3
"""
ParkingInMalaysia master-list scraper.

Purpose:
- Crawl ParkingInMalaysia.com parking pages.
- Extract one row per unique building/location.
- Capture normal parking operator, valet operator, city, state, source URL,
  and a confidence/review note.
- Save the raw result as CSV for easy review/import into Excel.

Usage:
    python parkinginmalaysia_scraper.py

Dependencies:
    pip install requests beautifulsoup4 lxml

Notes:
- The site changes over time. The scraper uses generic WordPress/sitemap logic
  plus heuristics, so operator fields should still be reviewed.
- It does NOT guess an operator when evidence is weak.
"""

from __future__ import annotations

import csv
import html
import re
import time
import unicodedata
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass, field
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

BASE = "https://parkinginmalaysia.com/"
OUT_CSV = "Malaysia_Parking_Operators.csv"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (compatible; ParkingInventoryResearch/1.0; "
        "+https://parkinginmalaysia.com/)"
    )
}

# Start with operators specifically supplied by the requester.
# Add more confirmed operator names here as you encounter them.
KNOWN_OPERATORS = [
    "Times Parking",
    "SCP Parking",
    "Semasa Parking",
]

# Do NOT include "Secure Parking" here by default because many pages use
# "secure parking" descriptively rather than as a company name.

CITY_TO_STATE = {
    "alor-setar": "Kedah",
    "sungai-petani": "Kedah",
    "kulim": "Kedah",
    "brinchang": "Pahang",
    "genting": "Pahang",
    "genting-highlands": "Pahang",
    "kuantan": "Pahang",
    "ipoh": "Perak",
    "johor-bahru": "Johor",
    "klang": "Selangor",
    "petaling-jaya": "Selangor",
    "puchong": "Selangor",
    "shah-alam": "Selangor",
    "subang-jaya": "Selangor",
    "cyberjaya": "Selangor",
    "kuala-lumpur": "Kuala Lumpur",
    "kuala-terengganu": "Terengganu",
    "kota-bharu": "Kelantan",
    "kota-kinabalu": "Sabah",
    "kuching": "Sarawak",
    "malacca": "Malacca",
    "melaka": "Malacca",
    "penang": "Penang",
    "george-town": "Penang",
    "putrajaya": "Putrajaya",
    "seremban": "Negeri Sembilan",
    "kangar": "Perlis",
    "labuan": "Labuan",
}

STATE_SLUGS = {
    "johor": "Johor",
    "kedah": "Kedah",
    "kelantan": "Kelantan",
    "malacca": "Malacca",
    "melaka": "Malacca",
    "negeri-sembilan": "Negeri Sembilan",
    "pahang": "Pahang",
    "penang": "Penang",
    "perak": "Perak",
    "perlis": "Perlis",
    "sabah": "Sabah",
    "sarawak": "Sarawak",
    "selangor": "Selangor",
    "terengganu": "Terengganu",
    "kuala-lumpur": "Kuala Lumpur",
    "labuan": "Labuan",
    "putrajaya": "Putrajaya",
}

EXCLUDE_TITLE_PHRASES = [
    "parking rate board",
    "parking rate boards",
    "parking payment terminal",
    "parking payment terminals",
    "floor directory",
    "mall directory",
    "f&b directory",
    "image gallery",
    "theme park map",
    "station line map",
]

INCLUDE_TITLE_PATTERNS = [
    r"\bparking info\b",
    r"\bvisitor parking info\b",
    r"\boutdoor parking info\b",
    r"\bparking rate\b",
    r"\bparking rates\b",
    r"\bparking fee\b",
]

BUILDING_SUFFIXES = [
    r"\s+visitor parking info\s*$",
    r"\s+outdoor parking info\s*$",
    r"\s+parking info\s*$",
    r"\s+parking rates?\s*$",
    r"\s+parking fee\s*$",
]


@dataclass
class Row:
    building: str
    normal_ops: set[str] = field(default_factory=set)
    valet_ops: set[str] = field(default_factory=set)
    city: str = ""
    state: str = ""
    sources: set[str] = field(default_factory=set)
    notes: set[str] = field(default_factory=set)


session = requests.Session()
session.headers.update(HEADERS)


def get(url: str, tries: int = 3, timeout: int = 30) -> requests.Response:
    last = None
    for i in range(tries):
        try:
            r = session.get(url, timeout=timeout)
            if r.status_code == 200:
                return r
            last = RuntimeError(f"{r.status_code} for {url}")
        except Exception as e:
            last = e
        time.sleep(1.5 * (i + 1))
    raise RuntimeError(f"Failed to fetch {url}: {last}")


def normalize_space(s: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(s or "")).strip()


def slugify_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = s.lower().replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def is_candidate_title(title: str) -> bool:
    low = title.lower()
    if any(x in low for x in EXCLUDE_TITLE_PHRASES):
        return False
    return any(re.search(p, low) for p in INCLUDE_TITLE_PATTERNS)


def clean_building_name(title: str) -> str:
    title = normalize_space(title)
    title = title.replace("’", "'")
    for pat in BUILDING_SUFFIXES:
        title = re.sub(pat, "", title, flags=re.I)
    return title.strip(" -–—|")


def parse_xml_locs(xml_text: str) -> list[str]:
    root = ET.fromstring(xml_text)
    locs = []
    for el in root.iter():
        if el.tag.endswith("loc") and el.text:
            locs.append(el.text.strip())
    return locs


def discover_sitemap_urls() -> list[str]:
    candidates = [
        urljoin(BASE, "wp-sitemap.xml"),
        urljoin(BASE, "sitemap_index.xml"),
        urljoin(BASE, "sitemap.xml"),
    ]
    for u in candidates:
        try:
            r = get(u)
            locs = parse_xml_locs(r.text)
            if locs:
                print(f"Using sitemap index: {u}")
                # Sitemap index -> child sitemaps; ordinary sitemap -> page URLs.
                if any(x.endswith(".xml") for x in locs):
                    return locs
                return [u]
        except Exception:
            pass
    raise RuntimeError("Could not locate a sitemap index.")


def discover_page_urls() -> list[str]:
    sitemap_urls = discover_sitemap_urls()
    urls = set()

    for sm in sitemap_urls:
        # Prefer post-like sitemaps but still allow generic sitemaps.
        if "image" in sm.lower():
            continue
        try:
            locs = parse_xml_locs(get(sm).text)
        except Exception as e:
            print(f"Skipping sitemap {sm}: {e}")
            continue

        for u in locs:
            if not u.startswith(BASE):
                continue
            p = urlparse(u).path.strip("/")
            if not p:
                continue
            # Exclude obvious archive/taxonomy/system URLs.
            if any(
                p.startswith(prefix)
                for prefix in (
                    "category/",
                    "tag/",
                    "author/",
                    "wp-",
                    "feed",
                )
            ):
                continue
            urls.add(u)

    print(f"Discovered {len(urls)} page URLs from sitemap(s).")
    return sorted(urls)


def extract_title_and_soup(url: str):
    r = get(url)
    soup = BeautifulSoup(r.text, "lxml")
    h1 = soup.find("h1")
    if h1:
        title = normalize_space(h1.get_text(" ", strip=True))
    else:
        title = normalize_space(soup.title.get_text(" ", strip=True) if soup.title else "")
        title = re.sub(r"\s*[|–—-]\s*Parking In Malaysia\s*$", "", title, flags=re.I)
    return title, soup


def extract_location(soup: BeautifulSoup) -> tuple[str, str]:
    city = ""
    state = ""

    # WordPress category links are the cleanest source when present.
    cats = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "/category/location/" in href:
            path = urlparse(href).path.strip("/").split("/")
            try:
                i = path.index("location")
                tail = path[i + 1 :]
                cats.extend(tail)
            except ValueError:
                pass

    # Also inspect breadcrumbs/text links that use recognizable slugs.
    ordered = []
    for x in cats:
        if x not in ordered:
            ordered.append(x)

    for slug in ordered:
        if slug in STATE_SLUGS:
            state = STATE_SLUGS[slug]
        if slug in CITY_TO_STATE:
            city = slug.replace("-", " ").title()
            state = state or CITY_TO_STATE[slug]

    # If category gave only a state/city hierarchy, choose the deepest non-state as city.
    for slug in reversed(ordered):
        if slug not in STATE_SLUGS and slug not in {"location"}:
            city = city or slug.replace("-", " ").title()
            break

    return city, state


def all_page_text(soup: BeautifulSoup) -> str:
    parts = [soup.get_text("\n", strip=True)]
    # Image alt/title/captions often contain the operator name.
    for img in soup.find_all("img"):
        for attr in ("alt", "title"):
            if img.get(attr):
                parts.append(str(img.get(attr)))
    for tag in soup.find_all(attrs={"title": True}):
        parts.append(str(tag.get("title")))
    return normalize_space("\n".join(parts))


def extract_operator_candidates(text: str) -> tuple[set[str], list[str]]:
    operators = set()
    evidence = []

    # Explicit wording: managed/operated by / parking operator / image credit.
    patterns = [
        r"(?:parking|car\s*park)\s+(?:is\s+)?(?:managed|operated)\s+by\s+([A-Z][A-Za-z0-9&().,'’\- ]{2,60})",
        r"(?:parking|car\s*park)\s+operator\s*[:\-]\s*([A-Z][A-Za-z0-9&().,'’\- ]{2,60})",
        r"(?:image(?:\s+credit)?|photo(?:\s+credit)?)\s*[:\-]\s*([A-Z][A-Za-z0-9&().,'’\- ]{2,60}Parking[A-Za-z0-9&().,'’\- ]*)",
    ]
    for pat in patterns:
        for m in re.finditer(pat, text, flags=re.I):
            cand = normalize_space(m.group(1))
            cand = re.split(r"\s{2,}|(?:More Info|Back|Last Updated|Rates)", cand, maxsplit=1)[0].strip(" .,:;-")
            if 2 < len(cand) <= 80:
                operators.add(cand)
                evidence.append(f"Explicit operator evidence: {cand}")

    for op in KNOWN_OPERATORS:
        if re.search(rf"\b{re.escape(op)}\b", text, flags=re.I):
            operators.add(op)
            evidence.append(f"Known operator name found: {op}")

    return operators, evidence


def extract_valet(text: str, normal_ops: set[str]) -> tuple[set[str], bool, list[str]]:
    if not re.search(r"\bvalet\b", text, flags=re.I):
        return set(), False, []

    valet_ops = set()
    notes = ["Valet service mentioned on page."]

    explicit = [
        r"valet\s+(?:parking\s+)?(?:is\s+)?(?:managed|operated)\s+by\s+([A-Z][A-Za-z0-9&().,'’\- ]{2,60})",
        r"valet\s+by\s+([A-Z][A-Za-z0-9&().,'’\- ]{2,60})",
        r"valet\s+operator\s*[:\-]\s*([A-Z][A-Za-z0-9&().,'’\- ]{2,60})",
    ]
    for pat in explicit:
        for m in re.finditer(pat, text, flags=re.I):
            cand = normalize_space(m.group(1))
            cand = re.split(r"\s{2,}|(?:Rate|RM|Last Updated|Directions)", cand, maxsplit=1)[0].strip(" .,:;-")
            if cand:
                valet_ops.add(cand)
                notes.append(f"Explicit valet operator evidence: {cand}")

    # Only associate a known operator with valet if it appears near the word valet.
    for m in re.finditer(r"\bvalet\b", text, flags=re.I):
        window = text[max(0, m.start()-220): m.end()+220]
        for op in KNOWN_OPERATORS:
            if re.search(rf"\b{re.escape(op)}\b", window, flags=re.I):
                valet_ops.add(op)
                notes.append(f"{op} appears near valet text.")

    # If the only detected name is the same as normal operator, keep it in valet only
    # when valet-specific evidence exists. Otherwise leave valet operator unstated.
    return valet_ops, True, notes


def confidence_note(normal_ops: set[str], evidence: list[str], valet_exists: bool, valet_ops: set[str]) -> str:
    bits = []
    if evidence:
        bits.append("Confirmed/Probable: operator evidence found")
    elif normal_ops:
        bits.append("Probable: operator name found")
    else:
        bits.append("Needs review: normal operator not stated/detected")

    if valet_exists and not valet_ops:
        bits.append("Valet offered; valet operator not stated/detected")
    return "; ".join(bits)


def dedupe_key(building: str, city: str, state: str) -> str:
    return " | ".join([slugify_key(building), slugify_key(city), slugify_key(state)])


def main():
    page_urls = discover_page_urls()
    rows: dict[str, Row] = {}
    processed = 0
    candidates = 0

    for idx, url in enumerate(page_urls, 1):
        try:
            title, soup = extract_title_and_soup(url)
        except Exception as e:
            print(f"[{idx}/{len(page_urls)}] fetch failed: {url} :: {e}")
            continue

        if not is_candidate_title(title):
            continue

        candidates += 1
        building = clean_building_name(title)
        city, state = extract_location(soup)
        text = all_page_text(soup)

        normal_ops, evidence = extract_operator_candidates(text)
        valet_ops, valet_exists, valet_notes = extract_valet(text, normal_ops)
        note = confidence_note(normal_ops, evidence, valet_exists, valet_ops)

        key = dedupe_key(building, city, state)
        if key not in rows:
            rows[key] = Row(building=building, city=city, state=state)

        row = rows[key]
        row.normal_ops.update(normal_ops)
        row.valet_ops.update(valet_ops)
        row.sources.add(url)
        row.notes.add(note)
        row.notes.update(valet_notes)

        # If later duplicate has a better location, retain it.
        row.city = row.city or city
        row.state = row.state or state

        processed += 1
        if processed % 25 == 0:
            print(f"Processed {processed} parking pages ({candidates} candidates seen).")

        # Be courteous to the site.
        time.sleep(0.15)

    output = sorted(rows.values(), key=lambda r: (r.state, r.city, r.building.lower()))

    with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow([
            "Building Name",
            "Parking Operator (Normal)",
            "Parking Operator (Valet)",
            "City",
            "State / Federal Territory",
            "Source URL",
            "Notes / Confidence",
        ])
        for r in output:
            w.writerow([
                r.building,
                "; ".join(sorted(r.normal_ops)) if r.normal_ops else "Not stated",
                "; ".join(sorted(r.valet_ops)) if r.valet_ops else ("Not stated" if any("Valet" in n for n in r.notes) else ""),
                r.city,
                r.state,
                "; ".join(sorted(r.sources)),
                "; ".join(sorted(r.notes)),
            ])

    print()
    print(f"Done. Unique building/location rows: {len(output)}")
    print(f"Saved: {OUT_CSV}")
    print("Open the CSV in Excel, then Save As .xlsx after review.")


if __name__ == "__main__":
    main()
