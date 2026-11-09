"""Catalog scraper: crawls a paginated product catalog and exports clean data.

Demo target: https://books.toscrape.com (a public sandbox built for scraping practice).

Features
- Follows pagination and product detail pages
- Polite: rate limit, retries with exponential backoff, custom User-Agent
- Cleans data: prices as numbers, ratings as 1-5, stock as integer
- Exports CSV, JSON and a formatted Excel file
- Incremental mode: --since-file skips products already scraped
- Run summary with counts, errors and timing

Usage
    python scraper.py --pages 3 --out output
    python scraper.py --pages 0 --details          # all pages, with detail pages
"""
from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import time
from dataclasses import asdict, dataclass, fields
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = "https://books.toscrape.com/"
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
log = logging.getLogger("scraper")


@dataclass
class Product:
    title: str
    price_gbp: float
    rating: int
    in_stock: bool
    stock_count: int | None
    category: str | None
    upc: str | None
    url: str


class Fetcher:
    """HTTP session with rate limiting and retry with backoff."""

    def __init__(self, delay: float = 0.5, retries: int = 3, timeout: float = 15):
        self.s = requests.Session()
        self.s.headers["User-Agent"] = "CatalogScraper/1.0 (+portfolio demo)"
        self.delay, self.retries, self.timeout = delay, retries, timeout
        self._last = 0.0
        self.requests = 0
        self.errors = 0

    def get(self, url: str) -> BeautifulSoup | None:
        for attempt in range(1, self.retries + 1):
            wait = self.delay - (time.monotonic() - self._last)
            if wait > 0:
                time.sleep(wait)
            self._last = time.monotonic()
            try:
                r = self.s.get(url, timeout=self.timeout)
                self.requests += 1
                if r.status_code == 404:
                    return None
                r.raise_for_status()
                r.encoding = "utf-8"
                return BeautifulSoup(r.text, "html.parser")
            except requests.RequestException as e:
                backoff = 2 ** attempt
                log.warning("attempt %d/%d failed for %s: %s (retry in %ds)",
                            attempt, self.retries, url, e, backoff)
                time.sleep(backoff)
        self.errors += 1
        log.error("giving up on %s", url)
        return None


def parse_listing(soup: BeautifulSoup, page_url: str) -> list[Product]:
    items = []
    for art in soup.select("article.product_pod"):
        a = art.select_one("h3 a")
        price = art.select_one(".price_color").get_text(strip=True)
        rating_cls = art.select_one(".star-rating")["class"]
        avail = art.select_one(".availability").get_text(strip=True)
        items.append(Product(
            title=a["title"].strip(),
            price_gbp=float(re.sub(r"[^\d.]", "", price)),
            rating=next((RATINGS[c] for c in rating_cls if c in RATINGS), 0),
            in_stock="In stock" in avail,
            stock_count=None,
            category=None,
            upc=None,
            url=urljoin(page_url, a["href"]),
        ))
    return items


def enrich(f: Fetcher, p: Product) -> None:
    soup = f.get(p.url)
    if not soup:
        return
    crumbs = soup.select("ul.breadcrumb li a")
    if len(crumbs) >= 3:
        p.category = crumbs[2].get_text(strip=True)
    table = {th.get_text(strip=True): td.get_text(strip=True)
             for th, td in zip(soup.select("table th"), soup.select("table td"))}
    p.upc = table.get("UPC")
    m = re.search(r"\((\d+) available\)", table.get("Availability", ""))
    p.stock_count = int(m.group(1)) if m else None


def crawl(pages: int, details: bool, seen: set[str], f: Fetcher) -> list[Product]:
    out: list[Product] = []
    url, n = urljoin(BASE, "catalogue/page-1.html"), 0
    while url and (pages == 0 or n < pages):
        soup = f.get(url)
        if not soup:
            break
        n += 1
        batch = [p for p in parse_listing(soup, url) if p.url not in seen]
        if details:
            for p in batch:
                enrich(f, p)
        out.extend(batch)
        log.info("page %d: %d new products (total %d)", n, len(batch), len(out))
        nxt = soup.select_one("li.next a")
        url = urljoin(url, nxt["href"]) if nxt else None
    return out


def export(items: list[Product], out: Path) -> None:
    out.mkdir(parents=True, exist_ok=True)
    cols = [x.name for x in fields(Product)]
    rows = [asdict(p) for p in items]

    with open(out / "products.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)

    (out / "products.json").write_text(json.dumps(rows, indent=2, ensure_ascii=False),
                                       encoding="utf-8")

    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    ws.append(cols)
    for r in rows:
        ws.append([r[c] for c in cols])
    head = PatternFill("solid", fgColor="1F3A5F")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head
    for i, c in enumerate(cols, 1):
        width = max(len(c), *(len(str(r[c] or "")) for r in rows)) if rows else len(c)
        ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(out / "products.xlsx")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--pages", type=int, default=3, help="listing pages to crawl (0 = all)")
    ap.add_argument("--details", action="store_true", help="also open each product page")
    ap.add_argument("--delay", type=float, default=0.5, help="seconds between requests")
    ap.add_argument("--out", type=Path, default=Path("output"))
    ap.add_argument("--since-file", type=Path, help="previous products.json; skip known URLs")
    a = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    seen: set[str] = set()
    if a.since_file and a.since_file.exists():
        seen = {r["url"] for r in json.loads(a.since_file.read_text(encoding="utf-8"))}
        log.info("incremental mode: %d known products skipped", len(seen))

    f = Fetcher(delay=a.delay)
    t0 = time.monotonic()
    items = crawl(a.pages, a.details, seen, f)
    export(items, a.out)
    log.info("done: %d products, %d requests, %d errors, %.1fs -> %s",
             len(items), f.requests, f.errors, time.monotonic() - t0, a.out.resolve())


if __name__ == "__main__":
    main()
