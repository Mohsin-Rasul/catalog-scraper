# Catalog Scraper

Crawls a paginated product catalog and exports clean, analysis-ready data to **CSV, JSON and Excel**.
Demo target: [books.toscrape.com](https://books.toscrape.com), a public sandbox made for scraping practice.

## What it does
- Follows pagination and (optionally) every product page
- Polite crawling: rate limit, retries with exponential backoff, custom User-Agent
- Cleans fields: price as a number, star rating as 1-5, stock as an integer, category, UPC
- **Incremental runs**: pass the last `products.json` and it only collects new products
- Excel output with styled header, frozen top row, filters and sized columns
- Run summary: products, requests, errors, time

## Run it
```bash
python -m venv .venv
.venv/Scripts/pip install -r requirements.txt      # macOS/Linux: .venv/bin/pip
python scraper.py --pages 2 --details               # first 2 pages, with product details
python scraper.py --pages 0 --details               # whole catalog (1000 products)
python scraper.py --since-file output/products.json --out new   # only new items
```

## Sample result (2 pages, with details)
```
title,price_gbp,rating,in_stock,stock_count,category,upc,url
A Light in the Attic,51.77,3,True,22,Poetry,a897fe39b1053632,https://books.toscrape.com/...
Tipping the Velvet,53.74,1,True,20,Historical Fiction,90fa61229261140a,https://books.toscrape.com/...
```
`40 products, 42 requests, 0 errors, 14.2s`

## Adapting it to a client site
Only `parse_listing()` and `enrich()` hold site-specific selectors. The fetcher, pagination,
incremental mode and exporters are reused as they are. Always check the target site's terms
and robots.txt first.
