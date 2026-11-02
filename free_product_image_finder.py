#!/usr/bin/env python3
"""
Free product-image URL finder for the supplied product catalog workbook.

What it does:
- Searches DuckDuckGo Images for each product title (no OpenAI API key/credits).
- Checks candidate URLs to see whether they return an actual image.
- Writes the first usable direct image URL into column J.
- Preserves the workbook's existing IMAGE formula in column A.

Notes:
- Search engines may rate-limit or block automated requests.
- Product matches are not guaranteed; review images before using commercially.
- Excel's IMAGE() formula generally requires a supported, current Excel version
  and a publicly accessible image URL.
"""
import argparse
import os
import sys
import time
from io import BytesIO
from urllib.parse import urlparse

import requests
from ddgs import DDGS
from openpyxl import load_workbook
from PIL import Image

DEFAULT_FILE = "PRODUCT_CATALOG_all_10313_image_search_links.xlsx"
SHEET_NAME = "Product Catalog"
TITLE_COL = 2       # B
DIRECT_URL_COL = 10 # J
IMAGE_FORMULA_COL = 1 # A

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/130.0.0.0 Safari/537.36"
    )
}

def is_valid_image_url(url, timeout=12):
    """Return True if URL responds with image bytes that Pillow can open."""
    if not url or not url.lower().startswith(("http://", "https://")):
        return False
    try:
        response = requests.get(
            url, headers=HEADERS, timeout=timeout, allow_redirects=True
        )
        if response.status_code != 200 or len(response.content) > 12 * 1024 * 1024:
            return False
        content_type = response.headers.get("Content-Type", "").lower()
        if "image/" not in content_type:
            # Some hosts mislabel image files, so verify the bytes as well.
            pass
        if not response.content:
            return False
        try:
            im = Image.open(BytesIO(response.content))
            im.verify()
            return im.format in {"JPEG", "PNG", "GIF", "WEBP", "BMP", "TIFF"}
        except Exception:
            return False
    except requests.RequestException:
        return False

def find_image_url(title, max_results=12):
    # Product title is the search query; results contain direct image URLs.
    try:
        with DDGS() as ddgs:
            results = ddgs.images(
                title,
                region="wt-wt",
                safesearch="moderate",
                max_results=max_results,
            )
            for result in results:
                image_url = result.get("image")
                if image_url and is_valid_image_url(image_url):
                    return image_url
    except Exception as exc:
        print(f"    Search warning: {exc}")
    return None

def main():
    parser = argparse.ArgumentParser(description="Find free product image URLs for Excel.")
    parser.add_argument("--file", default=DEFAULT_FILE, help="Workbook filename/path")
    parser.add_argument("--limit", type=int, default=10313, help="Maximum number of rows to process (default: 5)")
    parser.add_argument("--start-row", type=int, default=2, help="First product row (default: 2)")
    parser.add_argument("--delay", type=float, default=1.5, help="Seconds between searches (default: 1.5)")
    args = parser.parse_args()

    file_path = os.path.abspath(args.file)
    if not os.path.isfile(file_path):
        print(f"ERROR: Workbook not found: {file_path}")
        print("Put this script in the same folder as the workbook, or pass --file with its full path.")
        sys.exit(1)

    wb = load_workbook(file_path)
    if SHEET_NAME not in wb.sheetnames:
        print(f"ERROR: Expected sheet '{SHEET_NAME}' was not found.")
        sys.exit(1)
    ws = wb[SHEET_NAME]

    last_row = ws.max_row
    end_row = min(last_row, args.start_row + max(args.limit, 0) - 1)
    found = 0
    skipped = 0
    failed = 0

    print(f"Workbook: {file_path}")
    print(f"Processing rows {args.start_row}-{end_row} (limit {args.limit})")

    for row in range(args.start_row, end_row + 1):
        title = str(ws.cell(row, TITLE_COL).value or "").strip()
        existing = str(ws.cell(row, DIRECT_URL_COL).value or "").strip()

        if not title:
            skipped += 1
            continue
        if existing.startswith(("https://", "http://")) and is_valid_image_url(existing):
            print(f"Row {row}: already has a valid image URL; keeping it.")
            skipped += 1
            continue

        print(f"Row {row}: searching for {title!r} ...")
        url = find_image_url(title)
        if url:
            ws.cell(row, DIRECT_URL_COL).value = url
            found += 1
            print(f"    FOUND: {url}")
        else:
            failed += 1
            print("    No verified image URL found.")
        time.sleep(max(0, args.delay))

    root, ext = os.path.splitext(file_path)
    out_path = f"{root}_with_image_urls{ext or '.xlsx'}"
    wb.save(out_path)
    print("\nFinished.")
    print(f"Image URLs found: {found}; not found: {failed}; skipped: {skipped}")
    print(f"Saved a NEW workbook; original unchanged:\n{out_path}")
    print("Open the new workbook in Excel and allow time for IMAGE() formulas to load.")

if __name__ == "__main__":
    main()
