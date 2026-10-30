# Uniqlo Products Data Web Scraper

A Python web scraper that collects product data from the Uniqlo website and
writes it to a formatted Excel workbook.

## What it collects

For every product in the Women, Men, Kids and Baby categories:

- category
- product name
- SKU (product ID)
- material
- care instructions
- link to the product page

## Output

`product_list.xlsx` with one sheet per category (Women, Men, Kids, Baby). Each
sheet has styled column headers, column filters, and a clickable link column.

## How it works

`Scraper/Uniqlo_Scraper.py` crawls the site in three steps, using `requests`
and BeautifulSoup:

1. Read the main navigation to find the product category pages.
2. On each category page, collect the sub-category links, skipping lookbook
   pages and duplicates.
3. On each sub-category page, collect the product tiles. Then open every
   product page to read its material, care instructions and SKU.

The results are written to Excel with `openpyxl`.

## Setup and usage

Requires Python 3.

```bash
pip install requests beautifulsoup4 openpyxl
python Scraper/Uniqlo_Scraper.py
```

A full run opens one page per product, so it takes a while. Be considerate of
the site: don't run it more often than you need to.

## Limitations

- Built in 2020 against Uniqlo's page structure at the time. The scraper
  relies on specific CSS class names, so it will need updating whenever the
  site's markup changes.
- Check the site's terms of use before scraping, and use the data for personal
  or research purposes only.
