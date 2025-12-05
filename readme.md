# Excel 404 Link Scanner

This tool scans an **Excel (.xlsx)** file, checks URLs in selected columns, and **highlights cells in red** when a link returns **HTTP 404**.

**Authorship:** I built the original core script (Excel reading + link scanning + red highlighting). Then I asked **ChatGPT** to add the **GUI** (file picker + scan button + column input).

<a href="https://www.trustrepo.dev/Lorkydey/url-404-finder-in-excel" target="_blank" rel="noreferrer"><img src="https://www.trustrepo.dev/api/badge/Lorkydey/url-404-finder-in-excel?mode=ai" alt="TrustRepo badge (AI score) for Lorkydey/url-404-finder-in-excel" style="max-width:100%;height:auto;" loading="lazy" /></a>

## Features
- Select an `.xlsx` file
- Choose columns to scan (e.g. `17,18,19` or `Q,R,S`)
- Optional sheet name (empty = active sheet)
- Fast concurrent HTTP checks
- Saves `*{name}_checked.xlsx`

## Install
```bash
pip install openpyxl aiohttp
```

## run
```bash
python check.py
```

## Important
it only support xlsx not xls
