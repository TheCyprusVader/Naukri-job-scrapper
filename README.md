# Naukri Job Scraper

A personal automation tool that scrapes [Naukri.com](https://www.naukri.com) daily for fresh, junior-appropriate Data Analyst / Business Analyst / MIS roles across India, and saves them into a weekly-organized spreadsheet — so I don't have to manually search across dozens of pages every day.

## What it does

- Searches multiple job titles ("data analyst", "business analyst", "MIS executive") across 10 pages each, nationwide
- Filters results to only:
  - Jobs posted in the **last 24 hours**
  - Jobs with an experience range realistically fitting **0–3 years** (with an upper-bound cap, so senior-leaning postings like "2–7 Yrs" that technically overlap don't slip through)
- Deduplicates against everything already saved that week, so the same job never appears twice
- Saves results into `naukri_jobs_weekly.xlsx`, with one sheet per weekday (Monday–Friday). Running it on **Saturday** merges the week into a single deduped summary sheet instead of scraping again.

## How it works

1. **Fetching** — Naukri renders job listings via JavaScript after the initial page load (confirmed by inspecting page source before building this), so a simple HTTP request wouldn't return usable data. [Playwright](https://playwright.dev/) drives a real Chromium browser to load the page and wait for the listings to render before grabbing the HTML.
   - Note: this must run **non-headless** (`headless=False`). Naukri's bot detection silently returns zero results to headless browser sessions — this was discovered by testing, not documented anywhere, so it's called out directly in the code.
2. **Parsing** — [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/) extracts structured fields (title, company, experience, location, posted date, skills, description, link) from each job card using its CSS classes.
3. **Filtering** — title keyword match → experience range fit → freshness check, applied in that order.
4. **Storage** — [openpyxl](https://openpyxl.readthedocs.io/) writes to a multi-sheet `.xlsx` workbook (plain CSV can't hold multiple sheets), tracking which jobs have already been saved across the whole week to avoid duplicates.

## Setup

```bash
pip install playwright beautifulsoup4 openpyxl
playwright install chromium
```

## Usage

```bash
python naukri_scrapping.py
```

For hands-off daily runs, this is set up via Windows Task Scheduler with a **"When I log on"** trigger (rather than a fixed time), since login time varies day to day.

## Notes

- This project was built with AI pair-programming assistance (Claude), used to learn web scraping fundamentals (JS-rendered fetching, HTML parsing, filtering logic) step by step while building something genuinely useful for my own job search.
- For personal, non-commercial use.
