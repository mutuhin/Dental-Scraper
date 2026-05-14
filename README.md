# Dental Practice Scraper

Automated pipeline for scraping US dental practice websites and building a structured Excel dataset.
Covers up to **6,000 practices** (60 batches × 100 rows each) across all data fields.

---

## What It Captures

| Column | Field | How |
|--------|-------|-----|
| 3 | Doctor Name | Team page headings — per doctor |
| 10 | # of Hygienists | Credential title search (RDH / BSDH / RDHAP / Licensed Dental Hygienist) |
| 23 | CEREC (Same Day Crowns) | Keyword match across all pages |
| 24 | CBCT (3D Imaging) | Keyword match |
| 25 | Lasers | Keyword match |
| 26 | AI | Keyword match |
| 27 | Intraoral Scanners | Keyword match |
| 28 | Invisalign (Mentions) | Per-page count |
| 30–38 | Services (Veneers, Implants, Whitening, Sedation, etc.) | Per-page mention count |
| 40 | Associations / Memberships | Per-doctor bio (ADA, AGD, AACD, AAPD, etc.) |
| 41 | Doctor Specialty | Per-doctor bio (Cosmetic, Pediatric, Implants, etc.) |
| 42–45 | Google / Yelp Ratings | Google Places API + Yelp search |
| 46 | Testimonials | CSS class / schema.org / blockquote counting |

---

## Scripts

| File | Purpose |
|------|---------|
| `dental_scraper.py` | Core library — crawling, parsing, field extraction, doctor detection |
| `run_batch.py` | Entry point for the scraper workflow — accepts batch number and optional start row |
| `reprocess.py` | Post-processing — doctor name deduplication, data cleaning |
| `google_ratings_v12.py` | Google ratings engine — Places API + Playwright fallback |
| `google_ratings_workflow.py` | Wrapper that runs ratings on a deduped batch file |
| `refresh_tech_services.py` | Re-extracts tech / services / testimonials from saved page cache without re-crawling |

---

## Required Secrets

Go to **Settings → Secrets and variables → Actions** and add:

| Secret | Required | Description |
|--------|----------|-------------|
| `NORDVPN_SERVICE_USER` | Optional | NordVPN service username (not your login email). Enables US IP rotation to avoid rate limiting |
| `NORDVPN_SERVICE_PASS` | Optional | NordVPN service password |
| `GOOGLE_PLACES_API_KEY` | For ratings | Google Places API key used by the ratings workflow |

> Without NordVPN secrets the workflows skip VPN and run on GitHub's default IP.
> Without the Google Places key the ratings step will fall back to search scraping only.

---

## The 4 Workflows

### 1. Dental Scraper
**File:** `.github/workflows/scraper.yml`

Scrapes a full batch of 100 practices — crawls each website, extracts all fields, writes one row per doctor to an Excel file.

**Inputs:**
- `batch` — batch number (1 = rows 1–100, 2 = rows 101–200, up to 60)
- `start_row` *(optional)* — resume from a specific row within the batch (e.g. `45` to skip rows already done)

**Outputs (GitHub Artifacts):**
- `batch-N` — raw scraped xlsx
- `batch-N-deduped` — same file after doctor name deduplication
- `batch-N-cache` — compressed page cache (HTML of every visited page, kept 7 days for resume/refresh)

**Resume support:** If the run is interrupted, re-run with the same batch number and a `start_row`. The workflow automatically restores the previous page cache so already-scraped practices are not re-crawled.

**Timeout:** 5.5 hours (330 min) — enough for ~100 practices at ~3 min each.

---

### 2. Dedup + Google Ratings
**File:** `.github/workflows/collect_and_rate.yml`

Downloads the completed `batch-N` artifact, deduplicates doctor names, then runs Google ratings on the clean file.

**Inputs:**
- `batch` — batch number matching a previously completed Dental Scraper run
- `skip_ratings` *(optional, boolean)* — set `true` to run dedup only, skip Google ratings

**Outputs (GitHub Artifacts):**
- `batch-N-deduped` — deduplicated xlsx
- `google-ratings-batch-N` — xlsx with Google/Yelp ratings added

**Note:** Google ratings alone can take 2–4 hours per batch. After this workflow completes, run **Refresh Tech & Services** separately for the same batch number.

**Timeout:** 6 hours (360 min).

---

### 3. Dedup Doctor Names Only
**File:** `.github/workflows/dedup_only.yml`

Downloads an existing batch artifact and runs doctor deduplication only. Use this when the scraper finished successfully but the dedup step failed, or when you want to re-dedup without re-scraping.

**Inputs:**
- `batch` — batch number

**Outputs (GitHub Artifacts):**
- `batch-N-deduped` — deduplicated xlsx

**Timeout:** 10 minutes.

---

### 4. Refresh Tech & Services from Cache
**File:** `.github/workflows/refresh_tech_services.yml`

Re-extracts Technology, Services, Testimonials, Hygienists, and Doctor Specialty/Associations from the saved page cache — no live crawling. Also does a targeted live crawl for any practices that are still missing data. Produces a comparison report showing every changed cell.

**Inputs:**
- `batch` — batch number

**Outputs (GitHub Artifacts):**
- `batch-N-refreshed` — patched xlsx with updated fields
- `batch-N-comparison` — Excel before/after report (changed cells highlighted in green/yellow)

**Input priority:** Automatically picks the best available artifact in this order:
1. `google-ratings-batch-N` (most complete)
2. `batch-N-deduped`
3. `batch-N` (raw)

**Timeout:** 6 hours (360 min).

---

## Recommended Run Order Per Batch

```
1. Dental Scraper (batch N)
         ↓
2. Dedup + Google Ratings (batch N)
         ↓
3. Refresh Tech & Services from Cache (batch N)
         ↓
   Download: batch-N-refreshed.xlsx  ← final output
```

> Step 3 should always be run after step 2 so it picks up the ratings-enriched file as its input.

---

## Downloading Output Files

1. Go to **Actions** in the repository
2. Click the completed workflow run
3. Scroll to **Artifacts** at the bottom of the run page
4. Download the relevant artifact:
   - Final output: `batch-N-refreshed`
   - Ratings only: `google-ratings-batch-N`
   - Change log: `batch-N-comparison`

Artifacts expire after **30 days** (page cache after **7 days**).

---

## Batch Reference

| Batch | Rows | Batch | Rows |
|-------|------|-------|------|
| 1 | 1–100 | 11 | 1001–1100 |
| 2 | 101–200 | 12 | 1101–1200 |
| 3 | 201–300 | … | … |
| 4 | 301–400 | 30 | 2901–3000 |
| 5 | 401–500 | … | … |
| … | … | 60 | 5901–6000 |

---

## Known Limitations

- **Cloudflare-protected sites** return HTTP 403/202 and cannot be scraped automatically. These are skipped and show blank fields.
- **SPA / portal sites** (Kaiser Permanente, Healthgrades, Zocdoc, etc.) are skipped to prevent workflow freezes — they require a real login browser session.
- **Dynamically loaded reviews** (Google Maps widget embedded in page, JS-rendered testimonials) may not be counted in the Testimonials field.
- **Associations** are matched against a fixed list of 21 known US dental associations. Local or regional memberships not in the list will be missed.
- **Hygienists** require a credential title (RDH, BSDH, RDHAP, or "Dental Hygienist") to appear on the site. Staff listed only by first name with no title cannot be detected.
