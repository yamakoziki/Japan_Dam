# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**JapDam05** is a Japanese dam geology database system. It fetches geological data from a public government API (GeoNAVI by AIST), stores results in Excel, performs statistical analysis, and provides a web-based search interface.

## Setup

```bash
pip3 install openpyxl
```

No other third-party dependencies — `urllib` and `ssl` are stdlib.

## Running Scripts

**Data fetch (ETL) — `作業.py`:**
```bash
# Fetch geological data from GeoNAVI API and write to Excel
python3 作業.py --input data/zendam31_fixed.xlsx --output output.xlsx

# Process specific rows only
python3 作業.py --input data/zendam31_fixed.xlsx --output output.xlsx --rows 3-100

# Preview without writing
python3 作業.py --input data/zendam31_fixed.xlsx --output output.xlsx --dry-run

# Overwrite already-filled cells
python3 作業.py --input data/zendam31_fixed.xlsx --output output.xlsx --overwrite

# Retry rows that failed in a prior run (from CSV log)
python3 作業.py --input data/zendam31_fixed.xlsx --output output.xlsx --retry-log 作業ログ.csv
```

**Statistical analysis — `分析2.py`:**
```bash
python3 分析2.py --input data/zendam31_fixed.xlsx --output 分析報告書.xlsx
```

**Web interface:** open `dam_search.html` directly in a browser, then load an Excel file via the UI.

## Architecture

### Data Flow
```
Excel (dam locations) → 作業.py → GeoNAVI API → Excel (geology results)
                                                      ↓
                                              分析2.py → Excel (analysis sheets)
                                                      ↓
                                              dam_search.html (browser UI)
```

### Excel Data Structure

**Main sheet: `全国ダム地質DB`**
- Row 3+: dam records
- Col 20–21 (T–U): latitude/longitude used for API calls
- Cols 23, 35, 47, 59, 71 (W, AI, AU, BG, BS): legend IDs for 5 geological layers
- Each legend ID column is followed by: symbol, geo_surface, geo_era, geo_rock, formation age, group, lithology, risk

**Glossary sheet:** master reference for all geological units (row 4+), keyed by `id` and `symbol`.

### API: GeoNAVI (AIST)
- Endpoint: `https://gbank.gsj.jp/seamless/v2/api/1.2/legend.json`
- Public, no auth required
- Rate limit: 0.5s between calls; timeout 15s; up to 3 retries with backoff
- Special SSL context handling for macOS (custom cert bundle)

### Geological Layer Assignment Logic (`作業.py`)
1. Call API at exact dam coordinates → assign to geological era layer
2. If response is null (no geology): search 8 cardinal directions at 500m → 8000m expanding radius
3. If only "Late Quaternary (Q-H)" found: search surroundings for pre-Q-H geology

**Era → layer mapping:**
- Layer 1 (col W): Pre-Tertiary
- Layer 2 (col AI): Tertiary (N)
- Layer 3 (col AU): mid-Pleistocene (Q-old)
- Layer 4 (col BG): Late Pleistocene–Holocene (Q-H)
- Layer 5 (col BS): supplementary

### Analysis Sheets (`分析2.py`)
Generates 7 sheets: symbol hierarchy analysis, bearing capacity × permeability matrix, symbol clustering, 2-item combinations, Hokkaido 60 dams, national 100 dams selection, and coverage comparison.

**Key metrics:**
- Bearing capacity score: 1–5 (low → high)
- Permeability score: 1–4 (low → high)
- Risk = (5 − bearing_capacity) + permeability

### Logging
All runs emit a CSV log tracking per-row status (success/error/skipped), API search radius used, and error messages. Use `--retry-log` to re-process only failed rows from a prior log.
