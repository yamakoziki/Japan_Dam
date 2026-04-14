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

**Data fetch (ETL) — `firstset/作業.py`:**
```bash
# Fetch geological data from GeoNAVI API and write to Excel
python3 firstset/作業.py --input firstset/zendam31_nodata.xlsx --output output.xlsx

# Process specific rows only
python3 firstset/作業.py --input firstset/zendam31_nodata.xlsx --output output.xlsx --rows 3-100

# Preview without writing
python3 firstset/作業.py --input firstset/zendam31_nodata.xlsx --output output.xlsx --dry-run

# Overwrite already-filled cells
python3 firstset/作業.py --input firstset/zendam31_nodata.xlsx --output output.xlsx --overwrite

# Retry rows that failed in a prior run (from CSV log)
python3 firstset/作業.py --input firstset/zendam31_nodata.xlsx --output output.xlsx --retry-log firstset/作業ログ.csv
```

**Statistical analysis — `分析2.py`:**
```bash
python3 分析2.py --input data/zendam31_fixed.xlsx --output 分析報告書.xlsx
```

**Web interface:** open `index.html` (or the identical `dam_search.html`) directly in a browser, then load an Excel file via the UI.

## Architecture

### Data Flow
```
Excel (dam locations) → firstset/作業.py → GeoNAVI API → Excel (geology results)
                                                               ↓
                                                       分析2.py → Excel (analysis sheets)
                                                               ↓
                                                       index.html (browser UI)
```

### Directory Layout
- `firstset/` — preserved artifacts from the first completed run: `作業.py` (canonical ETL script), `zendam31_nodata.xlsx` (input), `出力20260412.xlsx` (output), `作業ログ.csv` (run log)
- `data/` — master input Excel files; `zendam31_fixed.xlsx` is the primary input for analysis
- `index.html` / `dam_search.html` — identical files; the browser search UI

### Excel Data Structure

**Main sheet: `全国ダム地質DB`**
- Row 3+: dam records
- Col 3: dam name; Col 10: height; Col 14: manager code; Col 17: location/prefecture; Col 20–21 (T–U): latitude/longitude used for API calls
- Cols 23, 35, 47, 59, 71 (W, AI, AU, BG, BS): legend IDs for 5 geological layers
- Each legend ID column is followed by: symbol, geo_surface, geo_era, geo_rock, formation age, group, lithology, risk

**Glossary sheet:** master reference for all geological units (row 4+). Column order: `id`, `symbol`, `geo_surface`, `geo_era`, `geo_rock`, `formationAge_ja`, `group_ja`, `lithology_ja`, `geo_rock_label`, `bearing_cap`, `permeability`, `main_risk`.

### API: GeoNAVI (AIST)
- Endpoint: `https://gbank.gsj.jp/seamless/v2/api/1.2/legend.json`
- Public, no auth required
- Rate limit: 0.5s between calls; timeout 15s; up to 3 retries with backoff
- SSL verification is disabled via a custom context (`_SSL_CTX`) to work around macOS certificate handling

### Geological Layer Assignment Logic (`firstset/作業.py`)
1. Call API at exact dam coordinates → assign to geological era layer
2. If response is null (no geology): search 8 cardinal directions at 500m → 8000m expanding radius
3. If only "Late Quaternary (Q-H)" found: search surroundings for pre-Q-H geology

**Era → layer mapping:**
- Layer 1 (col W,  col 23): Pre-N  — Pre-Tertiary
- Layer 2 (col AI, col 35): N      — Tertiary (older)
- Layer 3 (col AU, col 47): N      — Tertiary (newer; same era as layer 2, second slot)
- Layer 4 (col BG, col 59): Q-old  — mid-Pleistocene
- Layer 5 (col BS, col 71): Q-H    — Late Pleistocene–Holocene

### Symbol Format
Symbols follow the pattern `{era}_{rocktype}_{modifier}` (underscore-separated). `分析2.py` decomposes them into:
- 1-item: `parts[0]` (geological era)
- 2-item: `parts[0]_parts[1]` (era + rock type)
- 3-item: full symbol

### Analysis Sheets (`分析2.py`)
Generates 7 sheets: `S1_Symbol階層分析`, `S2_強度透水性マトリクス`, `S3_Symbol類似グループ`, `S4_2項目組合せ`, `S5_北海道60ダム`, `S6_全国100ダム選定`, `S7_カバレッジ比較`.

- Hokkaido 60 dams = Hokkaido × Ministry of Land, Infrastructure, Transport and Tourism managed (top 60 with data out of ~74)
- National 100 dams = selected from non-Hokkaido dams for geological diversity

**Key metrics:**
- Bearing capacity score (`bearing_cap`): 1–5 (low → high)
- Permeability score (`permeability`): 1–4 (low → high)
- Risk = (5 − bearing_capacity) + permeability  (max 8, min 1)

### Logging
All ETL runs emit a CSV log tracking per-row status (success/error/skipped), API search radius used, and error messages. Use `--retry-log` to re-process only failed rows from a prior log.
