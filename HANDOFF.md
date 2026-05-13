# MAX Comparable Valuation App — Developer Handoff

## What This Is

A Streamlit web app for MAX Solutions (Thai M&A advisory firm) that automates Comparable Company Valuation. It takes uploaded company documents, extracts financial data via Gemini AI, screens comparable peers, ingests LSEG trading data, selects precedent transactions, and produces a completed Excel valuation file.

**Live deployment:** Streamlit Cloud from `https://github.com/chawinw-max/max-valuation-engine.git` (branch: `main`)
**Developer:** Chawin — Thai M&A advisory analyst at MAX Solutions

---

## Project Location & Structure

```
MAX Comparable Valuation App Project/
├── app.py                          # Streamlit entry point, phase routing, nav bar
├── config.py                       # Secret loader (st.secrets → config_local.py fallback)
├── config_local.py                 # LOCAL SECRETS (git-ignored) — must recreate on new machine
├── requirements.txt                # Python dependencies
├── .gitignore                      # Excludes secrets, client data, Claude config
├── .streamlit/
│   └── secrets.toml.example        # Template for Streamlit Cloud secrets
├── core/
│   ├── ai_engine.py      (1410 lines) # Gemini AI calls: extraction, peer gen, deep dive, transactions, reports
│   ├── document_parser.py (200 lines) # PDF/DOCX/XLSX/CSV parsing + LSEG peer & transaction parsers
│   ├── excel_bridge.py    (424 lines) # Writes extracted data into the Excel template (openpyxl)
│   ├── flag_engine.py     (436 lines) # Validation flags across all phases (warnings, errors, info)
│   ├── drive_api.py        (64 lines) # Google Drive upload via service account
│   └── debug_report.py    (478 lines) # Debug/diagnostic reporting
├── ui/
│   ├── phase1_upload.py   (295 lines) # File upload, year selection, business type, extraction, breakdown report
│   ├── phase2_screening.py(337 lines) # Peer generation, verification, selection funnel, manual entry
│   ├── phase3_deepdive.py (167 lines) # LSEG data parse, AI deep dive, precedent transactions
│   └── phase4_export.py   (201 lines) # Sanity check, Excel generation, Drive upload, download
├── MEGAPROMPT_v5_Comparable_Valuation_App.md  # Full app specification (cell maps, schemas, prompts)
├── ANTIGRAVITY_INSTRUCTIONS.md                # Build guide for Antigravity (Claude Code plugin)
├── DF117 - FreshSupply (EXAMPLE CASE)/        # Real test data (git-ignored)
└── pdf_extracts/                              # Cached PDF extracts (git-ignored)
```

---

## Secrets & Config (Must Recreate on New Machine)

### Option A: `config_local.py` (for local dev)

Create `config_local.py` in project root (git-ignored). Copy it directly from the old laptop at:
```
MAX Comparable Valuation App Project/config_local.py
```

It contains: `GEMINI_API_KEY`, `DRIVE_TEMPLATE_FILE_ID`, `DRIVE_OUTPUT_FOLDER_ID`, and the full `SERVICE_ACCOUNT_INFO` dict.

**DO NOT put secrets in this handoff file** — GitHub Push Protection will block it.

### Option B: Streamlit Cloud

Secrets are configured in the Streamlit Cloud dashboard under the app's Settings > Secrets. Same keys in TOML format (see `.streamlit/secrets.toml.example`).

### How `config.py` Works

`config.py` tries `st.secrets` first (Streamlit Cloud), falls back to `config_local.py` (local dev). The Gemini client is lazy-initialized via `_get_client()` in `ai_engine.py` because `st.secrets` isn't available at import time on Streamlit Cloud.

---

## Dependencies & Setup

```bash
# Clone the repo
git clone https://github.com/chawinw-max/max-valuation-engine.git
cd max-valuation-engine

# Create venv
python3 -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
# requirements.txt contains:
#   streamlit, google-genai, google-api-python-client, google-auth-httplib2,
#   google-auth-oauthlib, openpyxl, PyMuPDF, python-docx, pandas, yfinance

# Create config_local.py with the secrets above

# Run the app
streamlit run app.py
```

**Gemini model constraint:** The API key ONLY works with `gemini-2.5-flash`. Do not use `gemini-pro` or other models.

---

## Antigravity Plugin

Antigravity (Claude Code financial analysis plugin) is installed on the OLD laptop. **Install it on the new laptop too.** It provides:
- 36 slash commands and 41 skills for financial analysis
- 11 MCP connectors (LSEG, FactSet, S&P Global, Moody's, PitchBook, etc.)
- Custom agent types (earnings-reviewer, model-builder, pitch-agent, etc.)

The `.mcp.json` in the project root has the MCP server URLs — this file IS committed to git, so the connectors will be available after clone. But the Antigravity plugin itself (the Claude Code skills and commands in `.claude/commands/` and `.claude/skills/`) needs to be reinstalled on the new machine.

---

## App Workflow (4 Phases)

### Phase 1: Upload & Extract
- Upload company info files (DOCX, PDF) + financial statements (PDF, XLSX)
- Select available years (2022/2023/2024)
- Select Business Type: **Service** / **Manufacturing** / **Trading / Distribution** (drives COGS classification rules)
- Gemini extracts a Preliminary P&L (9 line items per year) + Business Model summary
- Editable P&L table for analyst corrections
- "Generate Extraction Breakdown Report" button produces a Markdown audit report

### Phase 2: Peer Screening
- Gemini generates ~30 comparable companies using a 3-ring geography model
- yfinance verifies tickers (unreliable on Streamlit Cloud — unverified peers are kept with `verified=False`)
- 30 → 12 → 5-7 selection funnel with checkboxes
- Manual peer entry tab for delisted/uncovered companies (ticker lookup or manual form)
- Rejection rationales auto-generated for not-selected peers

### Phase 3: Deep Dive & Precedent Transactions
- Upload individual LSEG `.xlsx` files for selected peers (matches by ticker AND filename)
- AI generates qualitative + quantitative comparison tables
- Upload LSEG Precedent M&A file (`.xlsx` or `.pdf`) → AI selects top 10 relevant deals

### Phase 4: Export
- AI sanity check of all data before export
- Generates completed Excel file from Google Sheets template via openpyxl
- Formula guard: never overwrites cells starting with `=`
- Google Drive upload + local download button

---

## Known Bugs & In-Progress Work

### ACTIVE BUG: D&A Double-Counting When COGS Is Derived (Partially Fixed)

**The problem:** When the source financial statement shows Gross Profit but not COGS, the app derives COGS as `Revenue - Gross Profit`. This derived COGS includes embedded D&A (e.g., equipment depreciation). The app then ALSO reports that D&A as a separate line item below EBITDA, causing double-counting.

**Example (INNO-DENT / DF103, 2024):**
- Derived COGS = 8,137,871.71 (includes ~263K of dental equipment D&A from Note 10)
- Admin = 1,615,428.37 (source was 2,094,194.65, minus 478,766.28 of admin D&A from Note 11)
- D&A reported = 742,389.75 (478K from admin + 264K from COGS = double-counted!)

**Fix applied (in `ai_engine.py` Rule #3):** Added a "CRITICAL D&A + DERIVED COGS INTERACTION" sub-rule telling Gemini:
- When COGS is derived, only separate D&A from admin/opex
- Do NOT add COGS-embedded D&A to the depreciation_amortization line
- Set D&A = null if no admin D&A can be separated

**Status:** Rule is updated in the prompt but **not yet committed or pushed**. The fix needs testing with a real case (DF103 INNO-DENT dental clinic) to verify Gemini follows the new rule correctly.

### DEFERRED: LDC.BK Peer Generation

The app should generate at least LDC.BK (Lanna Dental Clinic) and D.BK (Dentists) as peers for dental clinic businesses. Currently it sometimes misses these. The user asked to "hold off" on this improvement — revisit later.

### KNOWN ISSUE: yfinance Rate Limiting on Streamlit Cloud

yfinance verification fails frequently on Streamlit Cloud (IP-based rate limiting). The app keeps unverified peers instead of removing them, but verification status is unreliable. ThreadPoolExecutor workers reduced from 8 to 4 to mitigate.

### KNOWN ISSUE: LSEG Ticker vs Yahoo Ticker Mismatch

LSEG files use identifiers like `DNTL.TO^A26` while Yahoo Finance uses `D.BK`. The app matches by both the LSEG internal ticker AND the filename (e.g., `D.BK.xlsx` → ticker `D.BK`). This fallback was added in commit `e255e9b`.

---

## Key Architecture Decisions

1. **Lazy Gemini client init** — `_get_client()` pattern in `ai_engine.py` because `st.secrets` isn't ready at import time on Streamlit Cloud.

2. **Thinking budget vs output tokens** — Gemini 2.5 Flash's "thinking" consumes output tokens. Peer generation uses `max_output_tokens=65536` with `thinking_budget=8192` to prevent thinking from consuming the entire budget.

3. **Business type → COGS classification** — Rule #11 in the extraction prompt uses the analyst-selected business type (Service/Manufacturing/Trading) to determine which line items go into COGS vs admin.

4. **Derived COGS flag** — When COGS is back-calculated from Gross Profit (Rule #2), the `_verification` field should note this. The D&A separation rule (Rule #3) behaves differently for derived vs directly-extracted COGS.

5. **Filename-based LSEG matching** — `document_parser.py` extracts `filename_ticker` from the uploaded filename and stores it alongside the LSEG internal identifier. `flag_engine.py` and `excel_bridge.py` use both for matching.

6. **All peers kept** — Phase 2 keeps unverified peers (marked `verified=False`) instead of removing them, because yfinance verification is unreliable on cloud servers.

---

## Git State

**Repo:** `https://github.com/chawinw-max/max-valuation-engine.git`
**Branch:** `main`
**Latest commit:** `e255e9b Fix: match LSEG files by filename when LSEG ticker differs`

**Uncommitted changes:**
- `core/ai_engine.py` — D&A double-counting fix (Rule #3 update + extraction report prompt update). 18 lines added. This should be committed and pushed after testing.

**Full commit history:**
```
e255e9b Fix: match LSEG files by filename when LSEG ticker differs
e1bdbc2 Fix: keep unverified peers instead of removing them
61352bf Fix: false truncation error and handle empty peer list gracefully
bc31ca4 Add manual peer entry tab for delisted/uncovered companies
5e7d885 Fix: handle empty not-selected peers in Phase 2 checkpoint
f721f95 Fix: peer generation truncation — increase output budget
b23eff1 Fix: lazy-initialize Gemini client for Streamlit Cloud
d4f4f0e Initial commit: MAX Comparable Valuation App
```

---

## Test Data

The example case folder `DF117 - FreshSupply (EXAMPLE CASE)/` is git-ignored but should be copied to the new machine for testing. It contains:
- Company info docs (English + Thai)
- 3 years of Thai financial statements (PDFs in Buddhist Era years)
- 6 LSEG peer trading data files
- A completed output Excel file for comparison

There's also a dental clinic case (DF103 / INNO-DENT) used to test the D&A double-counting fix — files are in `~/Downloads/DF103 - Dental1 /`.

---

## Reference Documents

- **`MEGAPROMPT_v5_Comparable_Valuation_App.md`** — The complete app specification. Every cell map, JSON schema, prompt template, and workflow. Read this before making any changes to the extraction or export logic.
- **`ANTIGRAVITY_INSTRUCTIONS.md`** — Build guide and verification checklist.

---

## Quick Start on New Laptop

1. Clone the repo: `git clone https://github.com/chawinw-max/max-valuation-engine.git`
2. Set up Python venv and install requirements
3. Create `config_local.py` with the secrets (copy from above)
4. Install Antigravity on the new Claude Code instance
5. Copy test data folders (`DF117`, `DF103`) to the new machine
6. Run `streamlit run app.py` and test with the example case
7. Review the uncommitted D&A fix in `ai_engine.py` — test with DF103, then commit and push
