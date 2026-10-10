# Case Study 1: Vendor AI Usage Analysis

## Overview
A reproducible, OSINT-driven tool that evaluates third-party vendors to determine their use of Artificial Intelligence and flags associated security risks for **Meridian Financial**. All evidence traces to a stored snapshot (URL, retrieval timestamp, content hash). Coverage gaps are recorded explicitly — never reported as a "No" finding.

## Setup
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Running the Pipeline

| Step | Command |
|------|---------|
| M1 — Criticality Engine | `python src/criticality.py` |
| M2 — OSINT Collectors | `python src/collectors.py` |
| M3 — AI Signal Classifier | `python src/classify.py` |
| M4 & M5 — Risk Engine & Excel Export | `python src/risk_export.py` |
| Full pipeline (via notebook) | `python notebooks/pipeline_demo.py` |

## Demo UI
```bash
streamlit run app/main.py
```

Upload `Meridian_Vendor_Input.xlsx`, choose LLM provider (mock or openai), and click **Run Pipeline**. Per-vendor evidence cards show tag, tier, verbatim quote, source URL, and retrieval date.

## LLM Configuration
Set via environment variables — **no keys in code**:
```bash
export LLM_PROVIDER=openai       # or "mock" (default)
export OPENAI_API_KEY=sk-...     # required if provider is openai
export OPENAI_MODEL=gpt-4o-mini  # optional, default: gpt-4o-mini
```

## Column V — Assessed By
The assessed-by field defaults to `"Optiv Assessment Team"`. Override via:
```bash
export ASSESSED_BY="Your Team Name"
```

## Evidence Log Sheet
`ADD_EVIDENCE_LOG_SHEET = True` (default). The export process adds a new `Evidence Log` sheet to the output workbook containing: source, tier, URL, retrieval date, verbatim excerpt, tag, snapshot hash.

> **Confirmation needed from Optiv**: Is the Evidence Log sheet acceptable as an addition to the workbook?

## Tests
```bash
python -m pytest tests/ -v
```
All 18 tests cover: criticality determinism, substring-quote guard, weighting logic, B-K unchanged round-trip, and coverage-gap handling.

## Project Structure
```
cs1/
├── config/
│   ├── rubric.yaml        # Criticality weights, anchors, thresholds, depth policy
│   └── sources.yaml       # Source tiers and crawl paths
├── src/
│   ├── criticality.py     # M1: Deterministic criticality engine
│   ├── collectors.py      # M2: OSINT collectors & snapshot store
│   ├── classify.py        # M3: AI signal classifier with quote verification
│   └── risk_export.py     # M4+M5: Aggregation, risk engine, Excel export
├── app/main.py            # M5: Streamlit demo UI
├── tests/test_pipeline.py # Full test suite
├── notebooks/pipeline_demo.py
├── data/                  # Input/output workbooks
├── cache/                 # Vendor OSINT snapshots (git-ignored)
└── docs/ARCHITECTURE.md
```
