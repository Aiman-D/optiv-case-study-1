# CS1 Architecture: Vendor AI Usage Analysis

## Functional Flow

```
[Excel Workbook: Meridian_Vendor_Input.xlsx]
        │
        ▼
┌───────────────────┐
│  M1: Criticality  │  Reads columns B-K; deterministic scoring on 5 scales
│  Engine           │  → tier (Critical/High/Medium/Low) + rationale + depth
└────────┬──────────┘
         │  tier drives depth policy
         ▼
┌───────────────────┐
│  M2: OSINT        │  Crawls vendor public footprint up to allowed depth
│  Collectors       │  Respects robots.txt; 1.2s rate limit; cached snapshots
└────────┬──────────┘  (cache/<VID>/<url_hash>.json + crawl_summary.json)
         │  snapshots
         ▼
┌───────────────────┐
│  M3: AI Signal    │  Keyword pre-filter → LLM/mock classifier (JSON schema)
│  Classifier       │  → verbatim quote verification (exact substring check)
└────────┬──────────┘  → signals.json: CONFIRMED/CORROBORATED/INDICATIVE/
         │               MARKETING-ONLY/NEGATIVE-FINDING/COVERAGE-GAP
         ▼
┌───────────────────┐
│  M4: Aggregation  │  Weighted score = base × tier × confidence × linkage × recency
│  & Risk Engine    │  → Yes/No/Inconclusive verdict
│                   │  → Risk: Critical/High/Medium/None identified
└────────┬──────────┘  → Recommended action (questionnaire, contract clauses)
         │
         ▼
┌───────────────────┐
│  M5: Excel        │  shutil.copy of input; fills columns L-V only
│  Export           │  B-K verified byte-for-byte unchanged; Evidence Log sheet
└────────┬──────────┘
         │
         ▼
[Meridian_Vendor_Output.xlsx]  +  [Streamlit UI / Notebook]
```

## Technical Components

### M1 — Criticality Engine (`src/criticality.py`)
- **Inputs**: pandas DataFrame row from Vendor Inventory (columns B-K)
- **Config**: `config/rubric.yaml` (scales, thresholds, floor rules, depth policy)
- **Scoring**: 5 anchored 0-3 scales × privileged-access modifier (+1)
  - data_sensitivity, payment_flow, regulatory_exposure, operational_dependency, data_volume
- **Floor rules**: "full customer master" → at least High; "Critical" dependency → Critical
- **Output**: `(tier, rationale_string, depth_string)`

### M2 — OSINT Collectors (`src/collectors.py`)
- **robots.txt**: `urllib.robotparser.RobotFileParser` — blocks if disallowed
- **Rate limiting**: 1.2s sleep between requests
- **Text extraction**: `trafilatura.extract()` with BeautifulSoup fallback
- **Snapshot schema**: `{url, url_hash, timestamp (ISO-8601 UTC), status, cleaned_text, content_hash}`
- **Coverage gaps**: any non-200 or empty-content URL → COVERAGE-GAP record

### M3 — AI Signal Classifier (`src/classify.py`)
- **Pre-filter**: 25+ regex patterns (AI keywords, model provider names)
- **LLM schema**: `{signal_type, verbatim_quote, ai_mechanism, ties_to_service, data_touched, provider_named, human_oversight_mentioned, confidence}`
- **Quote guard**: `quote in snapshot['cleaned_text']` — exact substring; discarded if fails
- **Fallback**: smart mock classifier (sentence-level analysis, no LLM required)

### M4 — Aggregation & Risk (`src/risk_export.py: aggregate_vendor()`)
- **Weight formula**: `base_score × tier_weight × confidence_weight × service_linkage × recency`
  - Recency: signals > 180 days old → 0.7× discount
- **Verdict thresholds**: ≥10 → Yes; ≥3 → Inconclusive; <3 → No
- **Risk mapping**: Yes+Critical tier → Critical risk; Inconclusive+Critical/High → High risk

### M5 — Excel Export (`src/risk_export.py: process_and_export()`)
- **Copy first**: `shutil.copy(input, output)` — preserves all formatting
- **Write columns L-V** only, never B-K
- **Evidence Log sheet**: source, tier, URL, retrieval date, verbatim excerpt, tag, snapshot hash

### App (`app/main.py`)
- Streamlit: upload → M1 → M2 → M3 → M4/M5 → per-vendor evidence cards → download
- Evidence cards: tag colour, tier, verbatim quote, source URL, retrieval date
- Coverage gaps listed explicitly per vendor

### Tests (`tests/test_pipeline.py`)
| Test class | What it verifies |
|---|---|
| TestCriticalityDeterminism | Same input → same tier + rationale (18 runs) |
| TestSubstringQuoteGuard | Invalid quotes rejected; valid quotes pass |
| TestPreFilter | AI keyword detection; no false positives on non-AI text |
| TestWeightingLogic | Score formula: 5×4×1.0×1.5 = 30 for Tier_A/High/service-linked |
| TestCoverageGapHandling | Unreachable source → COVERAGE-GAP, not "No" |
| TestBKUnchanged | Columns B-K byte-for-byte identical before/after export |
