# Case Study 1: Vendor AI Usage Analysis - Comprehensive Guide

## Executive Summary
This project delivers a reproducible, fully automated, evidence-driven tool to determine whether third-party vendors are utilizing Artificial Intelligence within the services they provide to **Meridian Financial**. 

To ensure compliance and respect vendor boundaries, the tool relies **100% on Open Source Intelligence (OSINT)**. We do not contact the vendors directly, bypass paywalls, or scrape authenticated platforms. All evidence traces back to a cryptographically hashed, stored snapshot (URL, retrieval timestamp, content hash). Unreachable sources are recorded explicitly as "Coverage Gaps"—never reported as a false "No" finding.

---

## 🛠️ Detailed System Architecture & Module Guide

The system is designed as a pipeline containing five modular stages (M1 to M5). 

### Module 1 (M1) — Criticality Engine (`src/criticality.py`)
**Role:** Determine how critical each vendor is based on strict, repeatable criteria so that another analyst would arrive at the exact same rating.

- **How it works:** Reads the vendor profile (Columns B-K) from the input Excel file. It uses a deterministic scoring rubric (`config/rubric.yaml`) on a 0-3 scale across five metrics:
  - **Data Sensitivity** (e.g., full customer master data = 3)
  - **Payment-Flow Involvement**
  - **Regulatory Exposure**
  - **Operational Dependency**
  - **Data Volume**
- **Floor Rules & Modifiers:** If a vendor holds production credentials but no customer data, a "privileged-access modifier" applies (+1). If a vendor holds "full customer master data", they are guaranteed at least a *High* rating.
- **Output:** Assigns a final Tier (`Critical`, `High`, `Medium`, or `Low`), generates a human-readable rationale, and determines the **Assessment Depth**.

### Module 2 (M2) — OSINT Collectors (`src/collectors.py`)
**Role:** Execute the web crawl to gather evidence based on the allowed depth for that vendor's tier.

- **v2 Crawling Engine:** Dynamically discovers links via `sitemap.xml` and homepage crawling. Enforces a strict rate limit and respects `robots.txt` per RFC 9309 (treating 4xx as unrestricted and 5xx as blocked) without silently dropping entire domains.
- **Offline Caching:** Saves raw HTML, cleaned text (stripping JS blobs via Trafilatura/BeautifulSoup), URL hashes, and text lengths to `cache/`. This makes the tool fully offline-capable once downloaded and provides an immutable audit trail.
- **Deduplication:** Soft-404s and identical duplicate pages are detected via content hashing and tracked as "Coverage Gaps" rather than artificially inflating the checked page count.

### Module 3 (M3) — AI Signal Classifier (`src/classify.py`)
**Role:** Analyze the scraped data to find definitive, verifiable proof of AI usage.

- **Pre-Filtering:** Runs a regex filter with 25+ AI keywords to drop irrelevant pages quickly, saving LLM tokens.
- **Classification Schema:** Uses an LLM (OpenAI or a deterministic Mock fallback) constrained to output a strict JSON schema identifying: `signal_type`, `verbatim_quote`, `ai_mechanism`, `data_touched`, etc.
- **The Quote Guard (Crucial):** To prevent LLM hallucinations, the pipeline mandates an **exact substring check**. The quote returned by the LLM *must* exactly match the **extracted text** (parsed by Trafilatura), ensuring no whitespace or HTML markup differences cause false rejections. If it doesn't match byte-for-byte against the cleaned text, the evidence is discarded.

### Module 4 (M4) — Risk Engine & Aggregation (`src/risk_export.py`)
**Role:** Surface the risk and classify the vendor's AI usage.

- **Source Strategy & Weighting:** Not all evidence is equal. Sources are tiered in `config/sources.yaml`:
  - **Tier A (Direct):** Product docs, sub-processor lists (Highest Weight)
  - **Tier B (Corroborating):** Job postings, GitHub
  - **Tier C (Contextual):** Press releases, Executive posts
  - **Tier D (Noise):** Generic marketing (Ignored)
- **Aggregation Formula:** Final Score = Base Score × Source Reliability × Service Linkage × Recency. (Signals >180 days old suffer a 0.7x discount).
- **Coverage-Aware Verdicts:** A vendor is only marked as "No" (AI usage) if the collector successfully read at least 3 distinct pages (including high-value Tier A/B pages) and found zero valid signals. Otherwise, weak signals or heavy bot-blocking (coverage gaps) result in an honest "Inconclusive".
- **Final Classification:** Maps the verdict and criticality to the final AI Security Risk (Critical/High/Medium/None).

### Module 5 (M5) — Excel Export & UI (`src/risk_export.py` & `app/main.py`)
**Role:** Output the results seamlessly.

- **Excel Export:** Creates a byte-for-byte copy of the input workbook (preserving formatting) and only injects data into columns L-V.
- **Evidence Log:** Automatically generates a brand new sheet called `Evidence Log` detailing the source URL, vendor tier, verbatim quote, tag, and snapshot hash for complete auditability.
- **Streamlit Demo:** A web UI that orchestrates M1-M5 visually, showing "evidence cards" to the user as the pipeline runs.

---

## 🚀 Setup & Execution Instructions

### Setup
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Running the Pipeline

| Step | Command |
|------|---------|
| M1 — Criticality Engine | `python src/criticality.py` |
| M2 — OSINT Collectors | `python src/collectors.py` |
| M3 — AI Signal Classifier | `python src/classify.py` |
| M4 & M5 — Risk Engine & Excel Export | `python src/risk_export.py` |
| Full pipeline (via notebook/script) | `python notebooks/pipeline_demo.py` |

### Interactive Demo UI
```bash
streamlit run app/main.py
```
Upload `Meridian_Vendor_Input.xlsx`, choose LLM provider (mock or openai), and click **Run Pipeline**. Per-vendor evidence cards show tag, tier, verbatim quote, source URL, and retrieval date.

### LLM Configuration
Set via environment variables — **no keys in code**:
```bash
# Option 1: OpenAI (Paid)
export LLM_PROVIDER=openai       
export OPENAI_API_KEY=sk-...     
export OPENAI_MODEL=gpt-4o-mini  

# Option 2: Ollama (Free & Local)
export LLM_PROVIDER=ollama
# Requires Ollama running locally: `ollama run llama3.1`

# Option 3: Mock (Default)
export LLM_PROVIDER=mock
```

### Column V — Assessed By
The assessed-by field defaults to `"Optiv Assessment Team"`. Override via:
```bash
export ASSESSED_BY="Your Team Name"
```

### Tests
```bash
python -m pytest tests/ -v
```
All 18 tests cover: criticality determinism, substring-quote guard, weighting logic, B-K unchanged round-trip, and coverage-gap handling.

### Project Structure
```
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
├── docs/ARCHITECTURE.md
├── generate_ppt.py        # Script to generate the presentation
└── Optiv_Standalone_Presentation.pptx # The final presentation deck
```
