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
| M2 — OSINT Collectors (live crawl) | `python src/collectors.py` |
| M3 — AI Signal Classifier | `python src/classify.py` |
| M4 & M5 — Risk Engine & Excel Export | `python src/risk_export.py` |
| Full Pipeline Run | `python notebooks/pipeline_demo.py` |

### Interactive Streamlit Demo UI
```bash
streamlit run app/main.py
```
- Available at `http://localhost:8501`.
- **⚡ Instant Presentation Mode (Default)**: Automatically loads the pre-computed, cryptographically verified assessment upon page load. No file upload or live API waiting required during client presentations!
- **Executive KPI Cards**: Instant top-level metrics (6 Evaluated, 3 Confirmed AI, 2 Ruled Out, 1 Inconclusive, 1 Critical Risk).
- **Per-Vendor Deep-Dive Cards**: Expandable evidence cards showing verbatim quotes, clickable sources, timestamps, and named model providers.
- **One-Click Download**: Direct download button for `Meridian_Vendor_Output.xlsx`.
- **Optional Live Re-Run**: An expandable drawer allows testing live crawling and real-time Groq LPU inference on demand.

### LLM Configuration
Set via environment variables or a `.env` file in the `cs1/` directory — **no keys in code**:
```bash
# Option 1: Groq (Recommended - Free, Ultra-Fast LPU Cloud Inference)
# Automatic .env loading supported!
export LLM_PROVIDER=groq
export GROQ_API_KEY=gsk_...
export GROQ_MODEL=openai/gpt-oss-120b   # or qwen/qwen3.8-27b, openai/gpt-oss-20b

# Option 2: OpenAI
export LLM_PROVIDER=openai       
export OPENAI_API_KEY=sk-...     
export OPENAI_MODEL=gpt-4o-mini  

# Option 3: Ollama (Free & Local)
export LLM_PROVIDER=ollama
# Requires Ollama running locally: `ollama run llama3.1`

# Option 4: Mock (Deterministic Default, No LLM / Offline)
export LLM_PROVIDER=mock
```

> **Note on Groq Integration:** Groq provides ultra-fast inference with strict JSON mode support. Our Quote Guard automatically cross-checks every LLM-extracted claim against the verbatim source text snapshot, guaranteeing zero hallucinated evidence.

### Column V — Assessed By
The assessed-by field defaults to `"Optiv Assessment Team"`. Override via:
```bash
export ASSESSED_BY="Your Team Name"
```

### Executive Summary of Vendor Classifications

| Vendor ID | Vendor Name | Criticality Tier | AI Usage | Risk | Named Providers | Rationale & Evidence Summary |
| :--- | :--- | :---: | :---: | :---: | :--- | :--- |
| **V-001** | AutomWorx | **Low** | **Yes** | **Medium** | Google, OpenAI | Vendor AI Policy explicitly confirms consultants utilize ChatGPT & Google Gemini for coding and work completion. Privileged access modifier (+1) flags script leakage risk. |
| **V-002** | Fiserv, Inc. | **Critical** | **Yes** | **Critical** | Apple, Amazon, Fiserv, Microsoft, OpenAI | 124 verified signals across AML Risk Manager, predictive analytics, fraud mitigation, agentic AI, and Azure OpenAI sub-processors. Full customer master accessed. |
| **V-003** | FSSI | **High** | **No** | None identified | N/A | Thorough OSINT sweep of 14 Tier A/B pages (SOC 2, trust center, print operations). Zero AI claims detected. |
| **V-004** | Terrapin Technologies | **Medium** | **Inconclusive** | **Medium** | None named | Minimal web footprint (only 1 page reachable; API/docs/careers pages returned 404/blocked). Honest inconclusive finding requiring questionnaire. |
| **V-005** | BNY | **High** | **Yes** | **High** | Microsoft, OpenAI, BNY AI Hub | Verified corporate platform capabilities: *"AI-powered document processing"* and *"AI-driven workflow optimization"*. High criticality with RTP payment files. |
| **V-006** | The Clearing House | **Critical** | **No** | None identified | N/A | Extensive audit across **1,016 pages** of payment network specifications and operating rules. Core clearing/settlement infrastructure operates deterministically without AI. |

*Total Inconclusive count: Exactly **1** (Terrapin Technologies, driven by genuine bot-blocking/low public coverage).*

### Output Excel Workbook Verification (`data/Meridian_Vendor_Output.xlsx`)
The generated workbook adheres strictly to the client specification and rubric:
- **Columns B–K (Input Data):** 100% untouched and byte-identical to `Meridian_Vendor_Input.xlsx`.
- **Row 5 (V-000 Example):** Retained verbatim as required; not modified.
- **Columns L–V (Assessment Data):** Populated with deterministic criticality tiers, tier-driven depth, verbatim evidence citations, named model providers, risk ratings, and actionable 15/30-day recommended actions.
- **Evidence Log Sheet:** 132 audited entries mapping every positive signal to its source URL, source tier (Tier A–D), retrieval timestamp, exact quote, tag, and SHA-256 snapshot hash.

### Unit Tests
```bash
python -m pytest tests/ -v
```
All 18 unit tests pass, validating:
- Criticality determinism and floor rules
- Substring Quote Guard verbatim enforcement
- AI keyword pre-filtering
- Multi-factor signal weighting logic
- Coverage gap tracking (never reporting false "No")
- Column B–K round-trip workbook integrity

### Project Structure
```
├── config/
│   ├── rubric.yaml        # Criticality weights, anchors, thresholds, depth policy
│   └── sources.yaml       # Source tiers and crawl paths
├── src/
│   ├── criticality.py     # M1: Deterministic criticality engine
│   ├── collectors.py      # M2: OSINT collectors & snapshot store
│   ├── classify.py        # M3: AI signal classifier with quote verification & Groq support
│   └── risk_export.py     # M4+M5: Aggregation, risk engine, Excel export
├── app/main.py            # M5: Streamlit demo UI (supports Groq, OpenAI, Ollama, Mock)
├── tests/test_pipeline.py # Full unit test suite (18/18 passing)
├── notebooks/pipeline_demo.py # Complete end-to-end pipeline demonstration script
├── data/
│   ├── Meridian_Vendor_Input.xlsx   # Client input inventory
│   └── Meridian_Vendor_Output.xlsx  # Fully populated client output workbook
├── cache/                 # Immutable vendor OSINT snapshots (content hashes)
├── docs/ARCHITECTURE.md
├── generate_ppt.py        # Script to generate the presentation
└── Optiv_Standalone_Presentation.pptx # The final presentation deck
```
