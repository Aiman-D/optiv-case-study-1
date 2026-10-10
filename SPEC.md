ROLE
You are a senior engineer and third-party-risk analyst building a repeatable, evidence-driven tool that determines whether vendors genuinely use AI in the services they provide to a client ("Meridian Financial"), using only each vendor's public footprint (OSINT). First produce a short plan, then STOP and wait for my approval. After approval, work one milestone at a time and verify by running.

HARD CONSTRAINTS
- Public sources only. No contact with any vendor. Respect robots.txt, terms of use, and rate limits. Do not log in to any site, bypass paywalls, or defeat anti-bot protection. Use vendors' own careers pages and permitted public sources for job postings, not scraped logged-in platforms.
- A source that cannot be reached is a COVERAGE GAP recorded in the output, never a "No" finding.
- Real organisations are involved. Never fabricate evidence. Every claim must trace to a stored snapshot (URL, retrieval timestamp, content hash).
- Evidence quotes must be VERBATIM. After the LLM returns a quote, code must verify it is an exact substring of the stored snapshot; otherwise discard it.
- "No" requires an explicit coverage statement listing the sources checked.
- Cache all fetched pages so the demo runs offline; provide a "refresh" option. Private local repo only; nothing published.

SETUP
Check for and install what is missing and tell me what you installed. LLM provider is configured through environment variables only; no keys in code.

INPUT (in /data/cs1/ - mapped to Optiv_proj/"Case Study 1 Supporting Material" locally)
Meridian_Vendor_Input.xlsx. Sheets: "Vendor Inventory", "Field Guide", "Points to consider". On Vendor Inventory the header is row 4; row 5 is V-000, a fictional worked example (use it for output style and as a format test only, never as a real vendor); vendors V-001 to V-006 are rows 6-11: AutomWorx, Fiserv, FSSI, Terrapin Technologies, BNY, The Clearing House. Columns B-K (blue headers) are provided: NEVER edit them, and do not add or remove vendors. Fill columns L-V (orange headers, cream cells): Criticality Tier, Criticality Rationale, Assessment Depth Applied, AI Usage Detected, Evidence Description along with OSINT Source(s)/Evidence Link, How Vendor May Be Using AI, Named AI Sub-processors/Model Providers, AI Security Risk Classification, Risk Rationale, Recommended Action, Assessed By/Date.
Quirks to handle:
- AI Usage Detected is headed "(Y/N)" but the Field Guide says Yes / No / Inconclusive. Use those three words.
- Evidence cell: the Field Guide wants only the exact supporting excerpt, no interpretation. Fill it with the verbatim excerpt in quotation marks, then source URL and retrieval date. No paraphrase, no summary. Put interpretation in "How Vendor May Be Using AI" and the rationale columns. If several excerpts exist, separate them clearly.
- The example refers to an "Evidence Log" sheet that does not exist. Add one new sheet "Evidence Log" (source, tier, URL, retrieval date, verbatim excerpt, tag, snapshot hash), controlled by a config flag ADD_EVIDENCE_LOG_SHEET (default true) and noted in the README as needing Optiv's confirmation. Do not alter existing sheets.
- Operational Dependency contains "Critical" for Fiserv and The Clearing House, which is outside the Field Guide scale (Total/High/Moderate/Low). Treat it as at least equal to Total and record the assumption.
- The sheet has no protection or validation, so add a test proving columns B-K are byte-for-byte unchanged after export.

BUILD
1. Criticality engine (deterministic, explainable, reproducible by another analyst): anchored 0-3 scales for data sensitivity, payment-flow involvement, regulatory exposure, operational dependency, and data volume (parse the volume text into bands), plus a privileged-access modifier (a vendor can hold no customer data yet have production credentials). Add floor rules (e.g. full-population customer data means at least High). Output Critical/High/Medium/Low, a rationale that quotes the input values and scores, and the warranted depth. Weights and anchors live in /config/rubric.yaml.
2. Depth policy (written into "Assessment Depth Applied" from the TIER, never from how much material was found): Critical = full sweep (product docs, trust centre/sub-processor list, AI policy, security/DPA pages, technical job postings, engineering blog, GitHub/SDK/API docs, regulatory filings and press/partnership releases, executive channels); High = product docs, trust centre, jobs, filings; Medium = website, docs, jobs; Low = trust centre/AI-policy check only. Stop when a confirmed signal exists and model providers are named or confirmed unfindable. Regulatory-filing search (e.g. SEC EDGAR) applies only where the vendor is a registrant.
3. Source strategy and weights (/config/sources.yaml): Tier A direct (product docs describing AI in the service, sub-processor lists, AI policies, filings); Tier B corroborating (technical job postings, engineering blogs, SDKs/repos, case studies); Tier C contextual (executive posts, press, interviews); Tier D noise (generic buzzwords). Signal weight = reliability x service-linkage (specific to the service Meridian buys vs company-wide) x recency x independent corroboration. Tags: CONFIRMED, CORROBORATED, INDICATIVE, MARKETING-ONLY, NEGATIVE-FINDING, COVERAGE-GAP. Large vendors almost certainly use AI somewhere in the company, so the verdict must hinge on service linkage.
4. Collectors: sitemap/site crawl of the public domain from the workbook, search-based discovery, vendor careers pages, GitHub, filings, trust-centre paths. Clean text with trafilatura or BeautifulSoup; dedupe; store snapshots with metadata.
5. AI signal classifier: keyword/regex pre-filter, then an LLM with a strict JSON schema {signal_type, verbatim_quote, ai_mechanism, ties_to_service, data_touched, provider_named, human_oversight_mentioned, confidence}. Enforce the substring check.
6. Aggregation to Yes/No/Inconclusive per vendor using the weights; list how the vendor may be using AI and any named providers/sub-processors.
7. Risk engine, classification Critical/High/Medium/Low/None identified. Yes: full rating from criticality, data class, evidence strength, provider transparency, and human oversight. Inconclusive: provisional rating from criticality and data class with an evidence-strength note, plus questionnaire actions. No: "None identified" with the coverage caveat. Map to Meridian risks: sensitive data exposed to AI models, undisclosed AI sub-processors, reduced transparency over decisions affecting Meridian, missing AI contract clauses. Write Recommended Action in the style of the V-000 example (questionnaire items, sub-processor register flag, contract clause asks, escalation).
8. Excel writer: fill L-V in a COPY of the workbook preserving formatting; add the Evidence Log sheet; column V = team name placeholder plus date.

DEMO
Streamlit or Gradio app plus a Colab/Jupyter notebook: upload the vendor workbook; show criticality and depth per vendor; run or load cached analysis with progress; show per-vendor evidence cards (tag, tier, verbatim quote, link, retrieval date), coverage gaps, and the risk classification; download the completed workbook.

STRUCTURE
/src (criticality, collectors, extract, classify, aggregate, risk, export), /config, /cache, /app, /notebooks, /tests (criticality determinism, substring-quote guard, weighting logic, B-K unchanged round trip, coverage-gap handling), README, docs/ARCHITECTURE.md (functional flow from OSINT collection to risk output, and the technical components inside each stage, written so I can turn it into diagrams).

MILESTONES
M1 workbook parse plus criticality for V-001..V-006 with rationales; M2 collectors plus snapshot store; M3 classifier with quote verification; M4 aggregation plus risk engine; M5 Excel export plus UI; M6 docs. After each: run tests, show results, flag every finding that rests on weak or single-source evidence, and never present an inference as confirmed.
