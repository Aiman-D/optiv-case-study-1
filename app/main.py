"""
CS1 Streamlit App — Vendor AI Risk Analyzer.

Wires the full pipeline:
  Upload → Criticality → (Cached) OSINT → Classify → Risk/Export → Download.
Per-vendor evidence cards with tag, tier, verbatim quote, link, retrieval date.
Coverage gaps and risk classification are surfaced explicitly.
"""

import streamlit as st
import pandas as pd
import sys
import os
import json
from pathlib import Path

# Ensure src is importable when run from the app/ directory
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

BASE_DIR = Path(__file__).parent.parent

st.set_page_config(page_title="Vendor AI Risk Analyzer — Meridian Financial", layout="wide")

st.title("🔍 Meridian Financial — Vendor AI Risk Analyzer")
st.markdown(
    "Upload the Vendor Inventory to run OSINT collection, AI signal classification, "
    "and risk assessment. Evidence is displayed per vendor; coverage gaps are flagged explicitly."
)

# ── Sidebar: settings ──────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    refresh_cache = st.checkbox("Refresh OSINT cache (live fetch)", value=False)
    add_evidence_log = st.checkbox("Add Evidence Log sheet to output", value=True)
    llm_provider = st.selectbox(
        "LLM Provider",
        options=["mock", "openai"],
        help="Use 'mock' for offline demo. Set OPENAI_API_KEY env var for 'openai'."
    )
    st.markdown("---")
    st.caption("Evidence quotes are verified as exact substrings of cached snapshots. "
               "Coverage gaps are recorded — never reported as 'No'.")

# ── Main: upload ───────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload Vendor Input Excel (.xlsx)",
    type=["xlsx"],
    help="Upload Meridian_Vendor_Input.xlsx. Vendors V-001 to V-006 will be processed."
)

col_run, col_load = st.columns([1, 3])
run_pipeline = col_run.button("▶ Run Pipeline", type="primary")
load_existing = col_load.button("📂 Load Existing Results (from cache)")

if run_pipeline and uploaded_file is None:
    st.warning("Please upload a vendor workbook first.")
    st.stop()

# ── Shared: write uploaded file ────────────────────────────────
if run_pipeline and uploaded_file is not None:
    data_dir = BASE_DIR / "data"
    data_dir.mkdir(exist_ok=True)
    input_path = data_dir / "Meridian_Vendor_Input.xlsx"
    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    os.environ["LLM_PROVIDER"] = llm_provider

    # M1: Criticality
    with st.spinner("🔹 M1 — Criticality Engine…"):
        from criticality import evaluate_vendor, load_rubric
        rubric = load_rubric(str(BASE_DIR / "config" / "rubric.yaml"))
        df = pd.read_excel(str(input_path), sheet_name="Vendor Inventory", header=3)
        df = df[df["Vendor ID"].str.startswith("V-", na=False) & (df["Vendor ID"] != "V-000")]
        criticality_results = {}
        for _, row in df.iterrows():
            tier, rationale, depth = evaluate_vendor(row, rubric)
            criticality_results[row["Vendor ID"]] = {
                "vendor": row["Vendor Name"],
                "tier": tier,
                "rationale": rationale,
                "depth": depth,
                "url": row.get("Public Website", ""),
            }
    st.success(f"✅ M1 complete — {len(criticality_results)} vendors tiered.")

    # M2: Collectors
    with st.spinner("🔹 M2 — OSINT Collectors (may take a few minutes)…"):
        from collectors import OSINTCollector
        collector = OSINTCollector(
            str(BASE_DIR / "config" / "sources.yaml"),
            cache_dir=str(BASE_DIR / "cache")
        )
        progress_bar = st.progress(0)
        total = len(criticality_results)
        for i, (vid, info) in enumerate(criticality_results.items()):
            url = str(info.get("url", "") or "")
            if url and not url.startswith("http"):
                url = "https://" + url
            if url:
                collector.crawl_vendor(vid, url, depth_tier=info["tier"], refresh=refresh_cache)
            progress_bar.progress((i + 1) / total)
    st.success("✅ M2 complete — snapshots stored.")

    # M3: Classify
    with st.spinner("🔹 M3 — AI Signal Classification…"):
        from classify import classify_vendor
        cache_path = BASE_DIR / "cache"
        for vid in criticality_results:
            classify_vendor(vid, cache_dir=str(cache_path))
    st.success("✅ M3 complete — signals classified with quote verification.")

    # M4 + M5: Risk Engine + Export
    with st.spinner("🔹 M4 & M5 — Risk Engine & Excel Export…"):
        from risk_export import process_and_export
        # Patch the ADD_EVIDENCE_LOG_SHEET flag at runtime
        import risk_export as _re
        _re.ADD_EVIDENCE_LOG_SHEET = add_evidence_log
        process_and_export()
    st.success("✅ M4/M5 complete — workbook exported.")

# ── Display results ────────────────────────────────────────────
output_path = BASE_DIR / "data" / "Meridian_Vendor_Output.xlsx"
cache_path = BASE_DIR / "cache"

if (run_pipeline and output_path.exists()) or (load_existing and output_path.exists()):
    st.markdown("---")
    st.header("📊 Results")

    # Summary table
    try:
        df_out = pd.read_excel(str(output_path), sheet_name="Vendor Inventory", header=3)
        df_out = df_out[
            df_out["Vendor ID"].str.startswith("V-", na=False) &
            (df_out["Vendor ID"] != "V-000")
        ]
        display_cols = [
            "Vendor ID", "Vendor Name", "Criticality Tier",
            "Assessment Depth Applied", "AI Usage Detected (Y/N)",
            "AI Security Risk Classification", "Recommended Action"
        ]
        present = [c for c in display_cols if c in df_out.columns]
        if present:
            st.subheader("Vendor Summary")
            st.dataframe(df_out[present], use_container_width=True)
    except Exception as e:
        st.warning(f"Could not load summary table: {e}")

    # Per-vendor evidence cards
    st.subheader("Per-Vendor Evidence Cards")
    if cache_path.exists():
        for vendor_dir in sorted(cache_path.iterdir()):
            if not vendor_dir.is_dir() or not vendor_dir.name.startswith("V-"):
                continue
            signals_file = vendor_dir / "signals.json"
            if not signals_file.exists():
                continue
            with open(signals_file) as f:
                signals_data = json.load(f)

            vid = vendor_dir.name
            signals = signals_data.get("signals", [])
            sources = signals_data.get("sources_checked", [])

            valid = [s for s in signals
                     if s.get("valid") and s["signal_type"] not in
                     ("COVERAGE-GAP", "NEGATIVE-FINDING", "INVALID_QUOTE")]
            gaps = [s for s in signals if s["signal_type"] == "COVERAGE-GAP"]

            tag_colors = {
                "CONFIRMED": "🟢", "CORROBORATED": "🔵",
                "INDICATIVE": "🟡", "MARKETING-ONLY": "🟠",
                "NEGATIVE-FINDING": "⚪", "COVERAGE-GAP": "🔴"
            }

            with st.expander(f"{vid} — {len(valid)} signals | {len(gaps)} coverage gap(s)"):
                if valid:
                    st.markdown("**Evidence signals:**")
                    for s in valid:
                        tag = s.get("signal_type", "")
                        tier = s.get("source_tier", "?")
                        quote = s.get("verbatim_quote", "")
                        url = s.get("url", "")
                        ts = s.get("timestamp", "")[:10] if s.get("timestamp") else "N/A"
                        conf = s.get("confidence", "?")
                        provider = s.get("provider_named", "None")
                        icon = tag_colors.get(tag, "❓")
                        st.markdown(
                            f"{icon} **{tag}** | Tier: `{tier}` | Confidence: `{conf}` | "
                            f"Provider: `{provider}`"
                        )
                        if quote:
                            st.markdown(f'> "{quote}"')
                        if url:
                            st.caption(f"Source: [{url}]({url}) — retrieved {ts}")
                        st.markdown("---")
                else:
                    st.info("No positive AI signals found.")

                if gaps:
                    st.markdown(f"**Coverage gaps ({len(gaps)}):**")
                    for g in gaps:
                        st.caption(f"🔴 {g.get('url', 'unknown')} — {g.get('reason', 'unreachable')}")

    # Download
    st.markdown("---")
    with open(str(output_path), "rb") as f:
        st.download_button(
            label="⬇️ Download Completed Workbook",
            data=f,
            file_name="Meridian_Vendor_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

elif load_existing and not output_path.exists():
    st.info("No existing output found. Upload a workbook and run the pipeline first.")
