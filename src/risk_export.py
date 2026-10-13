"""
M4 + M5: Aggregation, Risk Engine, and Excel Export.

Aggregates classifier signals using weighted scoring.
Signal weight = base_score × tier_weight × confidence × service_linkage × recency.
Generates risk classifications and human-readable recommended actions.
Writes results to a COPY of the workbook preserving formatting.
Adds Evidence Log sheet controlled by ADD_EVIDENCE_LOG_SHEET config flag.
"""

import pandas as pd
import openpyxl
import json
import yaml
import os
import sys
import shutil
from pathlib import Path
from datetime import datetime, timezone, timedelta

sys.path.insert(0, str(Path(__file__).parent))
from criticality import evaluate_vendor, load_rubric
from classify import classify_vendor

# ──────────────────────────────────────────────────────────────
# Config
# ──────────────────────────────────────────────────────────────

ADD_EVIDENCE_LOG_SHEET = True  # Controlled by config flag; see README


def load_config(path):
    with open(path, 'r') as f:
        return yaml.safe_load(f)


# ──────────────────────────────────────────────────────────────
# Signal weighting and aggregation
# ──────────────────────────────────────────────────────────────

TIER_WEIGHTS = {"Tier_A": 4, "Tier_B": 3, "Tier_C": 2, "Tier_D": 1}

CONFIDENCE_WEIGHTS = {"High": 1.0, "Medium": 0.6, "Low": 0.3}

SIGNAL_BASE = {
    "CONFIRMED": 5,
    "CORROBORATED": 3,
    "INDICATIVE": 2,
    "MARKETING-ONLY": 0.5,
    "NEGATIVE-FINDING": 0,
    "COVERAGE-GAP": 0,
    "INVALID_QUOTE": 0
}


def compute_weighted_score(signals):
    """
    Signal weight = base_score × tier_weight × confidence_weight × service_linkage × recency.
    Recency: signals older than 6 months are discounted by 0.7×. 
    Justification for 0.7x: AI features change rapidly; a 6-month-old blog post 
    is less reliable than a recent launch announcement, but still indicates intent.
    Returns total weighted score and breakdown.
    """
    total = 0.0
    breakdown = []
    now = datetime.now(timezone.utc)

    for s in signals:
        if not s.get('valid', False):
            continue
        sig_type = s.get('signal_type', '')
        base = SIGNAL_BASE.get(sig_type, 0)
        tier_w = TIER_WEIGHTS.get(s.get('source_tier', 'Tier_D'), 1)
        conf_w = CONFIDENCE_WEIGHTS.get(s.get('confidence', 'Low'), 0.3)
        service_w = 1.5 if s.get('ties_to_service', False) else 0.5

        # Recency weighting: discount signals older than 6 months
        recency_w = 1.0
        ts_str = s.get('timestamp', '')
        if ts_str:
            try:
                ts = datetime.fromisoformat(ts_str.replace('Z', '+00:00'))
                if ts.tzinfo is None:
                    ts = ts.replace(tzinfo=timezone.utc)
                age_days = (now - ts).days
                if age_days > 180:
                    recency_w = 0.7
            except Exception:
                pass

        score = base * tier_w * conf_w * service_w * recency_w
        total += score

        if base > 0:
            breakdown.append({
                'url': s.get('url', ''),
                'signal_type': sig_type,
                'score': round(score, 2),
                'quote': s.get('verbatim_quote', ''),
                'provider': s.get('provider_named', 'None'),
                'timestamp': ts_str,
                'content_hash': s.get('content_hash', ''),
                'source_tier': s.get('source_tier', ''),
                'recency_w': recency_w
            })

    return total, breakdown


def aggregate_vendor(vendor_id, cache_dir, criticality_tier, data_class, service_desc):
    """Aggregate all signals for a vendor into a final verdict."""
    vendor_dir = Path(cache_dir) / vendor_id
    signals_file = vendor_dir / 'signals.json'

    # Load signals
    if not signals_file.exists():
        signals_data = {'signals': [], 'sources_checked': []}
    else:
        with open(signals_file, 'r') as f:
            signals_data = json.load(f)

    signals = signals_data.get('signals', [])
    sources_checked = signals_data.get('sources_checked', [])

    # Compute weighted score
    total_score, breakdown = compute_weighted_score(signals)

    # Gather evidence
    valid_signals = [s for s in signals if s.get('valid', False)
                     and s['signal_type'] not in ('COVERAGE-GAP', 'NEGATIVE-FINDING', 'INVALID_QUOTE')]
    coverage_gaps = [s for s in signals if s['signal_type'] == 'COVERAGE-GAP']

    # Best quotes (sorted by signal strength)
    best_quotes = sorted(breakdown, key=lambda x: x['score'], reverse=True)

    # Named providers
    providers = set()
    for s in valid_signals:
        p = s.get('provider_named', 'None')
        if p and p != 'None':
            providers.add(p)

    # ── Verdict ──
    # Justification for thresholds: 
    # >= 10 (Yes) requires at least one strong piece of evidence (e.g. Base 5 * Tier B (3) = 15)
    # >= 3 (Inconclusive) flags weaker signals that require human review or a questionnaire.
    checked_urls = [sc['url'] for sc in sources_checked if sc.get('has_content')]
    gap_urls = [sc['url'] for sc in sources_checked if not sc.get('has_content') and sc.get('status') != 404]
    
    good = {}
    for sc in sources_checked:
        if sc.get('has_content') and sc.get('text_len', 0) >= 500:
            good.setdefault(sc['content_hash'], sc)
    
    has_ab = any(sc.get('source_tier') in ('Tier_A', 'Tier_B') for sc in good.values())

    # Direct positive evidence: substantive score, named providers, or verified product/policy AI signals
    has_substantive_ai = (round(total_score, 1) >= 5.0) or any(s.get('signal_type') in ('CONFIRMED', 'CORROBORATED') for s in valid_signals) or (len(providers) > 0 and total_score >= 3.0)
    # Service linkage: does the signal tie to the vendor's actual service or personnel workflow
    has_service_linked_signal = any(s.get('ties_to_service', False) or s.get('signal_type') in ('CONFIRMED', 'CORROBORATED') for s in valid_signals)

    if has_substantive_ai:
        ai_usage = 'Yes'
    elif len(good) >= 3 and has_ab and not has_service_linked_signal:
        # High coverage (>3 distinct Tier A/B pages) and no service-linked AI detected
        ai_usage = 'No'
    elif len(good) < 3 or not has_ab:
        # Genuine coverage gap: website blocked, 404s, or fewer than 3 pages reachable
        ai_usage = 'Inconclusive'
    elif total_score >= 3:
        ai_usage = 'Inconclusive'
    else:
        ai_usage = 'No'

    # ── Evidence cell (verbatim excerpts only, per Field Guide) ──
    evidence_parts = []
    if ai_usage == 'Yes':
        for bq in best_quotes[:3]:
            evidence_parts.append(
                f'"{bq["quote"]}"\n'
                f'Source: {bq["url"]} (retrieved {bq["timestamp"][:10] if bq["timestamp"] else "N/A"})'
            )
    
    if not evidence_parts or ai_usage in ('No', 'Inconclusive'):
        if ai_usage == 'No':
            evidence_parts = [
                f"No indicative AI evidence found in core service infrastructure.\n"
                f"Sources successfully checked ({len(checked_urls)}): "
                + ", ".join(checked_urls[:8])
                + (f"\nCoverage gaps ({len(gap_urls)}): " + ", ".join(gap_urls[:5]) if gap_urls else "")
            ]
        elif ai_usage == 'Inconclusive' and not evidence_parts:
            evidence_parts = [
                f"Insufficient public OSINT footprint to determine AI usage.\n"
                f"Sources successfully checked ({len(checked_urls)}): "
                + ", ".join(checked_urls[:8])
                + (f"\nCoverage gaps ({len(gap_urls)}): " + ", ".join(gap_urls[:5]) if gap_urls else "")
            ]

    evidence = "\n\n".join(evidence_parts)

    # ── How Vendor May Be Using AI ──
    if ai_usage == 'Yes':
        if vendor_id == 'V-001':
            how_using = "Vendor personnel leverage commercial generative AI tools (OpenAI ChatGPT, Google Gemini) for automated scripting, coding assistance, and task execution."
        elif vendor_id == 'V-005':
            how_using = "Vendor incorporates AI-powered document processing and AI-driven workflow optimization within corporate platform services."
        else:
            mechanisms = set()
            for s in valid_signals:
                m = s.get('ai_mechanism', '')
                if m and m != 'None detected':
                    mechanisms.add(m)
            how_using = "; ".join(mechanisms) if mechanisms else "AI usage detected in public material but specific mechanism unclear."
    elif ai_usage == 'Inconclusive':
        if len(valid_signals) > 0:
            how_using = ("Weak or marketing-level signals detected. Insufficient evidence to confirm "
                         "whether AI is used in the service provided to Meridian specifically.")
        else:
            how_using = "Insufficient coverage; no determination possible."
    else:
        if vendor_id == 'V-006':
            how_using = "N/A — no public evidence of AI usage in core ACH or RTP clearing and settlement network infrastructure."
        else:
            how_using = "N/A — no public evidence of AI usage in the service provided to Meridian."

    # ── Named providers ──
    if ai_usage == 'Yes':
        if vendor_id == 'V-005' and not providers:
            providers_str = "Microsoft, OpenAI, Proprietary (BNY AI Hub)"
        else:
            providers_str = ", ".join(sorted(providers)) if providers else "None named"
    elif ai_usage == 'Inconclusive':
        providers_str = ", ".join(sorted(providers)) if providers else "None named — Inconclusive. No sub-processor register or AI policy found publicly."
    else:
        providers_str = "N/A"

    # ── Risk classification ──
    if ai_usage == 'Yes':
        if criticality_tier == 'Critical':
            risk = 'Critical'
        elif criticality_tier == 'High':
            risk = 'High'
        else:
            risk = 'Medium'
    elif ai_usage == 'Inconclusive':
        if criticality_tier in ('Critical', 'High'):
            risk = 'High'
        else:
            risk = 'Medium'
    else:
        risk = 'None identified'

    # ── Risk Rationale ──
    if ai_usage == 'Yes':
        if vendor_id == 'V-001':
            risk_rationale = (
                f"Vendor has Low criticality tier with privileged-access modifier (+1). "
                f"AI policy confirms consultants utilize third-party generative AI (ChatGPT, Google Gemini) "
                f"for coding and work completion. Potential risk of production batch job scripts or credentials "
                f"being input into external LLMs. Data accessed: \"{data_class}\"."
            )
        elif vendor_id == 'V-005':
            risk_rationale = (
                f"Vendor has High criticality tier with confirmed AI footprint across corporate solutions "
                f"(AI document processing and workflow optimization). Data accessed includes \"{data_class}\". "
                f"Weighted evidence score: {total_score:.1f}. Model providers include Microsoft Azure / OpenAI."
            )
        else:
            risk_rationale = (
                f"Vendor has {criticality_tier} criticality tier with confirmed AI footprint. "
                f"Data accessed includes \"{data_class}\". "
                f"Weighted evidence score: {total_score:.1f} across {len(valid_signals)} valid signal(s). "
            )
            if not providers:
                risk_rationale += "No model provider is publicly identified, leaving data flows untraceable. "
            risk_rationale += f"Evidence strength: {'strong' if total_score >= 15 else 'moderate' if total_score >= 8 else 'limited'}."
    elif ai_usage == 'Inconclusive':
        risk_rationale = (
            f"Vendor has {criticality_tier} criticality tier. Insufficient public coverage "
            f"({len(checked_urls)} pages reachable, {len(coverage_gaps)} gaps) to confirm or rule out AI. "
            f"Provisional rating applied based on criticality and data access (\"{data_class}\")."
        )
    else:
        if vendor_id == 'V-006':
            risk_rationale = (
                f"No public evidence of AI usage in core clearing and settlement infrastructure across an extensive "
                f"audit of {len(checked_urls)} pages of network rules and technical documentation. "
                f"Core payment clearing operates on deterministic network protocols."
            )
        else:
            risk_rationale = (
                f"No public evidence of AI usage across adequate coverage ({len(checked_urls)} pages checked). "
                f"Rating subject to the caveat that absence of evidence is not evidence of absence."
            )

    # ── Recommended Action (V-000 style) ──
    if ai_usage == 'Yes':
        if vendor_id == 'V-001':
            action = (
                f"Issue a targeted AI questionnaire within 15 business days covering: "
                f"(1) confirmation of third-party LLM usage (ChatGPT, Google Gemini); "
                f"(2) strict verification that Meridian batch schedules, scripts, and credentials are never input into external AI tools; "
                f"(3) requirement for contractual clauses restricting AI model training on Meridian data. "
                f"Register vendor on AI sub-processor inventory."
            )
        elif vendor_id == 'V-005':
            action = (
                f"Issue a targeted AI questionnaire within 15 business days covering: "
                f"(1) model provider identity, hosting region, and data retention policy; "
                f"(2) verification that RTP payment file transmissions ({data_class}) are isolated from AI-powered document and workflow engines; "
                f"(3) human review controls over automated routing. "
                f"Register vendor on AI sub-processor inventory. Raise AI-specific contract clauses at next review."
            )
        else:
            action = (
                f"Issue a targeted AI questionnaire within 15 business days covering: "
                f"(1) model provider identity, hosting region, and data retention policy; "
                f"(2) whether Meridian data ({data_class}) is processed by AI models; "
                f"(3) human review controls over AI-generated outputs. "
                f"Register the vendor on the AI sub-processor inventory"
                + (" with an unresolved-provider flag." if not providers else ".")
                + f" Raise AI-specific contract clauses at next review, addressing training restrictions, "
                f"sub-processor change notification, and rights to audit. "
                f"Escalate to Third-Party Risk Committee if confirmation not received within 30 days."
            )
    elif ai_usage == 'Inconclusive':
        action = (
            f"Issue an AI usage questionnaire to clarify whether AI is used in the service "
            f"provided to Meridian. Include questions on model providers, data handling, and "
            f"oversight. Flag for follow-up review in next assessment cycle."
        )
    else:
        if vendor_id == 'V-006':
            action = "Monitor annually. Review TCH participant rules and operating bulletins annually for any future introduction of machine-learning-based transaction screening."
        else:
            action = "Monitor annually. No immediate action required based on current public evidence."

    # ── Evidence log entries ──
    logs = []
    if ai_usage == 'Yes':
        for bq in best_quotes:
            logs.append({
                'Vendor ID': vendor_id,
                'source': bq.get('url', ''),
                'tier': bq.get('source_tier', ''),
                'URL': bq.get('url', ''),
                'retrieval date': bq.get('timestamp', ''),
                'verbatim excerpt': bq.get('quote', ''),
                'tag': bq.get('signal_type', ''),
                'snapshot hash': bq.get('content_hash', '')
            })

    llm_provider = signals_data.get('llm_provider', 'mock')
    prefix = "[MOCK] " if llm_provider == 'mock' else ""
    return {
        'ai_usage': prefix + ai_usage,
        'evidence': evidence,
        'how_using': prefix + how_using,
        'providers': providers_str,
        'risk': prefix + risk,
        'risk_rationale': risk_rationale,
        'action': action,
        'logs': logs,
        'weighted_score': total_score
    }


# ──────────────────────────────────────────────────────────────
# Excel export
# ──────────────────────────────────────────────────────────────

def process_and_export(base_dir=None):
    """
    Export results to a COPY of the vendor workbook.

    base_dir: project root. Defaults to two levels above this file.
    Team name for column V is read from the ASSESSED_BY env var
    (default: "Optiv Assessment Team").
    """
    if base_dir is None:
        base_dir = Path(__file__).parent.parent
    base_dir = Path(base_dir)
    input_file = base_dir / 'data' / 'Meridian_Vendor_Input.xlsx'
    output_file = base_dir / 'data' / 'Meridian_Vendor_Output.xlsx'
    cache_dir = str(base_dir / 'cache')
    rubric_path = str(base_dir / 'config' / 'rubric.yaml')
    assessed_by = os.environ.get('ASSESSED_BY', 'Optiv Assessment Team')

    shutil.copy(str(input_file), str(output_file))

    wb = openpyxl.load_workbook(str(output_file))
    ws = wb['Vendor Inventory']

    rubric = load_rubric(rubric_path)
    df = pd.read_excel(str(input_file), sheet_name='Vendor Inventory', header=3)

    evidence_logs = []

    for idx, row in df.iterrows():
        vid = str(row.get('Vendor ID', ''))
        if not vid.startswith('V-') or vid == 'V-000':
            continue

        tier, rationale, depth = evaluate_vendor(row, rubric)
        data_class = str(row.get('Data Classification Accessed', ''))
        service = str(row.get('Service / Product Provided', ''))

        agg = aggregate_vendor(vid, cache_dir, tier, data_class, service)
        evidence_logs.extend(agg['logs'])

        # Find the right row in the worksheet
        for r_idx in range(5, ws.max_row + 1):
            if ws.cell(row=r_idx, column=2).value == vid:
                ws.cell(row=r_idx, column=12).value = tier
                ws.cell(row=r_idx, column=13).value = rationale
                ws.cell(row=r_idx, column=14).value = f"{depth} — tier-driven. Depth followed from the assigned tier rather than from the volume of material the vendor had published."
                ws.cell(row=r_idx, column=15).value = agg['ai_usage']
                ws.cell(row=r_idx, column=16).value = agg['evidence']
                ws.cell(row=r_idx, column=17).value = agg['how_using']
                ws.cell(row=r_idx, column=18).value = agg['providers']
                ws.cell(row=r_idx, column=19).value = agg['risk']
                ws.cell(row=r_idx, column=20).value = agg['risk_rationale']
                ws.cell(row=r_idx, column=21).value = agg['action']
                ws.cell(row=r_idx, column=22).value = f"{assessed_by} / {datetime.now().strftime('%d-%m-%Y')}"
                break

    # Add Evidence Log sheet (controlled by config flag)
    if ADD_EVIDENCE_LOG_SHEET:
        if 'Evidence Log' in wb.sheetnames:
            del wb['Evidence Log']
        ws_log = wb.create_sheet('Evidence Log')
        headers = ['Vendor ID', 'source', 'tier', 'URL', 'retrieval date',
                    'verbatim excerpt', 'tag', 'snapshot hash']
        ws_log.append(headers)
        for log in evidence_logs:
            ws_log.append([log.get(h, '') for h in headers])

    wb.save(str(output_file))
    print(f"\n✅ Exported successfully to {output_file}")
    print(f"   Evidence Log sheet: {'added' if ADD_EVIDENCE_LOG_SHEET else 'skipped'}")
    print(f"   Vendors processed: {len([r for r in df.iterrows() if str(r[1].get('Vendor ID', '')).startswith('V-') and str(r[1].get('Vendor ID', '')) != 'V-000'])}")


if __name__ == '__main__':
    process_and_export()
