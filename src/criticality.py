"""
M1: Criticality Engine — deterministic, explainable, reproducible.

Scores vendors on 5 anchored 0-3 scales plus a privileged-access modifier.
All weights and anchors live in /config/rubric.yaml.
Output: tier (Critical/High/Medium/Low), a human-readable rationale that
quotes input values and scores, and the warranted OSINT depth.
"""

import pandas as pd
import yaml
import re

def load_rubric(path):
    with open(path, 'r') as f:
        return yaml.safe_load(f)

def parse_volume(vol_str):
    """Parse volume text like '240 million transactions' into a number."""
    if pd.isna(vol_str):
        return 0
    vol_str = str(vol_str).lower().replace(',', '')
    match = re.search(r'([\d.]+)\s*(million|billion|k)?', vol_str)
    if not match:
        return 0
    num = float(match.group(1))
    mult = match.group(2)
    if mult == 'million':
        num *= 1_000_000
    elif mult == 'billion':
        num *= 1_000_000_000
    elif mult == 'k':
        num *= 1_000
    return num

def score_text(text, scale_dict):
    """Score text against a keyword-to-score mapping. Returns highest match."""
    if pd.isna(text):
        return 0
    text_lower = str(text).lower()
    best_score = 0
    for score, keywords in scale_dict.items():
        for kw in keywords:
            if kw.lower() in text_lower:
                best_score = max(best_score, int(score))
    return best_score

def evaluate_vendor(row, rubric):
    """Evaluate a single vendor row and return (tier, rationale, depth)."""
    scores = {}

    data_class = str(row.get('Data Classification Accessed', ''))
    service = str(row.get('Service / Product Provided', ''))
    biz_process = str(row.get('Business Process Supported', ''))
    category = str(row.get('Category', ''))
    op_dep = str(row.get('Operational Dependency', ''))
    vol_raw = row.get('Data Volume (annual)', '')

    # 1. Data Sensitivity (0-3)
    scores['data_sensitivity'] = score_text(data_class, rubric['scales']['data_sensitivity'])

    # 2. Payment Flow Involvement (0-3)
    combined_service = f"{service} {biz_process}"
    scores['payment_flow'] = score_text(combined_service, rubric['scales']['payment_flow'])

    # 3. Regulatory Exposure (0-3)
    combined_reg = f"{category} {service} {biz_process}"
    scores['regulatory_exposure'] = score_text(combined_reg, rubric['scales']['regulatory_exposure'])

    # 4. Operational Dependency (0-3)
    #    NOTE: "Critical" in the input is outside the Field Guide scale
    #    (Total/High/Moderate/Low). We treat it as >= Total (score 3).
    scores['operational_dependency'] = score_text(op_dep, rubric['scales']['operational_dependency'])

    # 5. Data Volume (0-3)
    vol = parse_volume(vol_raw)
    vol_score = 0
    if vol >= rubric['scales']['data_volume'][3]:
        vol_score = 3
    elif vol >= rubric['scales']['data_volume'][2]:
        vol_score = 2
    elif vol > rubric['scales']['data_volume'][1]:
        vol_score = 1
    scores['data_volume'] = vol_score

    total_score = sum(scores.values())

    # Privileged-access modifier
    modifier_applied = False
    if any(kw.lower() in data_class.lower() for kw in rubric['modifiers']['privileged_access']):
        total_score += 1
        modifier_applied = True

    # Tier mapping from thresholds
    tier = "Low"
    if total_score >= rubric['thresholds']['Critical']:
        tier = "Critical"
    elif total_score >= rubric['thresholds']['High']:
        tier = "High"
    elif total_score >= rubric['thresholds']['Medium']:
        tier = "Medium"

    # Floor rules
    floor_note = ""
    for kw in rubric['floors']['high']:
        if kw.lower() in data_class.lower() and tier in ("Low", "Medium"):
            tier = "High"
            floor_note = f" Floor rule applied: '{kw}' in data classification forces at least High."

    # Operational dependency override:
    # "Critical" in the input is outside the Field Guide scale. Assumption recorded.
    if 'Critical' in op_dep:
        if tier in ("Low", "Medium", "High"):
            tier = "Critical"
            floor_note += (" Assumption: Operational Dependency value 'Critical' is outside the "
                           "Field Guide scale (Total/High/Moderate/Low); treated as >= Total.")

    depth = rubric['depth_policy'].get(tier, "")

    # Build human-readable rationale in the style of V-000
    rationale_parts = [
        f"Data sensitivity scored {scores['data_sensitivity']}/3 "
        f"(data accessed: \"{data_class}\").",

        f"Payment-flow involvement scored {scores['payment_flow']}/3 "
        f"(service: \"{service}\"; process: \"{biz_process}\").",

        f"Regulatory exposure scored {scores['regulatory_exposure']}/3 "
        f"(category: \"{category}\").",

        f"Operational dependency scored {scores['operational_dependency']}/3 "
        f"(dependency: \"{op_dep}\").",

        f"Data volume scored {scores['data_volume']}/3 "
        f"(volume: \"{vol_raw}\").",
    ]

    if modifier_applied:
        rationale_parts.append(
            "Privileged-access modifier applied (+1): vendor holds production "
            "credentials despite limited customer data access."
        )

    rationale_parts.append(
        f"Total weighted score: {total_score}. Tier assigned: {tier}."
    )

    if floor_note:
        rationale_parts.append(floor_note.strip())

    rationale_parts.append(
        f"Assessment depth warranted: {depth}."
    )

    rationale = " ".join(rationale_parts)
    return tier, rationale, depth


def main():
    rubric = load_rubric('config/rubric.yaml')
    df = pd.read_excel('data/Meridian_Vendor_Input.xlsx', sheet_name='Vendor Inventory', header=3)
    df = df[df['Vendor ID'].str.startswith('V-', na=False) & (df['Vendor ID'] != 'V-000')]

    for idx, row in df.iterrows():
        tier, rationale, depth = evaluate_vendor(row, rubric)
        print(f"\n{'='*60}")
        print(f"{row['Vendor ID']} — {row['Vendor Name']}")
        print(f"Tier: {tier}")
        print(f"Rationale: {rationale}")
        print(f"Depth: {depth}")


if __name__ == '__main__':
    main()
