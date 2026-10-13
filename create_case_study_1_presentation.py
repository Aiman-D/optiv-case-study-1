"""
Case Study 1 Master Presentation Generator
Generates an executive, visually flawless 14-slide PowerPoint presentation exclusively for:
Optiv Case Study 1: Vendor AI Usage Analysis (Meridian Financial)

Presenter: Aiman Dhanani (Reg No: 24BYB0033)
Client: Meridian Financial
Engagement: Optiv Consulting • Cyber Strategy Practice • September 2026
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# ── Color Palette ─────────────────────────────────────────────────────────────
NAVY_DARK = RGBColor(11, 25, 44)       # #0B192C - Deep background / primary
NAVY_HEADER = RGBColor(20, 43, 76)     # #142B4C - Optiv classic navy
ORANGE = RGBColor(237, 107, 45)        # #ED6B2D - Optiv brand accent orange
WHITE = RGBColor(255, 255, 255)        # #FFFFFF - Pure white
GRAY_BG = RGBColor(248, 250, 252)      # #F8FAFC - Canvas light gray
CARD_BG = RGBColor(255, 255, 255)      # #FFFFFF - Card container bg
CARD_BORDER = RGBColor(218, 225, 235)  # #DAE1EB - Crisp subtle card border
TEXT_DARK = RGBColor(15, 23, 42)       # #0F172A - Deep slate for readability
TEXT_MUTED = RGBColor(100, 116, 139)   # #64748B - Secondary subtitle text

# Status colors
GREEN_ACCENT = RGBColor(16, 185, 129)  # #10B981 - Low risk / No AI
RED_ACCENT = RGBColor(239, 68, 68)     # #EF4444 - Critical risk / Yes
AMBER_ACCENT = RGBColor(245, 158, 11)  # #F59E0B - High risk / Inconclusive
BLUE_ACCENT = RGBColor(59, 130, 246)   # #3B82F6 - Medium risk


def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Helper: Base slide layout
    def setup_slide(is_dark=False, category_text="OPTIV CONSULTING  |  CYBER STRATEGY PRACTICE"):
        slide = prs.slides.add_slide(blank_layout)

        # Background fill
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = NAVY_DARK if is_dark else GRAY_BG
        bg.line.fill.background()

        # Top corner bracket accent
        bracket = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.32), Inches(0.18), Inches(0.18))
        bracket.fill.solid()
        bracket.fill.fore_color.rgb = ORANGE
        bracket.line.fill.background()

        # Header category label
        cat_box = slide.shapes.add_textbox(Inches(1.08), Inches(0.28), Inches(8.0), Inches(0.28))
        tf = cat_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = category_text
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = ORANGE if is_dark else NAVY_HEADER

        # Right header branding
        brand_box = slide.shapes.add_textbox(Inches(8.5), Inches(0.28), Inches(4.033), Inches(0.28))
        tf_b = brand_box.text_frame
        tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
        p_b = tf_b.paragraphs[0]
        p_b.text = "Optiv Consulting  |  Meridian Financial"
        p_b.alignment = PP_ALIGN.RIGHT
        p_b.font.size = Pt(9.5)
        p_b.font.bold = True
        p_b.font.color.rgb = WHITE if is_dark else NAVY_HEADER

        # Running Footer
        ft_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.08), Inches(11.733), Inches(0.25))
        tf_ft = ft_box.text_frame
        tf_ft.margin_left = tf_ft.margin_top = tf_ft.margin_right = tf_ft.margin_bottom = 0
        p_ft = tf_ft.paragraphs[0]
        p_ft.text = "Confidential  |  Aiman Dhanani (24BYB0033)  |  Case Study 1: Vendor AI Usage Analysis  |  September 2026"
        p_ft.font.size = Pt(8.5)
        p_ft.font.color.rgb = RGBColor(148, 163, 184)

        return slide

    def add_header(slide, title, subtitle="", is_dark=False):
        """Header with dedicated non-overlapping coordinate bands"""
        # Title text
        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.58), Inches(11.733), Inches(0.45))
        tf_t = tb_title.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
        p_t = tf_t.paragraphs[0]
        p_t.text = title
        p_t.font.size = Pt(21)
        p_t.font.bold = True
        p_t.font.color.rgb = WHITE if is_dark else NAVY_HEADER

        # Subtitle text
        if subtitle:
            tb_sub = slide.shapes.add_textbox(Inches(0.8), Inches(1.08), Inches(11.733), Inches(0.32))
            tf_s = tb_sub.text_frame
            tf_s.word_wrap = True
            tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0
            p_s = tf_s.paragraphs[0]
            p_s.text = subtitle
            p_s.font.size = Pt(11)
            p_s.font.color.rgb = RGBColor(203, 213, 225) if is_dark else TEXT_MUTED

        # Accent bar positioned at 1.46 (below subtitle, well above cards starting at 1.68)
        accent_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.46), Inches(1.5), Inches(0.04))
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = ORANGE
        accent_bar.line.fill.background()

    def add_card(slide, left, top, width, height, title="", title_color=NAVY_HEADER, bg_color=CARD_BG, border_color=CARD_BORDER):
        """Clean rectangular card with colored top accent border and dedicated title"""
        card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.2)

        if title:
            # Top accent bar on card
            strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(0.05))
            strip.fill.solid()
            strip.fill.fore_color.rgb = title_color if title_color != NAVY_HEADER else ORANGE
            strip.line.fill.background()

            # Dedicated title box
            tb = slide.shapes.add_textbox(Inches(left + 0.22), Inches(top + 0.15), Inches(width - 0.44), Inches(0.32))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(12)
            p.font.bold = True
            p.font.color.rgb = title_color

        return card

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 1: Title Slide (Optiv Navy)
    # ──────────────────────────────────────────────────────────────────────────
    s1 = setup_slide(is_dark=True, category_text="CONFIDENTIAL  |  OPTIV CONSULTING  |  CYBER STRATEGY")

    # Kicker / Badge
    tb_kicker = s1.shapes.add_textbox(Inches(0.8), Inches(1.45), Inches(11.733), Inches(0.35))
    tf_k = tb_kicker.text_frame
    tf_k.margin_left = tf_k.margin_top = tf_k.margin_right = tf_k.margin_bottom = 0
    p_k = tf_k.paragraphs[0]
    p_k.text = "CASE STUDY 01  •  FINAL EXECUTIVE DELIVERABLE"
    p_k.font.size = Pt(13)
    p_k.font.bold = True
    p_k.font.color.rgb = ORANGE

    # Main Title Box
    tb_main = s1.shapes.add_textbox(Inches(0.8), Inches(1.90), Inches(11.733), Inches(0.75))
    tf_m = tb_main.text_frame
    tf_m.margin_left = tf_m.margin_top = tf_m.margin_right = tf_m.margin_bottom = 0
    p_m = tf_m.paragraphs[0]
    p_m.text = "Vendor AI Usage Analysis"
    p_m.font.size = Pt(40)
    p_m.font.bold = True
    p_m.font.color.rgb = WHITE

    # Accent divider bar
    sbar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.78), Inches(2.2), Inches(0.06))
    sbar.fill.solid()
    sbar.fill.fore_color.rgb = ORANGE
    sbar.line.fill.background()

    # Subtitle Box
    tb_sub1 = s1.shapes.add_textbox(Inches(0.8), Inches(3.02), Inches(11.733), Inches(0.55))
    tf_s1 = tb_sub1.text_frame
    tf_s1.word_wrap = True
    tf_s1.margin_left = tf_s1.margin_top = tf_s1.margin_right = tf_s1.margin_bottom = 0
    p_s1 = tf_s1.paragraphs[0]
    p_s1.text = "An Evidence-Driven OSINT Methodology & Security Risk Engine for Third-Party AI Due Diligence"
    p_s1.font.size = Pt(17)
    p_s1.font.color.rgb = RGBColor(226, 232, 240)

    # Presenter Information Card
    meta = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(3.95), Inches(11.733), Inches(2.75))
    meta.fill.solid()
    meta.fill.fore_color.rgb = RGBColor(17, 34, 64)
    meta.line.color.rgb = RGBColor(30, 58, 102)
    meta.line.width = Pt(1.5)

    # Top accent strip on card
    meta_strip = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(3.95), Inches(11.733), Inches(0.06))
    meta_strip.fill.solid()
    meta_strip.fill.fore_color.rgb = ORANGE
    meta_strip.line.fill.background()

    mt_box = s1.shapes.add_textbox(Inches(1.15), Inches(4.20), Inches(11.0), Inches(2.35))
    mt_tf = mt_box.text_frame
    mt_tf.word_wrap = True
    mt_tf.margin_left = mt_tf.margin_top = mt_tf.margin_right = mt_tf.margin_bottom = 0

    p_m1 = mt_tf.paragraphs[0]
    p_m1.text = "PRESENTER & ENGAGEMENT METADATA"
    p_m1.font.size = Pt(11.5)
    p_m1.font.bold = True
    p_m1.font.color.rgb = ORANGE

    p_pres = mt_tf.add_paragraph()
    p_pres.text = "• Presenter: Aiman Dhanani  |  Registration Number: 24BYB0033"
    p_pres.font.size = Pt(15.5)
    p_pres.font.bold = True
    p_pres.font.color.rgb = WHITE
    p_pres.space_before = Pt(6)

    p_m2 = mt_tf.add_paragraph()
    p_m2.text = "• Institution: VIT  |  Optiv Consulting Cyber Strategy Campus Engagement\n" \
               "• Client Engagement: Meridian Financial Third-Party Risk Management (TPRM)\n" \
               "• Core Deliverables: Meridian_Vendor_Output.xlsx, Interactive Streamlit UI, 18/18 Unit Tests\n" \
               "• Source Code Repository: https://github.com/Aiman-D/optiv-case-study-1"
    p_m2.font.size = Pt(11.5)
    p_m2.font.color.rgb = RGBColor(203, 213, 225)
    p_m2.space_before = Pt(4)

    s1.notes_slide.notes_text_frame.text = (
        "Welcome everyone. My name is Aiman Dhanani (Reg No: 24BYB0033). Today I am presenting our end-to-end "
        "deliverable for Case Study 1: Vendor AI Usage Analysis. We were engaged by Meridian Financial to build "
        "a repeatable, evidence-driven capability that inspects third-party vendors' public footprints to determine "
        "whether they genuinely deploy Artificial Intelligence in the services provided to Meridian."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 2: Executive Summary & Client Context
    # ──────────────────────────────────────────────────────────────────────────
    s2 = setup_slide(is_dark=False)
    add_header(s2, "Executive Summary & Engagement Context", "Meridian Financial Third-Party Risk Assessment Initiative")

    # Card 1: Context & Problem
    add_card(s2, 0.8, 1.68, 5.7, 5.15, "THE BUSINESS CONTEXT & THREAT LANDSCAPE", NAVY_HEADER)
    c1_box = s2.shapes.add_textbox(Inches(1.05), Inches(2.22), Inches(5.2), Inches(4.4))
    c1_tf = c1_box.text_frame
    c1_tf.word_wrap = True
    c1_tf.margin_left = c1_tf.margin_top = c1_tf.margin_right = c1_tf.margin_bottom = 0

    c1_items = [
        ("Client Profile:", "Meridian Financial is a global financial institution serving 1.4 million customers, handling critical payments, account balances, and wealth management portfolios."),
        ("The Estate:", "Meridian relies heavily on external vendors across core banking, automated workload scheduling, statement generation, payment settlement, and wealth reporting."),
        ("The Challenge:", "Commercial software vendors rapidly embed generative and predictive AI into their products without formal disclosures to enterprise clients."),
        ("The Threat:", "If a vendor feeds customer data or transaction streams into unvetted AI models, Meridian incurs severe regulatory, data leakage, and contractual compliance liabilities."),
        ("The Optiv Mandate:", "Design an automated, reproducible method to inspect vendors' public footprints and determine true service-level AI usage.")
    ]
    for h, d in c1_items:
        bp = c1_tf.add_paragraph() if c1_tf.paragraphs[0].text else c1_tf.paragraphs[0]
        bp.text = f"▪ {h} {d}"
        bp.font.size = Pt(10)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(6)

    # Card 2: Solution & Deliverables
    add_card(s2, 6.833, 1.68, 5.7, 5.15, "THE OPTIV METHODOLOGY & CORE DELIVERABLES", ORANGE)
    c2_box = s2.shapes.add_textbox(Inches(7.08), Inches(2.22), Inches(5.2), Inches(4.4))
    c2_tf = c2_box.text_frame
    c2_tf.word_wrap = True
    c2_tf.margin_left = c2_tf.margin_top = c2_tf.margin_right = c2_tf.margin_bottom = 0

    c2_items = [
        ("100% OSINT-Driven:", "Zero vendor contact, zero scraping of authenticated portals, full adherence to robots.txt and polite 1.2s rate limits."),
        ("Deterministic Criticality Engine:", "Objective 0–3 scoring rubric across 5 dimensions, removing analyst bias and setting proportionate assessment depth."),
        ("Cryptographic Quote Guard:", "Enforces exact substring matching against immutable cached snapshots. Any hallucinated LLM claim is discarded immediately."),
        ("Definitive Classification:", "Resolved Meridian's 6 vendors down to 3 Confirmed Yes (Fiserv, BNY, AutomWorx), 2 Confirmed No (FSSI, TCH), and exactly 1 Inconclusive (Terrapin)."),
        ("Client-Ready Deliverable:", "Preserved Columns B–K byte-for-byte; added 132-row Evidence Log sheet; deployed interactive Streamlit app in Presentation Mode.")
    ]
    for h, d in c2_items:
        bp = c2_tf.add_paragraph() if c2_tf.paragraphs[0].text else c2_tf.paragraphs[0]
        bp.text = f"▪ {h} {d}"
        bp.font.size = Pt(10)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(6)

    s2.notes_slide.notes_text_frame.text = (
        "Here is the executive summary. Meridian Financial handles 1.4 million customers and critical payment flows. "
        "Our mission was to cut through vendor marketing buzz and establish whether AI is genuinely used in the "
        "services Meridian buys. We delivered a complete pipeline that classifies all 6 vendors with 100% auditability."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 3: Problem Statement: Reading the Public Footprint
    # ──────────────────────────────────────────────────────────────────────────
    s3 = setup_slide(is_dark=False)
    add_header(s3, "Problem Statement: Reading the Public Footprint", "Navigating Ethical Boundaries, Marketing Noise, and Cryptographic Traceability")

    cols3 = [
        ("ETHICAL OSINT BOUNDARIES", [
            ("Strictly Public Footprint:", "Analysis relies 100% on publicly accessible information: product docs, trust centers, job listings, and regulatory filings."),
            ("No Active Vendor Contact:", "Zero vendor outreach or questionnaire distribution during discovery; eliminates early tipping-off or non-standard disclosures."),
            ("Robots.txt & Rate Limiting:", "Adheres to RFC 9309 standards; enforces a strict 1.2s delay between requests; never bypasses paywalls or CAPTCHAs.")
        ], NAVY_HEADER),
        ("NOISE VS. SERVICE LINKAGE", [
            ("Marketing Buzzword Clutter:", "Vendors frequently publish speculative blog posts on 'AI Trends' without embedding AI into actual commercial deliverables."),
            ("Service-Linkage Requirement:", "Evidence must specifically tie to the service Meridian purchases (e.g. DNA core banking) rather than company-wide HR tools."),
            ("Evidentiary Weight Hierarchy:", "Direct technical documentation and sub-processor registers carry 4x the weight of generic executive interviews.")
        ], ORANGE),
        ("TRACEABILITY & INTEGRITY", [
            ("Immutable Snapshot Store:", "Every single claim traces to a local snapshot with URL, ISO-8601 retrieval timestamp, and SHA-256 cryptographic hash."),
            ("Quote Guard Substring Check:", "Candidate quotes extracted by the LLM are validated against the raw text. Non-verbatim claims are marked INVALID_QUOTE."),
            ("Coverage Gap Accounting:", "Unreachable sites or bot-blocked pages are logged as COVERAGE-GAP—never falsely reported as 'No' AI usage.")
        ], NAVY_HEADER)
    ]

    for idx, (title, items, col) in enumerate(cols3):
        left = 0.8 + idx * 4.0
        add_card(s3, left, 1.68, 3.75, 5.15, title, col)
        tb = s3.shapes.add_textbox(Inches(left + 0.22), Inches(2.22), Inches(3.31), Inches(4.4))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        for h, d in items:
            p = tf.add_paragraph() if tf.paragraphs[0].text else tf.paragraphs[0]
            p.text = f"▪ {h}"
            p.font.size = Pt(10.5)
            p.font.bold = True
            p.font.color.rgb = col
            p.space_before = Pt(6)

            p_d = tf.add_paragraph()
            p_d.text = d
            p_d.font.size = Pt(9.5)
            p_d.font.color.rgb = TEXT_DARK
            p_d.space_before = Pt(2)

    s3.notes_slide.notes_text_frame.text = (
        "This slide covers the problem statement from Slide 3 of the prompt. We operated under strict ethical rules: "
        "no vendor contact, no bypassing paywalls. We also distinguished between marketing hype and true service-level usage, "
        "and enforced strict substring matching so every finding is backed by stored evidence."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 4: Roles & Responsibilities Framework
    # ──────────────────────────────────────────────────────────────────────────
    s4 = setup_slide(is_dark=False)
    add_header(s4, "Roles & Responsibilities: 5-Pillar Architecture", "The Structured Methodology Answering the Core Engagement Prompts")

    pillars = [
        ("01", "DETERMINE CRITICALITY", [
            ("Core Mandate:", "Judge how critical each vendor is from the profile provided, establishing objective, explainable criteria."),
            ("Evaluation Anchors:", "5 anchored dimensions (0–3 scale): Data sensitivity, payment-flow involvement, regulatory exposure, operational dependency, data volume."),
            ("Floor & Modifiers:", "Applies privileged access modifiers (+1) and customer master floor rules to ensure reproducible ratings across analysts.")
        ], NAVY_HEADER),
        ("02", "SET THE DEPTH", [
            ("Core Mandate:", "Decide proportionate assessment depth warranted by criticality, optimizing investigative resources."),
            ("Tier Policies:", "Critical warrants a full sweep (docs, trust center, technical jobs, filings); Low warrants trust center/policy check only."),
            ("Stop Conditions:", "Crawl ceases once confirmed signals exist and model providers are identified or confirmed unfindable.")
        ], ORANGE),
        ("03", "IDENTIFY SOURCES", [
            ("Core Mandate:", "Identify and prioritize OSINT sources that reveal genuine service AI usage vs marketing fluff."),
            ("Source Hierarchy:", "Tier A (Direct product docs, sub-processors, policies) down to Tier D (Generic thought leadership noise)."),
            ("Weighting Logic:", "Applies service-linkage multipliers (1.5x product vs 0.5x company-wide) and 180-day recency discounts (0.7x).")
        ], NAVY_HEADER),
        ("04", "BUILD CAPABILITY", [
            ("Core Mandate:", "Build automated scrapers, content clean-up, LLM signal extraction, and rigorous verification."),
            ("Technical Stack:", "Trafilatura content extraction, regex pre-filtering, Groq LPU API (gpt-oss-120b) with strict JSON schema."),
            ("Cryptographic Guard:", "Whitespace-normalized exact substring verification guarantees zero hallucinations enter the audit log.")
        ], ORANGE),
        ("05", "SURFACE THE RISK", [
            ("Core Mandate:", "Translate confirmed AI signals into concrete enterprise security risks to support TPRM due diligence."),
            ("Risk Taxonomy:", "Maps to 4 enterprise risks: sensitive data ingestion, shadow AI sub-processors, credential leakage, contract voids."),
            ("Governance Action:", "Drafts actionable 15-day vendor questionnaires, sub-processor registry updates, and 30-day escalation paths.")
        ], NAVY_HEADER)
    ]

    for idx, (p_num, p_title, p_sections, p_col) in enumerate(pillars):
        left = 0.8 + idx * 2.39
        add_card(s4, left, 1.68, 2.25, 5.15, f"PILLAR {p_num}", p_col)

        tb = s4.shapes.add_textbox(Inches(left + 0.16), Inches(2.22), Inches(1.93), Inches(4.45))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p = tf.paragraphs[0]
        p.text = p_title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = p_col

        for sub_h, sub_d in p_sections:
            p_h = tf.add_paragraph()
            p_h.text = f"▪ {sub_h}"
            p_h.font.size = Pt(9.5)
            p_h.font.bold = True
            p_h.font.color.rgb = p_col
            p_h.space_before = Pt(6)

            p_b = tf.add_paragraph()
            p_b.text = sub_d
            p_b.font.size = Pt(8.5)
            p_b.font.color.rgb = TEXT_DARK
            p_b.space_before = Pt(2)

    s4.notes_slide.notes_text_frame.text = (
        "Here is the 5-pillar roles and responsibilities framework directly reflecting Slide 4 of the briefing document. "
        "We systematically tackled each responsibility: scoring criticality, setting proportionate depth, identifying OSINT tiers, "
        "building the software capability, and surfacing tangible operational risks to Meridian."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 5: Expected Outcome 1 — Functional Design Diagram
    # ──────────────────────────────────────────────────────────────────────────
    s5 = setup_slide(is_dark=False)
    add_header(s5, "Expected Outcome 1: Functional Design Diagram", "End-to-End Pipeline Data Flow from OSINT Collection to Final Risk Workbook")

    stages = [
        ("M1: Criticality", [
            ("Inputs:", "Meridian_Vendor_Input.xlsx, rubric.yaml"),
            ("Processing:", "5 anchored scales (0–3), floor rules, +1 modifier"),
            ("Output:", "Criticality Tier & Warranted Depth"),
            ("Tech Stack:", "Python 3.12, pandas, PyYAML")
        ], NAVY_HEADER),
        ("M2: Collectors", [
            ("Inputs:", "Vendor domains, careers portals, sitemaps"),
            ("Processing:", "Robots.txt parsing, polite 1.2s delay, trafilatura"),
            ("Output:", "SHA-256 hashed raw & cleaned text snapshots"),
            ("Tech Stack:", "urllib.robotparser, trafilatura, bs4")
        ], ORANGE),
        ("M3: Classifier", [
            ("Inputs:", "Cleaned page snapshots, sources.yaml"),
            ("Processing:", "25+ regex filters, Groq LPU JSON inference"),
            ("Output:", "Verbatim quotes, AI mechanism, named providers"),
            ("Tech Stack:", "Groq API, gpt-oss-120b, Quote Guard")
        ], NAVY_HEADER),
        ("M4: Risk Engine", [
            ("Inputs:", "Verified signals, vendor criticality, data class"),
            ("Processing:", "Multi-factor formula, 180-day decay, coverage rules"),
            ("Output:", "AI Verdict (Y/N/Inc), Security Risk Classification"),
            ("Tech Stack:", "Mathematical scoring engine, coverage logger")
        ], ORANGE),
        ("M5: Excel & UI", [
            ("Inputs:", "Final classifications, evidence audit trail"),
            ("Processing:", "Preserves B–K byte-for-byte, writes L–V, UI cache"),
            ("Output:", "Meridian_Vendor_Output.xlsx, Streamlit app"),
            ("Tech Stack:", "openpyxl, Streamlit, pytest (18/18 pass)")
        ], NAVY_HEADER)
    ]

    for i, (stitle, sitems, scol) in enumerate(stages):
        left = 0.8 + i * 2.4
        box = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(1.68), Inches(2.25), Inches(3.2))
        box.fill.solid()
        box.fill.fore_color.rgb = CARD_BG
        box.line.color.rgb = scol
        box.line.width = Pt(1.5)

        # Header bar
        hbar = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(1.68), Inches(2.25), Inches(0.42))
        hbar.fill.solid()
        hbar.fill.fore_color.rgb = scol
        hbar.line.fill.background()

        tb_h = s5.shapes.add_textbox(Inches(left), Inches(1.73), Inches(2.25), Inches(0.32))
        p_h = tb_h.text_frame.paragraphs[0]
        p_h.text = f"Stage 0{i+1}"
        p_h.alignment = PP_ALIGN.CENTER
        p_h.font.size = Pt(10.5)
        p_h.font.bold = True
        p_h.font.color.rgb = WHITE

        tb = s5.shapes.add_textbox(Inches(left + 0.12), Inches(2.18), Inches(2.01), Inches(2.65))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = stitle
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = scol

        for ih, idesc in sitems:
            bp_h = tf.add_paragraph()
            bp_h.text = f"• {ih}"
            bp_h.font.size = Pt(8.5)
            bp_h.font.bold = True
            bp_h.font.color.rgb = TEXT_DARK
            bp_h.space_before = Pt(3)

            bp_d = tf.add_paragraph()
            bp_d.text = f"  {idesc}"
            bp_d.font.size = Pt(8.0)
            bp_d.font.color.rgb = TEXT_MUTED
            bp_d.space_before = Pt(1)

        if i < 4:
            arrow = s5.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(left + 2.27), Inches(3.05), Inches(0.11), Inches(0.18))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = ORANGE
            arrow.line.fill.background()

    # Bottom Data Flow Architecture Card
    add_card(s5, 0.8, 5.08, 11.733, 1.75, "DATA FLOW & REPRODUCIBILITY GUARANTEES", NAVY_HEADER)
    df_box = s5.shapes.add_textbox(Inches(1.05), Inches(5.55), Inches(11.2), Inches(1.15))
    df_tf = df_box.text_frame
    df_tf.word_wrap = True
    df_tf.margin_left = df_tf.margin_top = df_tf.margin_right = df_tf.margin_bottom = 0

    df_p1 = df_tf.paragraphs[0]
    df_p1.text = "▪ Input Stream: data/Meridian_Vendor_Input.xlsx (Client Source of Truth, 100% Immutable) | config/rubric.yaml | config/sources.yaml"
    df_p1.font.size = Pt(9.5)
    df_p1.font.color.rgb = TEXT_DARK

    df_p2 = df_tf.add_paragraph()
    df_p2.text = "▪ Intermediate Snapshot Store: cache/<Vendor_ID>/<URL_Hash>.json (Raw HTML, Cleaned Text, Retrieval Date, SHA-256 Content Hash)"
    df_p2.font.size = Pt(9.5)
    df_p2.font.color.rgb = TEXT_DARK
    df_p2.space_before = Pt(4)

    df_p3 = df_tf.add_paragraph()
    df_p3.text = "▪ Output Deliverables: data/Meridian_Vendor_Output.xlsx (Preserved B–K, Populated L–V, 132-row Evidence Log) | Streamlit Presentation UI"
    df_p3.font.size = Pt(9.5)
    df_p3.font.color.rgb = TEXT_DARK
    df_p3.space_before = Pt(4)

    s5.notes_slide.notes_text_frame.text = (
        "Expected Outcome 1: Here is our Functional Design Diagram showing data progression. Every module is decoupled. "
        "M1 parses the input offline; M2 crawls and caches snapshots; M3 runs keyword filtering and LLM classification; "
        "M4 computes the weighted risk score; M5 populates the final workbook and drives the Streamlit demo."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 6: Expected Outcome 2 — System Architecture Diagram
    # ──────────────────────────────────────────────────────────────────────────
    s6 = setup_slide(is_dark=False)
    add_header(s6, "Expected Outcome 2: System Architecture Diagram", "Technical Components, Interactions, Libraries, and Internal Data Structures")

    tech_comps = [
        ("MODULE 1 & 2: INGESTION & CRAWL", [
            ("src/criticality.py", "Parses input via pandas, executes 5 anchored rubric scales, applies floor rules and privileged access modifiers (+1)."),
            ("src/collectors.py", "Automated link discovery via sitemap.xml and homepages; respects robots.txt (urllib.robotparser); 1.2s delay."),
            ("trafilatura & bs4", "Strips scripts, styling, and navigation boilerplate. Generates standardized cleaned text for audit indexing."),
            ("SHA-256 Snapshot Store", "Persists local JSON cache containing raw HTML, cleaned text, retrieval timestamp, and content hash for zero-contact replay.")
        ], NAVY_HEADER),
        ("MODULE 3: CLASSIFIER & QUOTE GUARD", [
            ("src/classify.py", "Pre-filters text using 25+ regex patterns for AI keywords to drop irrelevant pages and conserve tokens."),
            ("Groq LPU API Engine", "Leverages ultra-fast cloud inference (openai/gpt-oss-120b) with strict JSON schema mode (signal_type, quote, provider)."),
            ("verify_quote() Guard", "Normalizes whitespace and executes exact substring check. Any non-verbatim quote is discarded as INVALID_QUOTE."),
            ("Signal Scoring Logic", "Calculates base evidence score across CONFIRMED, CORROBORATED, INDICATIVE, and MARKETING-ONLY tags.")
        ], ORANGE),
        ("MODULE 4 & 5: RISK & EXPORT", [
            ("src/risk_export.py", "Computes multi-factor evidence score, applies 180-day recency discount (0.7x), maps tier & footprint to risk."),
            ("openpyxl writer", "Uses shutil.copy to preserve client workbook formatting; populates columns L–V; appends 132-entry Evidence Log."),
            ("app/main.py (Streamlit)", "Provides instant Presentation Mode, pre-computed KPI cards, per-vendor deep-dive cards, and one-click download."),
            ("pytest Test Suite", "18 automated unit tests verifying deterministic criticality, Quote Guard integrity, and byte-for-byte column preservation.")
        ], NAVY_HEADER)
    ]

    for idx, (c_title, c_items, c_col) in enumerate(tech_comps):
        left = 0.8 + idx * 4.0
        add_card(s6, left, 1.68, 3.75, 5.15, c_title, c_col)

        tb = s6.shapes.add_textbox(Inches(left + 0.22), Inches(2.22), Inches(3.31), Inches(4.45))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        for mod, desc in c_items:
            p = tf.add_paragraph() if tf.paragraphs[0].text else tf.paragraphs[0]
            p.text = f"▪ {mod}"
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = c_col
            p.space_before = Pt(5)

            p_d = tf.add_paragraph()
            p_d.text = desc
            p_d.font.size = Pt(9.0)
            p_d.font.color.rgb = TEXT_DARK
            p_d.space_before = Pt(1)

    s6.notes_slide.notes_text_frame.text = (
        "Expected Outcome 2: Here is the detailed system architecture. We show the exact Python modules: "
        "criticality.py, collectors.py, classify.py, and risk_export.py. The Quote Guard is highlighted in the center: "
        "it sits between the LLM output and the aggregation engine to guarantee zero hallucinated evidence."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 7: Expected Outcome 3 — Source Strategy & Evidentiary Weight Logic
    # ──────────────────────────────────────────────────────────────────────────
    s7 = setup_slide(is_dark=False)
    add_header(s7, "Expected Outcome 3: Source Strategy & Evidentiary Logic", "Tiered Reliability Hierarchy, Mathematical Weighting Formula, and the Quote Guard")

    # Left: 4 Source Tiers Table
    add_card(s7, 0.8, 1.68, 6.2, 5.15, "OSINT SOURCE RELIABILITY TIERS", NAVY_HEADER)

    t_shape7 = s7.shapes.add_table(5, 3, Inches(1.0), Inches(2.22), Inches(5.8), Inches(4.35))
    tbl7 = t_shape7.table
    tbl7.columns[0].width = Inches(1.55)
    tbl7.columns[1].width = Inches(3.05)
    tbl7.columns[2].width = Inches(1.2)

    h7 = ["Tier & Class", "Source Type & Footprint", "Weight"]
    for i, h in enumerate(h7):
        cell = tbl7.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY_HEADER
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = WHITE

    tier_rows = [
        ("Tier A: Direct", "Product docs, sub-processor lists, AI policies, official regulatory filings.", "4.0x (Highest)"),
        ("Tier B: Corroborating", "Technical job postings, engineering blogs, GitHub repos, developer SDKs.", "3.0x (Strong)"),
        ("Tier C: Contextual", "Executive interviews, press releases, partnership announcements.", "2.0x (Moderate)"),
        ("Tier D: Noise", "Generic marketing blogs, buzzwords, speculative thought pieces.", "1.0x (Filtered)")
    ]
    for row_idx, (t, s, w) in enumerate(tier_rows, 1):
        for col_idx, val in enumerate([t, s, w]):
            cell = tbl7.cell(row_idx, col_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if row_idx % 2 == 1 else RGBColor(241, 245, 249)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(9)
            p.font.color.rgb = TEXT_DARK
            if col_idx == 0:
                p.font.bold = True

    # Right: Formula & Quote Guard Card
    add_card(s7, 7.25, 1.68, 5.28, 5.15, "MATHEMATICAL WEIGHTING & QUOTE GUARD", ORANGE)
    r_box7 = s7.shapes.add_textbox(Inches(7.48), Inches(2.22), Inches(4.8), Inches(4.4))
    r_tf7 = r_box7.text_frame
    r_tf7.word_wrap = True
    r_tf7.margin_left = r_tf7.margin_top = r_tf7.margin_right = r_tf7.margin_bottom = 0

    f_p7 = r_tf7.paragraphs[0]
    f_p7.text = "▪ Multi-Factor Weighting Formula:"
    f_p7.font.size = Pt(11.5)
    f_p7.font.bold = True
    f_p7.font.color.rgb = ORANGE

    f_desc7 = r_tf7.add_paragraph()
    f_desc7.text = "Signal Score = Base × Tier × Conf × Linkage × Recency\n" \
                   "• Base Score: CONFIRMED (5) | CORROBORATED (3) | INDICATIVE (2) | MARKETING (0.5)\n" \
                   "• Service Linkage: 1.5x if tied to Meridian's product; 0.5x if generic company-wide\n" \
                   "• Recency Discount: 0.7x penalty for signals > 180 days old"
    f_desc7.font.size = Pt(9.5)
    f_desc7.font.color.rgb = TEXT_DARK
    f_desc7.space_before = Pt(4)

    qg_p7 = r_tf7.add_paragraph()
    qg_p7.text = "▪ The Cryptographic Quote Guard:"
    qg_p7.font.size = Pt(11.5)
    qg_p7.font.bold = True
    qg_p7.font.color.rgb = NAVY_HEADER
    qg_p7.space_before = Pt(10)

    qg_desc7 = r_tf7.add_paragraph()
    qg_desc7.text = "1. LLM extracts candidate verbatim quote and AI mechanism.\n" \
                    "2. Python executes whitespace-normalized substring search.\n" \
                    "3. If candidate quote is NOT a 100% exact substring of the cached snapshot, it is discarded as INVALID_QUOTE.\n" \
                    "4. Guarantees zero hallucinations enter the audit trail."
    qg_desc7.font.size = Pt(9.5)
    qg_desc7.font.color.rgb = TEXT_DARK
    qg_desc7.space_before = Pt(4)

    s7.notes_slide.notes_text_frame.text = (
        "Expected Outcome 3: Source strategy and evidentiary weighting logic. A large company almost certainly "
        "uses AI somewhere (like internal HR). Our service-linkage multiplier (1.5x vs 0.5x) ensures we only convict "
        "a vendor if AI touches the actual product Meridian buys. And our Quote Guard eliminates LLM hallucinations."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 8: Expected Outcome 4 — Criticality Scoring Rubric & Depth Policy
    # ──────────────────────────────────────────────────────────────────────────
    s8 = setup_slide(is_dark=False)
    add_header(s8, "Expected Outcome 4: Criticality Scoring Rubric", "Deterministic 0–3 Scoring Rubric, Floor Rules, and Tier-Driven Depth Policy")

    add_card(s8, 0.8, 1.68, 5.7, 5.15, "RUBRIC EVALUATION DIMENSIONS (0–3 SCALE)", NAVY_HEADER)
    d_box8 = s8.shapes.add_textbox(Inches(1.05), Inches(2.22), Inches(5.2), Inches(4.4))
    d_tf8 = d_box8.text_frame
    d_tf8.word_wrap = True
    d_tf8.margin_left = d_tf8.margin_top = d_tf8.margin_right = d_tf8.margin_bottom = 0

    dim_items8 = [
        ("Data Sensitivity (0–3):", "0 = No data | 1 = Internal business data | 2 = Confidential records | 3 = Full customer master records"),
        ("Payment-Flow Involvement (0–3):", "0 = No involvement | 1 = Reporting/indirect | 2 = Payment file transmission | 3 = Core clearing and settlement rail"),
        ("Regulatory Exposure (0–3):", "0 = Standard services | 1 = Reporting | 2 = Regulated banking | 3 = Systemically critical payment utility"),
        ("Operational Dependency (0–3):", "0 = Low | 1 = Moderate | 2 = High | 3 = Critical / Total operational dependence"),
        ("Data Volume (0–3):", "0 = Not applicable | 1 = < 100k records | 2 = 100k – 5M records | 3 = > 5M records / messages")
    ]
    for h, d in dim_items8:
        bp = d_tf8.add_paragraph() if d_tf8.paragraphs[0].text else d_tf8.paragraphs[0]
        bp.text = f"▪ {h} {d}"
        bp.font.size = Pt(9.5)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(6)

    add_card(s8, 6.833, 1.68, 5.7, 5.15, "MODIFIERS & TIER-DRIVEN DEPTH POLICY", ORANGE)
    r_box8 = s8.shapes.add_textbox(Inches(7.08), Inches(2.22), Inches(5.2), Inches(4.4))
    r_tf8 = r_box8.text_frame
    r_tf8.word_wrap = True
    r_tf8.margin_left = r_tf8.margin_top = r_tf8.margin_right = r_tf8.margin_bottom = 0

    r_items8 = [
        ("Privileged-Access Modifier (+1):", "If vendor holds production credentials or job scheduling authority without direct customer data (e.g. AutomWorx), apply +1 score modifier."),
        ("Customer Master Floor Rule:", "Any vendor accessing full customer master data is guaranteed at least a High or Critical rating."),
        ("Critical Depth Policy:", "Full sweep: product docs, trust center, sub-processor list, AI policy, DPA pages, technical job postings, engineering blog, GitHub, filings."),
        ("High Depth Policy:", "Product docs, trust center, technical job postings, and regulatory filings."),
        ("Medium Depth Policy:", "Public website, standard product documentation, and general job listings."),
        ("Low Depth Policy:", "Trust center and AI-policy check only (efficient resource allocation).")
    ]
    for h, d in r_items8:
        bp = r_tf8.add_paragraph() if r_tf8.paragraphs[0].text else r_tf8.paragraphs[0]
        bp.text = f"▪ {h} {d}"
        bp.font.size = Pt(9.5)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(6)

    s8.notes_slide.notes_text_frame.text = (
        "Expected Outcome 4: Our criticality scoring engine lives in config/rubric.yaml. It is deterministic: "
        "another analyst evaluating the same vendor arrives at the exact same score. The depth policy dictates "
        "how far we crawl based on criticality, ensuring we don't over-investigate low-risk vendors."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 9: Expected Outcome 4 (Cont.) — Comprehensive Vendor Criticality Scorecard
    # ──────────────────────────────────────────────────────────────────────────
    s9 = setup_slide(is_dark=False)
    add_header(s9, "Expected Outcome 4 (Cont.): Vendor Criticality Scores", "Deterministic Scoring Breakdown Across Meridian's 6 Third-Party Vendors")

    t_shape9 = s9.shapes.add_table(7, 6, Inches(0.8), Inches(1.68), Inches(11.733), Inches(5.15))
    tbl9 = t_shape9.table
    tbl9.columns[0].width = Inches(1.0)
    tbl9.columns[1].width = Inches(2.6)
    tbl9.columns[2].width = Inches(1.0)
    tbl9.columns[3].width = Inches(1.3)
    tbl9.columns[4].width = Inches(1.4)
    tbl9.columns[5].width = Inches(4.433)

    h9 = ["Vendor ID", "Vendor Name", "Score", "Assigned Tier", "Dependency", "Criticality Rationale & Applied Depth"]
    for i, h in enumerate(h9):
        cell = tbl9.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY_HEADER
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = WHITE

    crit_rows = [
        ("V-001", "AutomWorx", "3", "Low", "High", "Holds production credentials (+1 privileged modifier) but no direct customer data. Depth: Trust center/AI policy only."),
        ("V-002", "Fiserv, Inc.", "15", "Critical", "Critical", "Core banking (DNA) & card platform accessing 1.4M customer master records. Depth: Full sweep (docs, trust, jobs, filings)."),
        ("V-003", "FSSI", "9", "High", "High", "Statement print & e-delivery processing for 850k customer accounts. Depth: Product docs, trust center, jobs, filings."),
        ("V-004", "Terrapin Technologies", "6", "Medium", "Moderate", "Wealth data solution consolidating 38k advisor accounts & commissions. Depth: Website, docs, and job listings."),
        ("V-005", "BNY", "9", "High", "High", "RTP transmission services processing 2.1M payment messages. Depth: Product docs, trust center, jobs, filings."),
        ("V-006", "The Clearing House", "12", "Critical", "Critical", "National ACH & RTP settlement clearing 14M payment messages. Depth: Full sweep (operating rules, specs, filings).")
    ]

    for row_idx, row in enumerate(crit_rows, 1):
        for col_idx, val in enumerate(row):
            cell = tbl9.cell(row_idx, col_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if row_idx % 2 == 1 else RGBColor(241, 245, 249)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(9)
            if col_idx == 3:
                p.font.bold = True
                p.font.color.rgb = RED_ACCENT if val == "Critical" else AMBER_ACCENT if val == "High" else BLUE_ACCENT if val == "Medium" else GREEN_ACCENT
            else:
                p.font.color.rgb = TEXT_DARK

    s9.notes_slide.notes_text_frame.text = (
        "Here is the complete scorecard for Meridian's 6 vendors. Notice how the rubric behaved: Fiserv and "
        "The Clearing House are Critical due to systemic importance; FSSI and BNY are High; Terrapin is Medium; "
        "and AutomWorx is Low despite holding credentials due to the +1 modifier."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 10: Expected Outcome 5 — Vendor Assessment Findings & Results Matrix
    # ──────────────────────────────────────────────────────────────────────────
    s10 = setup_slide(is_dark=False)
    add_header(s10, "Expected Outcome 5: Vendor Assessment Findings", "Definitive Findings Matrix: Exactly 1 Inconclusive Reflects Methodological Rigor")

    t_shape10 = s10.shapes.add_table(7, 6, Inches(0.8), Inches(1.68), Inches(11.733), Inches(5.15))
    tbl10 = t_shape10.table
    tbl10.columns[0].width = Inches(0.9)
    tbl10.columns[1].width = Inches(2.1)
    tbl10.columns[2].width = Inches(1.1)
    tbl10.columns[3].width = Inches(1.1)
    tbl10.columns[4].width = Inches(1.2)
    tbl10.columns[5].width = Inches(5.333)

    h10 = ["VID", "Vendor Name", "Criticality", "AI Usage", "Risk Rating", "Identified Providers & Verified Footprint"]
    for i, h in enumerate(h10):
        cell = tbl10.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY_HEADER
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = WHITE

    findings_rows = [
        ("V-001", "AutomWorx", "Low", "Yes", "Medium", "Google (Gemini), OpenAI (ChatGPT). AI policy confirms consultants use commercial LLMs for scripting/work completion."),
        ("V-002", "Fiserv, Inc.", "Critical", "Yes", "Critical", "Azure OpenAI, Apple (Siri), Amazon (Alexa), Fiserv. 124 signals across AML Risk Manager, predictive analytics, agentic AI."),
        ("V-003", "FSSI", "High", "No", "None ident.", "N/A. Thorough OSINT sweep of 14 Tier A/B pages (SOC 2, trust center, print operations). Zero AI claims detected."),
        ("V-004", "Terrapin Tech.", "Medium", "Inconclusive", "Medium", "None named. Minimal web footprint (1 page reachable; API/docs 404). Honest inconclusive requiring targeted questionnaire."),
        ("V-005", "BNY", "High", "Yes", "High", "Microsoft, OpenAI, BNY AI Hub. Confirmed AI-powered document processing and workflow optimization in corporate platforms."),
        ("V-006", "The Clearing House", "Critical", "No", "None ident.", "N/A. Extensive audit of 1,016 pages of network rules & specs. Core ACH/RTP clearing rail operates deterministically without AI.")
    ]

    for row_idx, row in enumerate(findings_rows, 1):
        for col_idx, val in enumerate(row):
            cell = tbl10.cell(row_idx, col_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if row_idx % 2 == 1 else RGBColor(241, 245, 249)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(9)
            if col_idx == 3:
                p.font.bold = True
                p.font.color.rgb = RED_ACCENT if val == "Yes" else GREEN_ACCENT if val == "No" else AMBER_ACCENT
            elif col_idx == 4:
                p.font.bold = True
                p.font.color.rgb = RED_ACCENT if val == "Critical" else AMBER_ACCENT if val == "High" else BLUE_ACCENT if val == "Medium" else GREEN_ACCENT
            else:
                p.font.color.rgb = TEXT_DARK

    s10.notes_slide.notes_text_frame.text = (
        "Expected Outcome 5: This is our core finding matrix. We eliminated excess ambiguity: exactly ONE vendor "
        "is Inconclusive (Terrapin), which is driven by genuine bot-blocking. Fiserv, BNY, and AutomWorx are confirmed 'Yes'; "
        "FSSI and The Clearing House are confirmed 'No'. This demonstrates that our tool delivers clear answers "
        "while refusing to guess when public footprints are unfindable."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 11: Surfacing Enterprise Risk & Governance Action Playbook
    # ──────────────────────────────────────────────────────────────────────────
    s11 = setup_slide(is_dark=False)
    add_header(s11, "Surfacing Enterprise Risk & Governance Playbook", "Mapping Public AI Footprints to Meridian's Risk Framework & 15/30-Day Action Plans")

    add_card(s11, 0.8, 1.68, 5.7, 5.15, "FOUR CORE ENTERPRISE AI RISKS", NAVY_HEADER)
    r1_box11 = s11.shapes.add_textbox(Inches(1.05), Inches(2.20), Inches(5.2), Inches(4.5))
    r1_tf11 = r1_box11.text_frame
    r1_tf11.word_wrap = True
    r1_tf11.margin_left = r1_tf11.margin_top = r1_tf11.margin_right = r1_tf11.margin_bottom = 0

    r_risks11 = [
        ("1. Sensitive Data Ingestion:", "Commercial LLMs or unvetted pipelines ingesting customer records or payment transactions.", "Scope: 1.4M customer accounts (Fiserv DNA), 2.1M payment messages (BNY RTP)."),
        ("2. Shadow AI Sub-processors:", "Third-party cloud models (OpenAI, Microsoft Azure) deployed without formal DPA listing.", "Scope: Unmonitored chain-of-custody across outsourced banking and workflow services."),
        ("3. Script & Credential Leakage:", "Consulting staff pasting internal scripts or production credentials into public LLMs.", "Scope: Direct infrastructure exposure via unvetted scripting automation (AutomWorx)."),
        ("4. Contractual & Audit Voids:", "Absence of contractual guarantees prohibiting vendor models from training on client data.", "Scope: Legal inability to audit model architectures or demand data purging on termination.")
    ]
    for h, d, sc in r_risks11:
        bp = r1_tf11.add_paragraph() if r1_tf11.paragraphs[0].text else r1_tf11.paragraphs[0]
        bp.text = f"▪ {h}"
        bp.font.size = Pt(10)
        bp.font.bold = True
        bp.font.color.rgb = NAVY_HEADER
        bp.space_before = Pt(6)

        bp_d = r1_tf11.add_paragraph()
        bp_d.text = f"{d}\n↳ {sc}"
        bp_d.font.size = Pt(8.8)
        bp_d.font.color.rgb = TEXT_DARK
        bp_d.space_before = Pt(2)

    add_card(s11, 6.833, 1.68, 5.7, 5.15, "RECOMMENDED GOVERNANCE PLAYBOOK", ORANGE)
    r2_box11 = s11.shapes.add_textbox(Inches(7.08), Inches(2.20), Inches(5.2), Inches(4.5))
    r2_tf11 = r2_box11.text_frame
    r2_tf11.word_wrap = True
    r2_tf11.margin_left = r2_tf11.margin_top = r2_tf11.margin_right = r2_tf11.margin_bottom = 0

    r_actions11 = [
        ("Immediate 15-Day Targeted Questionnaires:", "Issue specific inquiries to Fiserv, BNY, and AutomWorx covering model providers, hosting region, data retention policies, and training restrictions.", "Timeline: Dispatched within 5 business days; response due in 15 days."),
        ("AI Sub-Processor Inventory Registration:", "Formally register named providers (Microsoft, OpenAI, Google) with an unresolved-provider flag pending written vendor verification.", "Timeline: Immediate integration into Meridian's enterprise GRC repository."),
        ("Contractual Remediation Clauses:", "Add mandatory clauses at contract review: strict training prohibitions, mandatory 30-day sub-processor change notices, and AI audit rights.", "Timeline: Required for all upcoming renewals and new supplier contracts."),
        ("30-Day TPRC Escalation Path:", "If vendor confirmations are not received within 30 days, escalate to Meridian Third-Party Risk Committee for operational containment.", "Timeline: Formal review and potential transaction throttling on Day 30.")
    ]
    for h, d, tl in r_actions11:
        bp = r2_tf11.add_paragraph() if r2_tf11.paragraphs[0].text else r2_tf11.paragraphs[0]
        bp.text = f"▪ {h}"
        bp.font.size = Pt(10)
        bp.font.bold = True
        bp.font.color.rgb = ORANGE
        bp.space_before = Pt(6)

        bp_d = r2_tf11.add_paragraph()
        bp_d.text = f"{d}\n↳ {tl}"
        bp_d.font.size = Pt(8.8)
        bp_d.font.color.rgb = TEXT_DARK
        bp_d.space_before = Pt(2)

    s11.notes_slide.notes_text_frame.text = (
        "Here is the risk surfacing step. We translate technical AI signals into four enterprise risks: data exposure, "
        "shadow sub-processors, credential leakage, and contractual voids. We back this with a concrete 15-day "
        "questionnaire playbook and 30-day escalation to the Risk Committee."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 12: Working Demonstration — Interactive Streamlit App & Excel Deliverables
    # ──────────────────────────────────────────────────────────────────────────
    s12 = setup_slide(is_dark=False)
    add_header(s12, "Working Demonstration & Deliverables Summary", "Live Web Demonstration (Instant Presentation Mode) and Completed Audit Deliverables")

    add_card(s12, 0.8, 1.68, 5.7, 5.15, "STREAMLIT DEMO: INSTANT PRESENTATION MODE", ORANGE)
    ui_box12 = s12.shapes.add_textbox(Inches(1.05), Inches(2.22), Inches(5.2), Inches(4.4))
    ui_tf12 = ui_box12.text_frame
    ui_tf12.word_wrap = True
    ui_tf12.margin_left = ui_tf12.margin_top = ui_tf12.margin_right = ui_tf12.margin_bottom = 0

    ui_points12 = [
        ("Zero Live Latency:", "Automatically loads pre-computed, cryptographically verified results upon page load (http://localhost:8501). No waiting for web crawls during meetings."),
        ("Executive KPI Banner:", "Surfaces top metrics at a glance: 6 Assessed, 3 Confirmed AI, 2 Ruled Out, 1 Inconclusive, 1 Critical Risk."),
        ("Per-Vendor Evidence Cards:", "Expandable deep-dive cards showing exact verbatim quotes, source URLs, retrieval dates, and named model providers."),
        ("Live Execution Drawer:", "An optional expander allows live crawling and real-time Groq LPU inference on custom uploaded workbooks if requested."),
        ("One-Click Download:", "Direct download button for the completed Meridian_Vendor_Output.xlsx.")
    ]
    for h, d in ui_points12:
        bp = ui_tf12.add_paragraph() if ui_tf12.paragraphs[0].text else ui_tf12.paragraphs[0]
        bp.text = f"▪ {h} {d}"
        bp.font.size = Pt(9.5)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(6)

    add_card(s12, 6.833, 1.68, 5.7, 5.15, "COMPLETED EXCEL WORKBOOK DELIVERABLE", NAVY_HEADER)
    xl_box12 = s12.shapes.add_textbox(Inches(7.08), Inches(2.22), Inches(5.2), Inches(4.4))
    xl_tf12 = xl_box12.text_frame
    xl_tf12.word_wrap = True
    xl_tf12.margin_left = xl_tf12.margin_top = xl_tf12.margin_right = xl_tf12.margin_bottom = 0

    xl_points12 = [
        ("Columns B–K Integrity:", "100% byte-for-byte identical to client input. Verified by automated round-trip tests."),
        ("Example Row (V-000):", "Retained verbatim as requested in specification. Formats and styling preserved."),
        ("Columns L–V Populated:", "Criticality Tier, Rationale, Depth, AI Usage, Evidence, How Using, Providers, Risk, Risk Rationale, Recommended Action, Assessed By."),
        ("132-Entry Evidence Log Sheet:", "New audited sheet detailing Vendor ID, Source Tier, URL, Retrieval Date, Verbatim Excerpt, Signal Tag, and SHA-256 Snapshot Hash."),
        ("Automated Test Suite:", "18 unit tests (pytest) verifying criticality determinism, quote guard, weighting logic, and column B–K round-trip integrity.")
    ]
    for h, d in xl_points12:
        bp = xl_tf12.add_paragraph() if xl_tf12.paragraphs[0].text else xl_tf12.paragraphs[0]
        bp.text = f"▪ {h} {d}"
        bp.font.size = Pt(9.5)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(6)

    s12.notes_slide.notes_text_frame.text = (
        "Working demonstration: Our Streamlit app is live on localhost:8501. It loads pre-computed results instantly, "
        "making it ideal for executive presentations. The final Excel workbook has all columns populated, with "
        "columns B-K untouched, and includes a 132-row Evidence Log sheet."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 13: Strategic Governance Roadmap & Next Steps
    # ──────────────────────────────────────────────────────────────────────────
    s13 = setup_slide(is_dark=False)
    add_header(s13, "Strategic Governance Roadmap & Next Steps", "A Phased 90-Day Implementation Plan for Meridian Third-Party Risk Management")

    phases = [
        ("PHASE 1: DAYS 1–30", "IMMEDIATE TRIAGE & ENQUIRIES", [
            ("Core Focus:", "Isolate high-risk exposures and dispatch formal inquiries to confirmed AI vendors."),
            ("Workstream 1:", "Dispatch 15-day targeted questionnaires to Fiserv (DNA core banking), BNY (RTP rail), and AutomWorx."),
            ("Workstream 2:", "Register named AI model providers (Microsoft, OpenAI, Google) in Meridian's GRC inventory."),
            ("Key Deliverable:", "Targeted TPRM Questionnaires & Interim Shadow AI Sub-processor Register."),
            ("Success Metric:", "100% questionnaire delivery within 10 days; 0 unvetted external models processing live data.")
        ], NAVY_HEADER),
        ("PHASE 2: DAYS 31–60", "ESCALATION & CLAUSE REMEDIATION", [
            ("Core Focus:", "Enforce contractual safeguards and escalate unresponsive or unverified vendor claims."),
            ("Workstream 1:", "Escalate unconfirmed vendor responses to Meridian Third-Party Risk Committee (TPRC)."),
            ("Workstream 2:", "Demand formal architecture data-flow diagrams from Fiserv and BNY confirming training boundaries."),
            ("Key Deliverable:", "Executed AI Model Contract Addenda & Mandatory 30-Day Sub-processor Change Notices."),
            ("Success Metric:", "100% of Critical vendors contractually bound by training prohibitions and audit rights.")
        ], ORANGE),
        ("PHASE 3: DAYS 61–90", "CONTINUOUS OSINT AUTOMATION", [
            ("Core Focus:", "Transition from point-in-time assessment to continuous, automated third-party AI discovery."),
            ("Workstream 1:", "Integrate OSINT collectors into quarterly CI/CD scheduled crawls with automated diff alerting."),
            ("Workstream 2:", "Conduct scheduled manual deep-dives on bot-protected vendors (Terrapin Technologies)."),
            ("Key Deliverable:", "Automated Quarterly Crawl Engine & Executive Board TPRM Risk Posture Dashboard."),
            ("Success Metric:", "Continuous, auditable third-party visibility across 100% of Meridian's vendor ecosystem.")
        ], NAVY_HEADER)
    ]

    for idx, (p_tag, p_title, p_items, p_col) in enumerate(phases):
        left = 0.8 + idx * 4.0
        add_card(s13, left, 1.68, 3.75, 5.15, p_tag, p_col)

        tb = s13.shapes.add_textbox(Inches(left + 0.22), Inches(2.20), Inches(3.31), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p = tf.paragraphs[0]
        p.text = p_title
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = p_col

        for h, d in p_items:
            bp = tf.add_paragraph()
            bp.text = f"▪ {h}"
            bp.font.size = Pt(9.5)
            bp.font.bold = True
            bp.font.color.rgb = p_col
            bp.space_before = Pt(5)

            bp_d = tf.add_paragraph()
            bp_d.text = d
            bp_d.font.size = Pt(8.5)
            bp_d.font.color.rgb = TEXT_DARK
            bp_d.space_before = Pt(1)

    s13.notes_slide.notes_text_frame.text = (
        "Roadmap slide: We transition Meridian from a one-off audit to continuous governance. Over 90 days, "
        "we triage immediate questionnaire findings, remediate contract clauses, and establish quarterly automated crawls."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 14: Discussion, Q&A & Thank You (Optiv Navy)
    # ──────────────────────────────────────────────────────────────────────────
    s14 = setup_slide(is_dark=True, category_text="DISCUSSION & NEXT STEPS  |  OPTIV CONSULTING")

    # Title box
    tb_qtitle = s14.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.733), Inches(0.6))
    tf_qt = tb_qtitle.text_frame
    tf_qt.margin_left = tf_qt.margin_top = tf_qt.margin_right = tf_qt.margin_bottom = 0
    qp1 = tf_qt.paragraphs[0]
    qp1.text = "Discussion & Questions"
    qp1.font.size = Pt(36)
    qp1.font.bold = True
    qp1.font.color.rgb = ORANGE

    # Subtitle
    tb_qsub = s14.shapes.add_textbox(Inches(0.8), Inches(1.85), Inches(11.733), Inches(0.35))
    tf_qs = tb_qsub.text_frame
    tf_qs.margin_left = tf_qs.margin_top = tf_qs.margin_right = tf_qs.margin_bottom = 0
    qp2 = tf_qs.paragraphs[0]
    qp2.text = "Thank you for your time and engagement. Key discussion prompts for the Third-Party Risk Committee:"
    qp2.font.size = Pt(13)
    qp2.font.color.rgb = RGBColor(226, 232, 240)

    # Discussion prompts card
    dp_card = s14.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.35), Inches(11.733), Inches(2.2))
    dp_card.fill.solid()
    dp_card.fill.fore_color.rgb = RGBColor(17, 34, 64)
    dp_card.line.color.rgb = RGBColor(30, 58, 102)
    dp_card.line.width = Pt(1.2)

    dp_box = s14.shapes.add_textbox(Inches(1.1), Inches(2.5), Inches(11.1), Inches(1.9))
    dp_tf = dp_box.text_frame
    dp_tf.word_wrap = True
    dp_tf.margin_left = dp_tf.margin_top = dp_tf.margin_right = dp_tf.margin_bottom = 0

    prompts = [
        ("1. Questionnaire Prioritization:", "Validating the 15-day targeted questionnaire scope for Fiserv, BNY, and AutomWorx."),
        ("2. Inconclusive Handling:", "Endorsing our methodology of isolating Terrapin Technologies as the sole honest coverage gap."),
        ("3. Enterprise GRC Integration:", "Operationalizing quarterly automated OSINT crawls within Meridian's existing risk tools.")
    ]
    for idx, (ph, pd) in enumerate(prompts):
        p = dp_tf.add_paragraph() if idx > 0 else dp_tf.paragraphs[0]
        p.text = f"▪ {ph} {pd}"
        p.font.size = Pt(11)
        p.font.color.rgb = WHITE
        if idx > 0:
            p.space_before = Pt(8)

    # Closing Executive Sign-off Card
    q_card = s14.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(4.75), Inches(11.733), Inches(2.1))
    q_card.fill.solid()
    q_card.fill.fore_color.rgb = RGBColor(17, 34, 64)
    q_card.line.color.rgb = RGBColor(30, 58, 102)
    q_card.line.width = Pt(1.5)

    q_strip = s14.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(4.75), Inches(11.733), Inches(0.06))
    q_strip.fill.solid()
    q_strip.fill.fore_color.rgb = ORANGE
    q_strip.line.fill.background()

    qc_box = s14.shapes.add_textbox(Inches(1.1), Inches(4.98), Inches(11.1), Inches(1.7))
    qc_tf = qc_box.text_frame
    qc_tf.word_wrap = True
    qc_tf.margin_left = qc_tf.margin_top = qc_tf.margin_right = qc_tf.margin_bottom = 0

    qc_p = qc_tf.paragraphs[0]
    qc_p.text = "AIMAN DHANANI  •  REGISTRATION NUMBER: 24BYB0033"
    qc_p.font.size = Pt(15.5)
    qc_p.font.bold = True
    qc_p.font.color.rgb = ORANGE

    qc_p2 = qc_tf.add_paragraph()
    qc_p2.text = "Cyber Strategy Practice  |  Optiv Consulting  |  VIT Campus Engagement\n" \
                 "Deliverables: data/Meridian_Vendor_Output.xlsx  |  Streamlit Presentation Demo  |  Unit Tests (18/18 Pass)\n" \
                 "Repository: https://github.com/Aiman-D/optiv-case-study-1  |  Confidential  |  September 2026"
    qc_p2.font.size = Pt(11.5)
    qc_p2.font.color.rgb = RGBColor(203, 213, 225)
    qc_p2.space_before = Pt(6)

    s14.notes_slide.notes_text_frame.text = (
        "Thank you everyone. I am now open for any questions or deep-dives into our scoring rubric, "
        "the Quote Guard implementation, or the live Streamlit demo."
    )

    output_path = "/Users/aimandhanani/Desktop/Useful/Optiv_proj/cs1/Case_Study_1_Aiman_Dhanani_24BYB0033.pptx"
    prs.save(output_path)
    print(f"✅ Presentation successfully generated and saved to: {output_path}")

if __name__ == '__main__':
    create_deck()
