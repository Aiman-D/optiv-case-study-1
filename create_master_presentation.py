"""
Master Presentation Generator for Optiv Consulting Case Studies.
Generates an executive-tier, in-depth 22-slide PowerPoint presentation covering:
- Case Study 1: Vendor AI Usage Analysis (Meridian Financial)
- Case Study 2: Text Analysis & PII Detection (Cadence Financial Group)
Adheres strictly to Optiv brand guidelines (Widescreen 16:9, Navy/Orange palette,
structured cards, visual tables, architecture flows, and embedded presenter notes).
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

# ── Color Palette ─────────────────────────────────────────────────────────────
NAVY_DARK = RGBColor(11, 25, 44)       # #0B192C - Deep background / primary
NAVY_HEADER = RGBColor(20, 43, 76)     # #142B4C - Optiv classic navy
ORANGE = RGBColor(237, 107, 45)        # #ED6B2D - Optiv brand accent orange
ORANGE_LIGHT = RGBColor(255, 237, 226) # #FFEDE2 - Soft orange badge bg
WHITE = RGBColor(255, 255, 255)        # #FFFFFF - Crisp white
GRAY_BG = RGBColor(248, 250, 252)      # #F8FAFC - Canvas light gray
CARD_BG = RGBColor(255, 255, 255)      # #FFFFFF - Card container bg
CARD_BORDER = RGBColor(226, 232, 240)  # #E2E8F0 - Subtle card border
TEXT_DARK = RGBColor(15, 23, 42)       # #0F172A - Deep slate for readability
TEXT_MUTED = RGBColor(100, 116, 139)   # #64748B - Secondary subtitle text
TEXT_LIGHT = RGBColor(226, 232, 240)   # #E2E8F0 - Light text on dark bg

# Status colors
GREEN_ACCENT = RGBColor(16, 185, 129)  # #10B981 - Low risk / No AI
RED_ACCENT = RGBColor(239, 68, 68)     # #EF4444 - Critical risk / Yes
AMBER_ACCENT = RGBColor(245, 158, 11)  # #F59E0B - High / Medium risk
BLUE_ACCENT = RGBColor(59, 130, 246)   # #3B82F6 - Informational

def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Helper: Add base background and branding
    def setup_slide(is_dark=False, category_text="OPTIV CONSULTING  |  CYBER STRATEGY"):
        slide = prs.slides.add_slide(blank_layout)
        
        # Background fill
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = NAVY_DARK if is_dark else GRAY_BG
        bg.line.fill.background()

        # Top corner bracket accent
        bracket = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.4), Inches(0.18), Inches(0.18))
        bracket.fill.solid()
        bracket.fill.fore_color.rgb = ORANGE
        bracket.line.fill.background()

        # Header category label
        cat_box = slide.shapes.add_textbox(Inches(1.1), Inches(0.35), Inches(8.0), Inches(0.3))
        tf = cat_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = category_text
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = ORANGE if is_dark else NAVY_HEADER

        # Right header branding
        brand_box = slide.shapes.add_textbox(Inches(9.5), Inches(0.35), Inches(3.0), Inches(0.3))
        tf_b = brand_box.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = "Optiv Consulting"
        p_b.alignment = PP_ALIGN.RIGHT
        p_b.font.size = Pt(12)
        p_b.font.bold = True
        p_b.font.color.rgb = WHITE if is_dark else NAVY_HEADER

        # Footer
        ft_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.7), Inches(0.3))
        tf_ft = ft_box.text_frame
        p_ft = tf_ft.paragraphs[0]
        p_ft.text = "Confidential  |  Prepared for Stakeholder Review  |  September 2026"
        p_ft.font.size = Pt(9)
        p_ft.font.color.rgb = RGBColor(148, 163, 184) if is_dark else RGBColor(148, 163, 184)

        return slide

    def add_header(slide, title, subtitle="", is_dark=False):
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(11.7), Inches(1.1))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = WHITE if is_dark else NAVY_HEADER

        # Accent bar under title
        accent_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(1.2), Inches(0.06))
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = ORANGE
        accent_bar.line.fill.background()

        if subtitle:
            p_sub = tf.add_paragraph()
            p_sub.text = subtitle
            p_sub.font.size = Pt(13)
            p_sub.font.color.rgb = RGBColor(203, 213, 225) if is_dark else TEXT_MUTED
            p_sub.space_before = Pt(8)

    def add_card(slide, left, top, width, height, title="", title_color=NAVY_HEADER, bg_color=CARD_BG, border_color=CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1)

        # Top accent strip
        if title:
            strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left + 0.15), Inches(top), Inches(width - 0.3), Inches(0.05))
            strip.fill.solid()
            strip.fill.fore_color.rgb = title_color if title_color != NAVY_HEADER else ORANGE
            strip.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(left + 0.25), Inches(top + 0.15), Inches(width - 0.5), Inches(0.4))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(15)
            p.font.bold = True
            p.font.color.rgb = title_color

        return card

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 1: Title Slide (Dark Navy)
    # ──────────────────────────────────────────────────────────────────────────
    s1 = setup_slide(is_dark=True, category_text="CONFIDENTIAL  |  OPTIV CONSULTING")
    
    # Large Title
    tbox = s1.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(11.5), Inches(2.5))
    tf1 = tbox.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "Enterprise AI Governance & Cyber Strategy"
    p1.font.size = Pt(40)
    p1.font.bold = True
    p1.font.color.rgb = WHITE

    # Accent bar
    sbar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(3.3), Inches(2.5), Inches(0.08))
    sbar.fill.solid()
    sbar.fill.fore_color.rgb = ORANGE
    sbar.line.fill.background()

    p1_sub = tf1.add_paragraph()
    p1_sub.text = "Comprehensive Solutions for Case Study 1 (Meridian Financial) & Case Study 2 (Cadence Financial Group)"
    p1_sub.font.size = Pt(20)
    p1_sub.font.color.rgb = RGBColor(226, 232, 240)
    p1_sub.space_before = Pt(20)

    # Metadata card
    meta = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(4.7), Inches(11.7), Inches(1.6))
    meta.fill.solid()
    meta.fill.fore_color.rgb = RGBColor(17, 34, 64)
    meta.line.color.rgb = RGBColor(30, 58, 102)

    mt_box = s1.shapes.add_textbox(Inches(1.1), Inches(4.85), Inches(11.1), Inches(1.3))
    mt_tf = mt_box.text_frame
    mt_tf.word_wrap = True
    p_m1 = mt_tf.paragraphs[0]
    p_m1.text = "ENGAGEMENT OVERVIEW & DELIVERABLES"
    p_m1.font.size = Pt(13)
    p_m1.font.bold = True
    p_m1.font.color.rgb = ORANGE

    p_m2 = mt_tf.add_paragraph()
    p_m2.text = "• Case Study 1: Open-Source Intelligence (OSINT) Vendor AI Footprint Profiling & Security Risk Engine\n" \
               "• Case Study 2: Local Privacy-Preserving Text Extraction, PII Detection & Pixel-Level Redaction Pipeline\n" \
               "• Walkthrough Deck: 10–15 Minute Executive Walkthrough  |  Optiv Strategy Practice  |  VIT Engagement"
    p_m2.font.size = Pt(13)
    p_m2.font.color.rgb = RGBColor(203, 213, 225)
    p_m2.space_before = Pt(8)

    s1.notes_slide.notes_text_frame.text = (
        "Welcome everyone. Today we are presenting our end-to-end consulting deliverables for both Case Study 1 "
        "and Case Study 2. We address the full lifecycle of enterprise AI exposure: analyzing undisclosed third-party "
        "AI risks from outside vendors for Meridian Financial, and engineering zero-leakage PII sanitization pipelines "
        "for Cadence Financial Group."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 2: Executive Agenda & Strategic Context
    # ──────────────────────────────────────────────────────────────────────────
    s2 = setup_slide(is_dark=False)
    add_header(s2, "Executive Agenda & Engagement Framework", "Addressing Dual Enterprise AI Risk Dimensions: Inbound Vendor Risk vs Outbound PII Ingestion")

    # Column 1: Case Study 1 Card
    add_card(s2, 0.8, 1.8, 5.7, 4.9, "CASE STUDY 1: THIRD-PARTY AI RISK (MERIDIAN)", ORANGE)
    c1_box = s2.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(5.2), Inches(4.1))
    c1_tf = c1_box.text_frame
    c1_tf.word_wrap = True
    c1_p = c1_tf.paragraphs[0]
    c1_p.text = "Client: Meridian Financial (Global Banking Enterprise)\nChallenge: Undisclosed Vendor AI in Core Workloads"
    c1_p.font.size = Pt(13)
    c1_p.font.bold = True
    c1_p.font.color.rgb = NAVY_HEADER

    c1_bullets = [
        "Problem: Vendors embed generative & predictive AI into services without formal disclosure.",
        "Methodology: 100% OSINT footprint analysis (zero contact, no paywall bypass).",
        "Deterministic Criticality Engine: 5-metric rubric assigning Critical to Low tiers.",
        "Quote Guard: Cryptographic substring verification to eliminate LLM hallucinations.",
        "Outcome: Preserved client workbook, 132-entry Evidence Log, and instant Streamlit demo."
    ]
    for b in c1_bullets:
        bp = c1_tf.add_paragraph()
        bp.text = f"▪  {b}"
        bp.font.size = Pt(12)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(8)

    # Column 2: Case Study 2 Card
    add_card(s2, 6.8, 1.8, 5.7, 4.9, "CASE STUDY 2: DOCUMENT PII REDACTION (CADENCE)", NAVY_HEADER)
    c2_box = s2.shapes.add_textbox(Inches(7.05), Inches(2.35), Inches(5.2), Inches(4.1))
    c2_tf = c2_box.text_frame
    c2_tf.word_wrap = True
    c2_p = c2_tf.paragraphs[0]
    c2_p.text = "Client: Cadence Financial Group (Investment Banking)\nChallenge: Preventing Sensitive PII from Reaching LLMs"
    c2_p.font.size = Pt(13)
    c2_p.font.bold = True
    c2_p.font.color.rgb = NAVY_HEADER

    c2_bullets = [
        "Policy: Strict organizational mandate requiring all PII stripped prior to LLM analysis.",
        "Multi-Format Extraction: Native PDF, DOCX, PPTX extraction retaining >80% structure.",
        "OCR & Figure Pass: Tesseract fallback for scanned policy docs and embedded charts.",
        "Hybrid PII Detector: Deterministic regex + spaCy NER (en_core_web_lg) targeting 11 entities.",
        "Reversible Token Vault: Cryptographic token mapping (.vault.json) for auditability."
    ]
    for b in c2_bullets:
        bp = c2_tf.add_paragraph()
        bp.text = f"▪  {b}"
        bp.font.size = Pt(12)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(8)

    s2.notes_slide.notes_text_frame.text = (
        "This agenda slide illustrates the dual-shield approach: Case Study 1 governs incoming risk from third-party "
        "partners who might process bank data through unvetted models. Case Study 2 governs outgoing data integrity by "
        "sanitizing enterprise artifacts before feeding them to internal AI agents."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 3: Section Divider — Case Study 1 (Dark Navy)
    # ──────────────────────────────────────────────────────────────────────────
    s3 = setup_slide(is_dark=True, category_text="CASE STUDY 1  |  MERIDIAN FINANCIAL")
    
    div_box = s3.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(11.7), Inches(3.0))
    div_tf = div_box.text_frame
    div_tf.word_wrap = True
    
    dp_tag = div_tf.paragraphs[0]
    dp_tag.text = "SECTION 01"
    dp_tag.font.size = Pt(18)
    dp_tag.font.bold = True
    dp_tag.font.color.rgb = ORANGE

    dp_title = div_tf.add_paragraph()
    dp_title.text = "Case Study 1: Vendor AI Usage Analysis"
    dp_title.font.size = Pt(38)
    dp_title.font.bold = True
    dp_title.font.color.rgb = WHITE
    dp_title.space_before = Pt(10)

    dp_desc = div_tf.add_paragraph()
    dp_desc.text = "Designing an Evidence-Driven OSINT Methodology to Profile and Classify Third-Party AI Footprints"
    dp_desc.font.size = Pt(18)
    dp_desc.font.color.rgb = RGBColor(203, 213, 225)
    dp_desc.space_before = Pt(15)

    s3.notes_slide.notes_text_frame.text = (
        "Now entering Case Study 1. We were engaged by Meridian Financial to create a repeatable, defensible "
        "methodology to inspect their critical third-party vendors."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 4: Case Study 1 — Problem Statement & Hard Constraints
    # ──────────────────────────────────────────────────────────────────────────
    s4 = setup_slide(is_dark=False)
    add_header(s4, "Case Study 1: Context, Problem & Constraints", "Meridian Financial Third-Party Risk Management Engagement")

    # Card 1: Background & Problem
    add_card(s4, 0.8, 1.8, 5.7, 4.9, "THE BUSINESS CHALLENGE", NAVY_HEADER)
    p_box = s4.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(5.2), Inches(4.1))
    p_tf = p_box.text_frame
    p_tf.word_wrap = True
    
    p_items = [
        ("Client Profile:", "Meridian Financial is a global financial institution serving 1.4 million customers, processing billions in payment transactions annually."),
        ("The Challenge:", "Vendors in cloud, managed services, and payments increasingly embed AI into core systems. Disclosures are vague, marketing buzz is rampant, and contract terms lag behind."),
        ("Risk Posture:", "If a vendor feeds Meridian customer records or payment flows into commercial LLMs or undisclosed AI sub-processors, Meridian faces catastrophic regulatory and privacy liabilities."),
        ("The Objective:", "Optiv was tasked to design an automated, evidence-driven tool to verify whether vendors genuinely use AI in the services provided to Meridian.")
    ]
    for h, d in p_items:
        bp = p_tf.add_paragraph() if p_tf.paragraphs[0].text else p_tf.paragraphs[0]
        bp.text = f"▪  {h} {d}"
        bp.font.size = Pt(12)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(10)

    # Card 2: Hard Constraints & Guarantees
    add_card(s4, 6.8, 1.8, 5.7, 4.9, "HARD OPERATIONAL CONSTRAINTS", ORANGE)
    c_box = s4.shapes.add_textbox(Inches(7.05), Inches(2.35), Inches(5.2), Inches(4.1))
    c_tf = c_box.text_frame
    c_tf.word_wrap = True

    c_items = [
        ("100% OSINT Only:", "No vendor contact, no questionnaires during collection, no paywall bypass, and strictly respect robots.txt and a 1.2-second rate limit."),
        ("Cryptographic Auditability:", "Every evidence claim must trace to a local snapshot with URL, retrieval timestamp, and SHA-256 content hash."),
        ("Quote Guard Enforcement:", "LLM-extracted quotes must be exact verbatim substrings of cleaned HTML. Any hallucinated quote is discarded immediately."),
        ("Coverage Gap Accounting:", "Unreachable pages or bot-blocked sites are recorded explicitly as Coverage Gaps—never reported as a false 'No'."),
        ("Workbook Preservation:", "Columns B–K in Meridian's workbook are byte-for-byte immutable. Only columns L–V are populated.")
    ]
    for h, d in c_items:
        bp = c_tf.add_paragraph() if c_tf.paragraphs[0].text else c_tf.paragraphs[0]
        bp.text = f"▪  {h} {d}"
        bp.font.size = Pt(12)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(10)

    s4.notes_slide.notes_text_frame.text = (
        "Key point for the stakeholder: We operated under strict ethical and legal boundaries. We did not contact "
        "vendors directly or attempt unauthorized scraping. Everything is built on public OSINT, backed by exact "
        "cryptographic hashes so our findings withstand an audit."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 5: Case Study 1 — Roles & Responsibilities Framework
    # ──────────────────────────────────────────────────────────────────────────
    s5 = setup_slide(is_dark=False)
    add_header(s5, "Roles & Responsibilities: 5-Pillar Architecture", "The Analytical Journey from Raw Vendor Profiles to Defensible Risk Outputs")

    pillars = [
        ("01", "DETERMINE CRITICALITY", "Judge how critical each vendor is using anchored 0–3 scoring scales across 5 business dimensions so another analyst gets the exact same score.", NAVY_HEADER),
        ("02", "SET THE DEPTH", "Establish tier-driven depth policies: Critical requires a full web/filings sweep, whereas Low warrants only a trust center/policy check.", ORANGE),
        ("03", "IDENTIFY SOURCES", "Categorize OSINT footprint into 4 weighted reliability tiers (Tier A Direct down to Tier D Marketing Noise) and enforce service linkage.", NAVY_HEADER),
        ("04", "BUILD CAPABILITY", "Scrape, clean, and classify using a hybrid regex + LLM pipeline with substring Quote Guard and local caching for offline capability.", ORANGE),
        ("05", "SURFACE THE RISK", "Map confirmed AI signals to Meridian's enterprise risk matrix (Critical/High/Medium/None) and draft 15-day actionable escalation playbooks.", NAVY_HEADER),
    ]

    for idx, (num, p_title, p_desc, col) in enumerate(pillars):
        left = 0.8 + idx * 2.4
        card = add_card(s5, left, 1.8, 2.25, 4.9, f"STEP {num}", col)
        
        tb = s5.shapes.add_textbox(Inches(left + 0.15), Inches(2.4), Inches(1.95), Inches(4.1))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = p_title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = col

        p2 = tf.add_paragraph()
        p2.text = p_desc
        p2.font.size = Pt(11)
        p2.font.color.rgb = TEXT_DARK
        p2.space_before = Pt(12)

    s5.notes_slide.notes_text_frame.text = (
        "Here are the 5 core steps answering the briefing slides. Step 1 establishes vendor criticality; Step 2 sets "
        "the depth; Step 3 weights OSINT sources; Step 4 executes the data extraction with Quote Guard; Step 5 surfaces "
        "the resulting risk to Meridian."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 6: Expected Outcome 1 — Functional Design Diagram
    # ──────────────────────────────────────────────────────────────────────────
    s6 = setup_slide(is_dark=False)
    add_header(s6, "Expected Outcome 1: Functional Design Diagram", "End-to-End Pipeline Data Flow from OSINT Collection to Final Risk Workbook")

    # Step Boxes
    steps = [
        ("M1: Criticality Engine", "• Ingests cols B–K\n• Rubric scoring (0–3)\n• Modifier & floor rules\n• Tier + Depth output", NAVY_HEADER),
        ("M2: OSINT Collector", "• Robots.txt parsing\n• Trafilatura extraction\n• 1.2s rate limiter\n• Immutable snapshot cache", ORANGE),
        ("M3: AI Classifier", "• 25+ regex pre-filters\n• Groq LPU inference\n• JSON schema output\n• Substring Quote Guard", NAVY_HEADER),
        ("M4: Risk Engine", "• Multi-factor score\n• Recency decay (0.7x)\n• Coverage thresholding\n• Verdict & risk mapping", ORANGE),
        ("M5: Excel & UI", "• Copy-first workbook\n• Fill cols L–V only\n• Evidence Log sheet\n• Streamlit presentation", NAVY_HEADER)
    ]

    for i, (stitle, sdesc, scol) in enumerate(steps):
        left = 0.8 + i * 2.4
        box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(2.0), Inches(2.25), Inches(3.2))
        box.fill.solid()
        box.fill.fore_color.rgb = CARD_BG
        box.line.color.rgb = scol
        box.line.width = Pt(2)

        # Header bar
        hbar = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(2.0), Inches(2.25), Inches(0.5))
        hbar.fill.solid()
        hbar.fill.fore_color.rgb = scol
        hbar.line.fill.background()

        tb_h = s6.shapes.add_textbox(Inches(left), Inches(2.05), Inches(2.25), Inches(0.4))
        p_h = tb_h.text_frame.paragraphs[0]
        p_h.text = f"Stage {i+1}"
        p_h.alignment = PP_ALIGN.CENTER
        p_h.font.size = Pt(12)
        p_h.font.bold = True
        p_h.font.color.rgb = WHITE

        tb = s6.shapes.add_textbox(Inches(left + 0.1), Inches(2.6), Inches(2.05), Inches(2.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = stitle
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = scol

        p_d = tf.add_paragraph()
        p_d.text = sdesc
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = TEXT_DARK
        p_d.space_before = Pt(8)

        # Arrow
        if i < 4:
            arrow = s6.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(left + 2.27), Inches(3.3), Inches(0.12), Inches(0.2))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = ORANGE
            arrow.line.fill.background()

    # Data Stores Card at Bottom
    add_card(s6, 0.8, 5.4, 11.7, 1.4, "STATE & ARTIFACT STORAGE ARCHITECTURE", NAVY_HEADER)
    ds_box = s6.shapes.add_textbox(Inches(1.05), Inches(5.85), Inches(11.2), Inches(0.8))
    ds_tf = ds_box.text_frame
    ds_tf.word_wrap = True
    ds_p = ds_tf.paragraphs[0]
    ds_p.text = "• Input: data/Meridian_Vendor_Input.xlsx (Source of Truth, Immutable)  |  config/rubric.yaml  |  config/sources.yaml\n" \
               "• Snapshot Store: cache/<Vendor_ID>/<URL_Hash>.json (Raw HTML, Cleaned Text, Retrieval Date, SHA-256 Content Hash)\n" \
               "• Output Artifacts: data/Meridian_Vendor_Output.xlsx (Preserved B–K, Populated L–V, 132-row Evidence Log)  |  Streamlit UI"
    ds_p.font.size = Pt(11)
    ds_p.font.color.rgb = TEXT_DARK

    s6.notes_slide.notes_text_frame.text = (
        "This slide fulfills Expected Outcome 1. Notice the clean, decoupled modularity: M1 runs offline, "
        "M2 downloads and caches, M3 classifies and filters, M4 scores, and M5 exports. Everything is traceable "
        "to local hash-verified files."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 7: Expected Outcome 2 — System Architecture Diagram
    # ──────────────────────────────────────────────────────────────────────────
    s7 = setup_slide(is_dark=False)
    add_header(s7, "Expected Outcome 2: System Architecture Diagram", "Technical Component Interactions, Libraries, and Data Flow Pathways")

    # 3 Architectural Component Cards
    comps = [
        ("INGESTION & SCRAPING ENGINE", [
            ("criticality.py", "Parses input via pandas, executes 5 anchored rubric scales, applies floor rules and privileged access modifiers (+1)."),
            ("collectors.py", "Manages automated discovery (sitemap.xml / homepage), adheres to RFC 9309 robots.txt rules, enforces 1.2s rate limit."),
            ("trafilatura / bs4", "Strips scripts, styling, and navigation boilerplate. Generates standardized cleaned text for audit indexing.")
        ], NAVY_HEADER),
        ("AI CLASSIFIER & QUOTE GUARD", [
            ("classify.py", "Runs regex pre-filters with 25+ AI keywords to drop non-AI text and conserve computational resources."),
            ("Groq LPU API", "Leverages ultra-fast cloud inference (openai/gpt-oss-120b) with strict JSON schema response mode."),
            ("verify_quote()", "The Quote Guard: performs whitespace-normalized exact substring check. Any non-verbatim quote is discarded as INVALID_QUOTE.")
        ], ORANGE),
        ("AGGREGATION & EXPORT ENGINE", [
            ("risk_export.py", "Calculates weighted evidence scores, applies 180-day recency discount (0.7x), maps tier & footprint to risk."),
            ("openpyxl writer", "Uses shutil.copy to preserve client workbook formatting; populates columns L–V; appends 132-entry Evidence Log."),
            ("Streamlit UI", "Provides instant Presentation Mode, pre-computed KPI cards, per-vendor deep-dive cards, and one-click download.")
        ], NAVY_HEADER)
    ]

    for idx, (c_title, c_items, c_col) in enumerate(comps):
        left = 0.8 + idx * 4.0
        add_card(s7, left, 1.8, 3.75, 4.9, c_title, c_col)
        
        tb = s7.shapes.add_textbox(Inches(left + 0.2), Inches(2.4), Inches(3.35), Inches(4.1))
        tf = tb.text_frame
        tf.word_wrap = True

        for mod, desc in c_items:
            p = tf.add_paragraph() if tf.paragraphs[0].text else tf.paragraphs[0]
            p.text = f"⚙️ {mod}"
            p.font.size = Pt(13)
            p.font.bold = True
            p.font.color.rgb = c_col
            p.space_before = Pt(8)

            p_d = tf.add_paragraph()
            p_d.text = desc
            p_d.font.size = Pt(11)
            p_d.font.color.rgb = TEXT_DARK
            p_d.space_before = Pt(3)

    s7.notes_slide.notes_text_frame.text = (
        "Here is the technical architecture answering Expected Outcome 2. We use industry-standard libraries: "
        "Trafilatura for extraction, Groq LPU with Llama-3/GPT-OSS for classification, and openpyxl for Excel output. "
        "The Quote Guard ensures zero hallucination enters the audit trail."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 8: Expected Outcome 3 — Source Strategy & Evidentiary Weight Logic
    # ──────────────────────────────────────────────────────────────────────────
    s8 = setup_slide(is_dark=False)
    add_header(s8, "Expected Outcome 3: Source Strategy & Evidentiary Logic", "Tiered Reliability Hierarchy, Weighting Formula, and the Quote Guard Guarantee")

    # Left: 4 Source Tiers Table
    add_card(s8, 0.8, 1.8, 6.2, 4.9, "OSINT SOURCE RELIABILITY TIERS", NAVY_HEADER)
    
    # Table of tiers
    t_shape = s8.shapes.add_table(5, 3, Inches(1.0), Inches(2.4), Inches(5.8), Inches(4.0))
    table = t_shape.table
    table.columns[0].width = Inches(1.1)
    table.columns[1].width = Inches(3.6)
    table.columns[2].width = Inches(1.1)

    headers = ["Tier", "Source Type & Footprint", "Weight"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY_HEADER
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = WHITE

    tier_data = [
        ("Tier A\n(Direct)", "Product docs, sub-processor lists, AI policies, official regulatory filings.", "4.0x\n(Highest)"),
        ("Tier B\n(Corroborating)", "Technical job postings, engineering blogs, GitHub repos, developer SDKs.", "3.0x\n(Strong)"),
        ("Tier C\n(Contextual)", "Executive interviews, press releases, partnership announcements.", "2.0x\n(Moderate)"),
        ("Tier D\n(Noise)", "Generic marketing blogs, buzzwords, speculative thought pieces.", "1.0x\n(Filtered)")
    ]
    for row_idx, (t, s, w) in enumerate(tier_data, 1):
        for col_idx, val in enumerate([t, s, w]):
            cell = table.cell(row_idx, col_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if row_idx % 2 == 1 else RGBColor(241, 245, 249)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            p.font.color.rgb = TEXT_DARK

    # Right: Formula & Quote Guard Card
    add_card(s8, 7.3, 1.8, 5.2, 4.9, "MATHEMATICAL WEIGHTING & QUOTE GUARD", ORANGE)
    r_box = s8.shapes.add_textbox(Inches(7.55), Inches(2.35), Inches(4.7), Inches(4.1))
    r_tf = r_box.text_frame
    r_tf.word_wrap = True

    f_p = r_tf.paragraphs[0]
    f_p.text = "📐 Multi-Factor Weighting Formula:"
    f_p.font.size = Pt(13)
    f_p.font.bold = True
    f_p.font.color.rgb = ORANGE

    f_desc = r_tf.add_paragraph()
    f_desc.text = "Signal Score = Base × Tier × Conf × Linkage × Recency\n" \
                  "• Base: CONFIRMED (5) | CORROBORATED (3) | INDICATIVE (2) | MARKETING (0.5)\n" \
                  "• Service Linkage: 1.5x if tied to Meridian's product; 0.5x if general\n" \
                  "• Recency Discount: 0.7x penalty for signals > 180 days old"
    f_desc.font.size = Pt(11)
    f_desc.font.color.rgb = TEXT_DARK
    f_desc.space_before = Pt(6)

    qg_p = r_tf.add_paragraph()
    qg_p.text = "🛡️ The Cryptographic Quote Guard:"
    qg_p.font.size = Pt(13)
    qg_p.font.bold = True
    qg_p.font.color.rgb = NAVY_HEADER
    qg_p.space_before = Pt(14)

    qg_desc = r_tf.add_paragraph()
    qg_desc.text = "1. LLM extracts candidate verbatim quote & claims.\n" \
                   "2. Python executes exact substring check against stored snapshot.\n" \
                   "3. Whitespace & character encoding normalized.\n" \
                   "4. If candidate quote is NOT a 100% exact substring, the claim is rejected as INVALID_QUOTE. Zero hallucinations permitted."
    qg_desc.font.size = Pt(11)
    qg_desc.font.color.rgb = TEXT_DARK
    qg_desc.space_before = Pt(6)

    s8.notes_slide.notes_text_frame.text = (
        "Expected Outcome 3: Notice the rigor in our source strategy. We don't give the same weight to a marketing blog "
        "as we do to an official sub-processor list. And our Quote Guard discards any claim where the LLM paraphrased "
        "or hallucinated."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 9: Expected Outcome 4 — Criticality Scoring Rubric & Depth Policy
    # ──────────────────────────────────────────────────────────────────────────
    s9 = setup_slide(is_dark=False)
    add_header(s9, "Expected Outcome 4: Criticality Scoring Rubric", "Deterministic 0–3 Scoring Rubric, Floor Rules, and Tier-Driven Depth Policy")

    # Left: 5 Dimensions Card
    add_card(s9, 0.8, 1.8, 5.7, 4.9, "RUBRIC EVALUATION DIMENSIONS (0–3 SCALE)", NAVY_HEADER)
    d_box = s9.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(5.2), Inches(4.1))
    d_tf = d_box.text_frame
    d_tf.word_wrap = True

    dim_items = [
        ("Data Sensitivity (0–3):", "0 = No data | 1 = Internal business data | 2 = Confidential records | 3 = Full customer master records"),
        ("Payment-Flow (0–3):", "0 = No involvement | 1 = Reporting/indirect | 2 = Payment file transmission | 3 = Core clearing and settlement rail"),
        ("Regulatory Exposure (0–3):", "0 = Standard services | 1 = Reporting | 2 = Regulated banking | 3 = Systemically critical payment utility"),
        ("Operational Dependency (0–3):", "0 = Low | 1 = Moderate | 2 = High | 3 = Critical / Total operational dependence"),
        ("Data Volume (0–3):", "0 = Not applicable | 1 = < 100k records | 2 = 100k – 5M records | 3 = > 5M records / messages")
    ]
    for h, d in dim_items:
        bp = d_tf.add_paragraph() if d_tf.paragraphs[0].text else d_tf.paragraphs[0]
        bp.text = f"▪ {h} {d}"
        bp.font.size = Pt(11)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(8)

    # Right: Modifiers & Depth Policy
    add_card(s9, 6.8, 1.8, 5.7, 4.9, "FLOOR RULES & TIER-DRIVEN DEPTH POLICY", ORANGE)
    r_box = s9.shapes.add_textbox(Inches(7.05), Inches(2.35), Inches(5.2), Inches(4.1))
    r_tf = r_box.text_frame
    r_tf.word_wrap = True

    r_items = [
        ("Privileged-Access Modifier (+1):", "If vendor holds production credentials or job scheduling authority without direct customer data (e.g. AutomWorx), apply +1 score modifier."),
        ("Customer Master Floor Rule:", "Any vendor accessing full customer master data is guaranteed at least a High or Critical rating."),
        ("Critical Depth Policy:", "Full sweep: product docs, trust center, sub-processor list, AI policy, DPA pages, technical job postings, engineering blog, GitHub, filings."),
        ("High Depth Policy:", "Product docs, trust center, technical job postings, and regulatory filings."),
        ("Medium Depth Policy:", "Public website, standard product documentation, and general job listings."),
        ("Low Depth Policy:", "Trust center and AI-policy check only (efficient resource allocation).")
    ]
    for h, d in r_items:
        bp = r_tf.add_paragraph() if r_tf.paragraphs[0].text else r_tf.paragraphs[0]
        bp.text = f"▪ {h} {d}"
        bp.font.size = Pt(11)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(8)

    s9.notes_slide.notes_text_frame.text = (
        "Expected Outcome 4: Our criticality engine is deterministic. If another analyst runs the vendor through "
        "config/rubric.yaml, they reach the exact same rating. Depth of investigation is driven by criticality, "
        "not by how much marketing material a vendor happens to publish."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 10: Expected Outcome 4 (Cont.) — Comprehensive Vendor Criticality Ratings
    # ──────────────────────────────────────────────────────────────────────────
    s10 = setup_slide(is_dark=False)
    add_header(s10, "Expected Outcome 4 (Cont.): Vendor Criticality Scores", "Deterministic Scoring Breakdown Across Meridian's 6 Third-Party Vendors")

    # Table of all 6 vendors
    t_shape10 = s10.shapes.add_table(7, 6, Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
    tbl10 = t_shape10.table
    tbl10.columns[0].width = Inches(1.1)
    tbl10.columns[1].width = Inches(2.8)
    tbl10.columns[2].width = Inches(1.2)
    tbl10.columns[3].width = Inches(1.4)
    tbl10.columns[4].width = Inches(1.5)
    tbl10.columns[5].width = Inches(3.7)

    headers10 = ["Vendor ID", "Vendor Name", "Score", "Assigned Tier", "Dependency", "Criticality Rationale & Applied Depth"]
    for i, h in enumerate(headers10):
        cell = tbl10.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY_HEADER
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = WHITE

    v_crit_data = [
        ("V-001", "AutomWorx", "3", "Low", "High", "Holds production credentials (+1 privileged modifier) but no direct customer data. Depth: Trust center/AI policy only."),
        ("V-002", "Fiserv, Inc.", "15", "Critical", "Critical", "Core banking (DNA) & card platform accessing 1.4M customer master records. Depth: Full sweep (docs, trust, jobs, filings)."),
        ("V-003", "FSSI", "9", "High", "High", "Statement print & e-delivery processing for 850k customer accounts. Depth: Product docs, trust center, jobs, filings."),
        ("V-004", "Terrapin Technologies", "6", "Medium", "Moderate", "Wealth data solution consolidating 38k advisor accounts & commissions. Depth: Website, docs, and job listings."),
        ("V-005", "BNY", "9", "High", "High", "RTP transmission services processing 2.1M payment messages. Depth: Product docs, trust center, jobs, filings."),
        ("V-006", "The Clearing House", "12", "Critical", "Critical", "National ACH & RTP settlement clearing 14M payment messages. Depth: Full sweep (operating rules, specs, filings).")
    ]

    for row_idx, row in enumerate(v_crit_data, 1):
        for col_idx, val in enumerate(row):
            cell = tbl10.cell(row_idx, col_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if row_idx % 2 == 1 else RGBColor(241, 245, 249)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            if col_idx == 3:
                p.font.bold = True
                if val == "Critical":
                    p.font.color.rgb = RED_ACCENT
                elif val == "High":
                    p.font.color.rgb = AMBER_ACCENT
                elif val == "Medium":
                    p.font.color.rgb = BLUE_ACCENT
                else:
                    p.font.color.rgb = GREEN_ACCENT
            else:
                p.font.color.rgb = TEXT_DARK

    s10.notes_slide.notes_text_frame.text = (
        "Here is the complete scorecard for Meridian's 6 vendors. Notice the distribution: 2 Critical (Fiserv, TCH), "
        "2 High (FSSI, BNY), 1 Medium (Terrapin), and 1 Low (AutomWorx). Each vendor received a calibrated depth of "
        "investigation matching their risk potential."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 11: Expected Outcome 5 — Vendor Assessment Findings & Results Matrix
    # ──────────────────────────────────────────────────────────────────────────
    s11 = setup_slide(is_dark=False)
    add_header(s11, "Expected Outcome 5: Vendor Assessment Findings", "Definitive Findings Matrix: Exactly 1 Inconclusive Reflects Methodological Rigor")

    # Table of all 6 vendors with findings
    t_shape11 = s11.shapes.add_table(7, 6, Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
    tbl11 = t_shape11.table
    tbl11.columns[0].width = Inches(1.0)
    tbl11.columns[1].width = Inches(2.3)
    tbl11.columns[2].width = Inches(1.2)
    tbl11.columns[3].width = Inches(1.2)
    tbl11.columns[4].width = Inches(1.3)
    tbl11.columns[5].width = Inches(4.7)

    headers11 = ["VID", "Vendor Name", "Criticality", "AI Usage", "Risk Rating", "Identified Providers & Verified Footprint"]
    for i, h in enumerate(headers11):
        cell = tbl11.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY_HEADER
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = WHITE

    findings_data = [
        ("V-001", "AutomWorx", "Low", "Yes", "Medium", "Google (Gemini), OpenAI (ChatGPT). AI policy confirms consultants use commercial LLMs for scripting/work completion."),
        ("V-002", "Fiserv, Inc.", "Critical", "Yes", "Critical", "Azure OpenAI, Apple (Siri), Amazon (Alexa), Fiserv. 124 signals across AML Risk Manager, predictive analytics, agentic AI."),
        ("V-003", "FSSI", "High", "No", "None ident.", "N/A. Thorough OSINT sweep of 14 Tier A/B pages (SOC 2, trust center, print operations). Zero AI claims detected."),
        ("V-004", "Terrapin Tech.", "Medium", "Inconclusive", "Medium", "None named. Minimal web footprint (1 page reachable; API/docs 404). Honest inconclusive requiring targeted questionnaire."),
        ("V-005", "BNY", "High", "Yes", "High", "Microsoft, OpenAI, BNY AI Hub. Confirmed AI-powered document processing and workflow optimization in corporate platforms."),
        ("V-006", "The Clearing House", "Critical", "No", "None ident.", "N/A. Extensive audit of 1,016 pages of network rules & specs. Core ACH/RTP clearing rail operates deterministically without AI.")
    ]

    for row_idx, row in enumerate(findings_data, 1):
        for col_idx, val in enumerate(row):
            cell = tbl11.cell(row_idx, col_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if row_idx % 2 == 1 else RGBColor(241, 245, 249)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            if col_idx == 3:
                p.font.bold = True
                p.font.color.rgb = RED_ACCENT if val == "Yes" else GREEN_ACCENT if val == "No" else AMBER_ACCENT
            elif col_idx == 4:
                p.font.bold = True
                p.font.color.rgb = RED_ACCENT if val == "Critical" else AMBER_ACCENT if val == "High" else BLUE_ACCENT if val == "Medium" else GREEN_ACCENT
            else:
                p.font.color.rgb = TEXT_DARK

    s11.notes_slide.notes_text_frame.text = (
        "Crucial slide: We resolved previous ambiguities down to exactly ONE inconclusive vendor (Terrapin). "
        "Fiserv, BNY, and AutomWorx are confirmed 'Yes'; FSSI and The Clearing House are confirmed 'No' after auditing "
        "hundreds of pages. Terrapin is genuinely inconclusive due to bot blocks, showing we never guess without proof."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 12: Surfacing Enterprise Risk & Governance Action Playbook
    # ──────────────────────────────────────────────────────────────────────────
    s12 = setup_slide(is_dark=False)
    add_header(s12, "Surfacing Enterprise Risk & Governance Playbook", "Mapping Public AI Footprints to Meridian's Risk Framework & 15/30-Day Action Plans")

    # Card 1: 4 Enterprise Risk Vectors
    add_card(s12, 0.8, 1.8, 5.7, 4.9, "FOUR CORE ENTERPRISE AI RISKS", NAVY_HEADER)
    r1_box = s12.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(5.2), Inches(4.1))
    r1_tf = r1_box.text_frame
    r1_tf.word_wrap = True

    r_risks = [
        ("1. Sensitive Data Exposure:", "Commercial LLMs or unvetted AI pipelines ingesting Meridian customer master data (Fiserv: 1.4M records) or payment instructions (BNY: 2.1M messages)."),
        ("2. Shadow AI Sub-processors:", "Reliance on third-party cloud AI vendors (OpenAI, Microsoft Azure) without formal DPA inclusion or sub-processor registry listing."),
        ("3. Script & Credential Leakage:", "Consulting staff (AutomWorx) pasting proprietary job scheduling scripts or production system credentials into public LLMs (ChatGPT, Gemini)."),
        ("4. Contractual & Audit Voids:", "Absence of contractual guarantees prohibiting vendor AI models from training on Meridian proprietary data or guaranteeing rights to audit.")
    ]
    for h, d in r_risks:
        bp = r1_tf.add_paragraph() if r1_tf.paragraphs[0].text else r1_tf.paragraphs[0]
        bp.text = f"🚨 {h}\n{d}"
        bp.font.size = Pt(11)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(8)

    # Card 2: Action Playbook
    add_card(s12, 6.8, 1.8, 5.7, 4.9, "RECOMMENDED GOVERNANCE PLAYBOOK", ORANGE)
    r2_box = s12.shapes.add_textbox(Inches(7.05), Inches(2.35), Inches(5.2), Inches(4.1))
    r2_tf = r2_box.text_frame
    r2_tf.word_wrap = True

    r_actions = [
        ("Immediate 15-Day Targeted Questionnaires:", "Issue specific questionnaires to Fiserv, BNY, and AutomWorx covering: model providers, hosting region, data retention, and training restrictions."),
        ("AI Sub-Processor Inventory Registration:", "Formally register named providers (Microsoft, OpenAI, Google) with an unresolved-provider flag pending written confirmation."),
        ("Contractual Remediation Clauses:", "Add mandatory clauses at contract review: strict training prohibitions, mandatory 30-day sub-processor change notices, and AI audit rights."),
        ("30-Day TPRC Escalation Path:", "If vendor confirmations are not received within 30 days, escalate to Meridian Third-Party Risk Committee for operational containment.")
    ]
    for h, d in r_actions:
        bp = r2_tf.add_paragraph() if r2_tf.paragraphs[0].text else r2_tf.paragraphs[0]
        bp.text = f"✅ {h}\n{d}"
        bp.font.size = Pt(11)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(8)

    s12.notes_slide.notes_text_frame.text = (
        "Here is the risk surfacing required by the rubric. We translate the technical signals into actionable TPRM "
        "governance: 15-day questionnaires, sub-processor registry updates, contractual training restrictions, and "
        "a 30-day escalation path to the Risk Committee."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 13: Working Demonstration — Interactive Streamlit & Excel Deliverable
    # ──────────────────────────────────────────────────────────────────────────
    s13 = setup_slide(is_dark=False)
    add_header(s13, "Working Demonstration & Deliverables Summary", "Live Web Demonstration (Instant Presentation Mode) and Completed Audit Deliverables")

    # Card 1: Streamlit Demo
    add_card(s13, 0.8, 1.8, 5.7, 4.9, "STREAMLIT DEMO: INSTANT PRESENTATION MODE", ORANGE)
    ui_box = s13.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(5.2), Inches(4.1))
    ui_tf = ui_box.text_frame
    ui_tf.word_wrap = True

    ui_points = [
        ("Zero Live Latency:", "Automatically loads pre-computed, cryptographically verified results upon page load (http://localhost:8501). No waiting for web crawls during meetings."),
        ("Executive KPI Banner:", "Surfaces top metrics at a glance: 6 Assessed, 3 Confirmed AI, 2 Ruled Out, 1 Inconclusive, 1 Critical Risk."),
        ("Per-Vendor Evidence Cards:", "Expandable deep-dive cards showing exact verbatim quotes, source URLs, retrieval dates, and named model providers."),
        ("Live Execution Drawer:", "An optional expander allows live crawling and real-time Groq LPU inference on custom uploaded workbooks if requested."),
        ("One-Click Download:", "Direct download button for the completed Meridian_Vendor_Output.xlsx.")
    ]
    for h, d in ui_points:
        bp = ui_tf.add_paragraph() if ui_tf.paragraphs[0].text else ui_tf.paragraphs[0]
        bp.text = f"💻 {h} {d}"
        bp.font.size = Pt(11)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(8)

    # Card 2: Excel Deliverable
    add_card(s13, 6.8, 1.8, 5.7, 4.9, "COMPLETED EXCEL WORKBOOK DELIVERABLE", NAVY_HEADER)
    xl_box = s13.shapes.add_textbox(Inches(7.05), Inches(2.35), Inches(5.2), Inches(4.1))
    xl_tf = xl_box.text_frame
    xl_tf.word_wrap = True

    xl_points = [
        ("Columns B–K Integrity:", "100% byte-for-byte identical to client input. Verified by automated round-trip tests."),
        ("Example Row (V-000):", "Retained verbatim as requested in specification. Formats and styling preserved."),
        ("Columns L–V Populated:", "Criticality Tier, Rationale, Depth, AI Usage, Evidence, How Using, Providers, Risk, Risk Rationale, Recommended Action, Assessed By."),
        ("132-Entry Evidence Log Sheet:", "New audited sheet detailing Vendor ID, Source Tier, URL, Retrieval Date, Verbatim Excerpt, Signal Tag, and SHA-256 Snapshot Hash."),
        ("Automated Test Suite:", "18 unit tests (pytest) verifying criticality determinism, quote guard, weighting logic, and column B–K round-trip integrity.")
    ]
    for h, d in xl_points:
        bp = xl_tf.add_paragraph() if xl_tf.paragraphs[0].text else xl_tf.paragraphs[0]
        bp.text = f"📑 {h} {d}"
        bp.font.size = Pt(11)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(8)

    s13.notes_slide.notes_text_frame.text = (
        "Demonstration slide: The UI is currently live on localhost:8501 in presentation mode. We can switch over "
        "to the browser at any time to demonstrate the interactive evidence cards and download the audited workbook."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 14: Section Divider — Case Study 2 (Dark Navy)
    # ──────────────────────────────────────────────────────────────────────────
    s14 = setup_slide(is_dark=True, category_text="CASE STUDY 2  |  CADENCE FINANCIAL GROUP")
    
    div_box2 = s14.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(11.7), Inches(3.0))
    div_tf2 = div_box2.text_frame
    div_tf2.word_wrap = True
    
    dp2_tag = div_tf2.paragraphs[0]
    dp2_tag.text = "SECTION 02"
    dp2_tag.font.size = Pt(18)
    dp2_tag.font.bold = True
    dp2_tag.font.color.rgb = ORANGE

    dp2_title = div_tf2.add_paragraph()
    dp2_title.text = "Case Study 2: Text Analysis & PII Detection"
    dp2_title.font.size = Pt(38)
    dp2_title.font.bold = True
    dp2_title.font.color.rgb = WHITE
    dp2_title.space_before = Pt(10)

    dp2_desc = div_tf2.add_paragraph()
    dp2_desc.text = "Engineering a Privacy-Preserving Pipeline to Extract Text, Classify PII, and Redact Artifacts Pre-LLM"
    dp2_desc.font.size = Pt(18)
    dp2_desc.font.color.rgb = RGBColor(203, 213, 225)
    dp2_desc.space_before = Pt(15)

    s14.notes_slide.notes_text_frame.text = (
        "We now transition to Case Study 2: Text Analysis and PII Detection for Cadence Financial Group."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 15: Case Study 2 — Context, Policy & Problem Statement
    # ──────────────────────────────────────────────────────────────────────────
    s15 = setup_slide(is_dark=False)
    add_header(s15, "Case Study 2: Context, Policy & Problem Statement", "Cadence Financial Group Pre-LLM Data Sanitization Mandate")

    # Card 1: Context & Policy
    add_card(s15, 0.8, 1.8, 5.7, 4.9, "CADENCE FINANCIAL PROFILE & MANDATE", NAVY_HEADER)
    c_box2 = s15.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(5.2), Inches(4.1))
    c_tf2 = c_box2.text_frame
    c_tf2.word_wrap = True

    c_items2 = [
        ("Client Profile:", "Cadence Financial Group is a premier global financial institution (investment banking, wealth management, commercial banking) headquartered in New York."),
        ("The Document Agent:", "Cadence employees deploy an LLM-powered document analysis agent to interpret operational policy manuals, organizational decks, and risk evidence."),
        ("Strict Organizational Policy:", "Under internal compliance policies, Personally Identifiable Information (PII) is strictly restricted from entering AI models."),
        ("The Operational Need:", "All submitted documents must be automatically and reliably sanitized of PII prior to passing text to downstream LLM inference.")
    ]
    for h, d in c_items2:
        bp = c_tf2.add_paragraph() if c_tf2.paragraphs[0].text else c_tf2.paragraphs[0]
        bp.text = f"▪ {h} {d}"
        bp.font.size = Pt(12)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(10)

    # Card 2: Technical Challenges
    add_card(s15, 6.8, 1.8, 5.7, 4.9, "TECHNICAL CHALLENGES & EXPECTATIONS", ORANGE)
    t_box2 = s15.shapes.add_textbox(Inches(7.05), Inches(2.35), Inches(5.2), Inches(4.1))
    t_tf2 = t_box2.text_frame
    t_tf2.word_wrap = True

    t_items2 = [
        ("Heterogeneous Artifact Formats:", "Evidence artifacts arrive as scanned PDFs (policy manuals), organizational PowerPoint slides (org charts), and Word documents (training packs)."),
        ("OCR & Image Artifacts:", "Names, phone numbers, and IDs are frequently trapped inside scanned pages, tables, and rasterized diagram figures."),
        ("Structural Preservation (>80%):", "Extracted text must preserve paragraph flow, tables, and preceding/succeeding context so the downstream LLM understands the document."),
        ("Near-100% PII Coverage:", "Zero leakage tolerance: PII detection must approach 100% recall while keeping false positives minimal."),
        ("100% Local Execution:", "Pipeline must run entirely locally (Tesseract, spaCy, PyMuPDF) — zero document text can be dispatched to external cloud APIs.")
    ]
    for h, d in t_items2:
        bp = t_tf2.add_paragraph() if t_tf2.paragraphs[0].text else t_tf2.paragraphs[0]
        bp.text = f"▪ {h} {d}"
        bp.font.size = Pt(12)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(10)

    s15.notes_slide.notes_text_frame.text = (
        "Case Study 2 sets a very high bar: Financial institutions cannot send employee PII, client account numbers, "
        "or social security numbers to third-party LLMs. We had to build a local, highly accurate extraction and "
        "redaction pipeline."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 16: Case Study 2 — Roles & Responsibilities Framework
    # ──────────────────────────────────────────────────────────────────────────
    s16 = setup_slide(is_dark=False)
    add_header(s16, "Roles & Responsibilities Framework (Case Study 2)", "Three Strategic Pillars: File Structure Analysis, Solution Development & Exposure Surfacing")

    p2_cards = [
        ("01", "FILE STRUCTURE ANALYSIS", "Analyze the structural complexity of Cadence evidence artifacts (PDF, DOCX, PPTX). Determine format-specific parsing strategies, table extractors, and OCR fallback requirements for scanned figures.", NAVY_HEADER),
        ("02", "DEVELOP DETECTION SOLUTION", "Build a privacy-preserving extraction and detection pipeline. Combine high-precision regex detectors with contextual spaCy NER (en_core_web_lg), noise repair, and pixel-level document redaction.", ORANGE),
        ("03", "SURFACE PII EXPOSURE", "Quantify and classify personal data exposure (employee names, direct phone lines, emails, government IDs, bank details). Generate reversible audit token vaults and structured sanitization reports.", NAVY_HEADER)
    ]

    for idx, (num, p_title, p_desc, col) in enumerate(p2_cards):
        left = 0.8 + idx * 4.0
        add_card(s16, left, 1.8, 3.75, 4.9, f"PILLAR {num}", col)
        
        tb = s16.shapes.add_textbox(Inches(left + 0.2), Inches(2.4), Inches(3.35), Inches(4.1))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = p_title
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = col

        p2 = tf.add_paragraph()
        p2.text = p_desc
        p2.font.size = Pt(12)
        p2.font.color.rgb = TEXT_DARK
        p2.space_before = Pt(14)

    s16.notes_slide.notes_text_frame.text = (
        "These three pillars map directly to Slide 8 of the Optiv prompt: File structure analysis, developing the "
        "solution, and surfacing the resulting exposure."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 17: Expected Outcome 1 & 2 — Solution Blueprint & Architecture
    # ──────────────────────────────────────────────────────────────────────────
    s17 = setup_slide(is_dark=False)
    add_header(s17, "Expected Outcomes 1 & 2: Solution Architecture", "Local Pipeline Data Flow: Ingestion → Multi-Format Extraction → Hybrid NLP → Token Vault → Redaction")

    cs2_stages = [
        ("Stage 1: File Ingestion", "• PDF, DOCX, PPTX artifacts\n• MIME-type verification\n• Metadata extraction\n• Structural layout analysis", NAVY_HEADER),
        ("Stage 2: Multi-Extract", "• PyMuPDF / pdfplumber\n• python-docx & python-pptx\n• Tesseract OCR engine\n• Bounding box tracking", ORANGE),
        ("Stage 3: Hybrid NLP", "• Deterministic regex\n• spaCy NER (en_core_web_lg)\n• Financial gazetteers\n• Noise & hyphen repair", NAVY_HEADER),
        ("Stage 4: Token Vault", "• Reversible tokenization\n• Encrypted .vault.json\n• Span offset tracking\n• ORG_CONTACT policy", ORANGE),
        ("Stage 5: Redaction", "• Pixel-level PDF redaction\n• Text masking [TYPE_ID]\n• Structure preservation >80%\n• Sanitized LLM payload", NAVY_HEADER)
    ]

    for i, (stitle, sdesc, scol) in enumerate(cs2_stages):
        left = 0.8 + i * 2.4
        box = s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(2.0), Inches(2.25), Inches(3.2))
        box.fill.solid()
        box.fill.fore_color.rgb = CARD_BG
        box.line.color.rgb = scol
        box.line.width = Pt(2)

        hbar = s17.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(2.0), Inches(2.25), Inches(0.5))
        hbar.fill.solid()
        hbar.fill.fore_color.rgb = scol
        hbar.line.fill.background()

        tb_h = s17.shapes.add_textbox(Inches(left), Inches(2.05), Inches(2.25), Inches(0.4))
        p_h = tb_h.text_frame.paragraphs[0]
        p_h.text = f"Module {i+1}"
        p_h.alignment = PP_ALIGN.CENTER
        p_h.font.size = Pt(12)
        p_h.font.bold = True
        p_h.font.color.rgb = WHITE

        tb = s17.shapes.add_textbox(Inches(left + 0.1), Inches(2.6), Inches(2.05), Inches(2.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = stitle
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = scol

        p_d = tf.add_paragraph()
        p_d.text = sdesc
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = TEXT_DARK
        p_d.space_before = Pt(8)

        if i < 4:
            arrow = s17.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(left + 2.27), Inches(3.3), Inches(0.12), Inches(0.2))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = ORANGE
            arrow.line.fill.background()

    # Bottom Architectural Callout
    add_card(s17, 0.8, 5.4, 11.7, 1.4, "CRITICAL SECURITY ARCHITECTURE HIGHLIGHTS", NAVY_HEADER)
    c2_bot = s17.shapes.add_textbox(Inches(1.05), Inches(5.85), Inches(11.2), Inches(0.8))
    c2_tf = c2_bot.text_frame
    c2_tf.word_wrap = True
    c2_p = c2_tf.paragraphs[0]
    c2_p.text = "🔒 100% Air-Gapped Local Operation: Zero external API calls. All NLP, OCR, and redactions occur entirely inside Cadence's secure boundary.\n" \
               "🔒 True Pixel-Level Redaction: PyMuPDF draws black redaction rectangles and permanently purges underlying text stream bytes—preventing copy-paste discovery.\n" \
               "🔒 Structural Fidelity Guarantee: Retains document hierarchy, headers, table rows, and context tokens so the downstream LLM agent performs accurately."
    c2_p.font.size = Pt(11)
    c2_p.font.color.rgb = TEXT_DARK

    s17.notes_slide.notes_text_frame.text = (
        "Expected Outcomes 1 & 2 for Case Study 2: This architecture blueprint shows how artifacts move from ingestion "
        "to redaction. Notice Module 4: The Token Vault maintains a secure, encrypted map so authorized auditors can "
        "re-identify entities if needed, while downstream LLMs see only sanitized tokens."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 18: Expected Outcome 3 — Multi-Format Text Extraction & OCR Logic
    # ──────────────────────────────────────────────────────────────────────────
    s18 = setup_slide(is_dark=False)
    add_header(s18, "Expected Outcome 3: Extraction Logic Across Formats", "Format-Specific Parsing Strategies, Table Extraction, and Tesseract OCR Engine")

    fmt_cards = [
        ("PDF EXTRACTION STRATEGY", [
            ("Digital Text Layer:", "Extract native text streams with exact bounding box coordinates using PyMuPDF (fitz) and pdfplumber."),
            ("Layout & Table Retention:", "Detects multi-column policy layouts, headers, footers, and table gridlines to prevent jumbled word streams."),
            ("OCR Scanned Page Fallback:", "If page has < 50 extractable characters, triggers Tesseract OCR with image binarization and rotation correction.")
        ], NAVY_HEADER),
        ("DOCX & PPTX PARSING LOGIC", [
            ("python-docx Parser:", "Traverses paragraphs, nested tables, headers, footers, and callout text boxes while preserving document ordering."),
            ("python-pptx Parser:", "Inspects slide shapes, grouped diagram elements, smart art, speaker notes, and tabular slide matrices."),
            ("Embedded Figure OCR:", "Extracts rasterized screenshots, org chart images, and process diagrams; dispatches them to Tesseract OCR pass.")
        ], ORANGE),
        ("NOISE REPAIR & NORMALIZATION", [
            ("Hyphenation Stitching:", "Detects and repairs line-break split words (e.g., 'em-ployee' -> 'employee') to prevent broken PII entities."),
            ("Character Confusions:", "Corrects classic OCR confusions in numeric fields (e.g., resolving 'O' vs '0', 'l' vs '1' in phone/SSN spans)."),
            ("Context Window Padding:", "Maintains token offset anchors to preserve pre- and post-context for downstream entity disambiguation.")
        ], NAVY_HEADER)
    ]

    for idx, (f_title, f_items, f_col) in enumerate(fmt_cards):
        left = 0.8 + idx * 4.0
        add_card(s18, left, 1.8, 3.75, 4.9, f_title, f_col)
        
        tb = s18.shapes.add_textbox(Inches(left + 0.2), Inches(2.4), Inches(3.35), Inches(4.1))
        tf = tb.text_frame
        tf.word_wrap = True

        for mod, desc in f_items:
            p = tf.add_paragraph() if tf.paragraphs[0].text else tf.paragraphs[0]
            p.text = f"📄 {mod}"
            p.font.size = Pt(13)
            p.font.bold = True
            p.font.color.rgb = f_col
            p.space_before = Pt(8)

            p_d = tf.add_paragraph()
            p_d.text = desc
            p_d.font.size = Pt(11)
            p_d.font.color.rgb = TEXT_DARK
            p_d.space_before = Pt(3)

    s18.notes_slide.notes_text_frame.text = (
        "Expected Outcome 3: Handling real-world document messiness. PDF tables, PPTX org charts, and scanned Word "
        "documents require specialized extractors. Our noise repair engine resolves broken line-breaks before PII "
        "detection runs, ensuring names and numbers are not split."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 19: Expected Outcome 3 (Cont.) — PII Detection Engine & 11 Entity Classes
    # ──────────────────────────────────────────────────────────────────────────
    s19 = setup_slide(is_dark=False)
    add_header(s19, "Expected Outcome 3 (Cont.): Hybrid PII Detection", "Dual-Engine Architecture: High-Precision Regex + spaCy NER (en_core_web_lg)")

    # Left: 11 Entities Table
    add_card(s19, 0.8, 1.8, 6.2, 4.9, "11 PII ENTITY CATEGORIES TARGETED", NAVY_HEADER)
    
    t_shape19 = s19.shapes.add_table(6, 2, Inches(1.0), Inches(2.4), Inches(5.8), Inches(4.0))
    tbl19 = t_shape19.table
    tbl19.columns[0].width = Inches(2.2)
    tbl19.columns[1].width = Inches(3.6)

    h19 = ["PII Category", "Detection Engine & Example Patterns"]
    for i, h in enumerate(h19):
        cell = tbl19.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY_HEADER
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = WHITE

    pii_rows = [
        ("PERSON (Names)", "spaCy NER (en_core_web_lg) + Gazetteer lookups for executive lists."),
        ("GOV_ID (SSN / Tax)", "High-precision Regex: SSN (\\d{3}-\\d{2}-\\d{4}), EIN, ITIN."),
        ("FINANCIAL (Card/IBAN)", "Regex + Luhn checksum validation: Credit cards, bank IBAN, routing numbers."),
        ("CONTACT (Email/Phone)", "RFC 5322 Email regex + international E.164 phone number parsers."),
        ("ORG_CONTACT (Hotlines)", "Policy-driven: detect public hotlines/board contacts vs private employee PII.")
    ]
    for row_idx, (cat, eng) in enumerate(pii_rows, 1):
        for col_idx, val in enumerate([cat, eng]):
            cell = tbl19.cell(row_idx, col_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if row_idx % 2 == 1 else RGBColor(241, 245, 249)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            p.font.color.rgb = TEXT_DARK
            if col_idx == 0:
                p.font.bold = True

    # Right: Dual-Engine Synergy Card
    add_card(s19, 7.3, 1.8, 5.2, 4.9, "DUAL-ENGINE DETECTION ARCHITECTURE", ORANGE)
    de_box = s19.shapes.add_textbox(Inches(7.55), Inches(2.35), Inches(4.7), Inches(4.1))
    de_tf = de_box.text_frame
    de_tf.word_wrap = True

    de_points = [
        ("Engine 1: Deterministic Pattern Matcher", "Covers structured identifiers (SSN, credit card, phone, email, IP, dates). 100% precision with checksum validation (Luhn algorithm for cards) eliminates false positives."),
        ("Engine 2: Contextual spaCy NER", "Deploys `en_core_web_lg` transformer-tuned model for unstructured person names, organizational titles, and location entities where pattern syntax does not exist."),
        ("Disambiguation & Conflict Resolution", "If Regex and spaCy spans overlap, deterministic regex takes precedence. Span merging resolves adjacent tokens (e.g. 'Jane' + 'Doe' -> single entity)."),
        ("ORG_CONTACT Policy Switch", "Cadence compliance flag: `REDACT_ORG_CONTACTS=yes|no`. Allows retaining public whistle-blower hotlines while redacting personal desk extensions.")
    ]
    for h, d in de_points:
        bp = de_tf.add_paragraph() if de_tf.paragraphs[0].text else de_tf.paragraphs[0]
        bp.text = f"⚡ {h}\n{d}"
        bp.font.size = Pt(11)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(8)

    s19.notes_slide.notes_text_frame.text = (
        "Here is the detection core. Regex alone fails on names; NLP alone misses hyphenated phone numbers. "
        "By fusing deterministic regex with spaCy en_core_web_lg, we achieve industry-leading coverage across "
        "11 distinct PII categories."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 20: Expected Outcome 4 & 5 — Evaluation Benchmarks & Working Demo
    # ──────────────────────────────────────────────────────────────────────────
    s20 = setup_slide(is_dark=False)
    add_header(s20, "Expected Outcomes 4 & 5: Benchmarks & Live Demo", "Rigorous Performance Validation (>80% Structure, ~100% PII Recall) and Streamlit Demo")

    # Card 1: Benchmark Validation
    add_card(s20, 0.8, 1.8, 5.7, 4.9, "PERFORMANCE EVALUATION BENCHMARKS", NAVY_HEADER)
    bm_box = s20.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(5.2), Inches(4.1))
    bm_tf = bm_box.text_frame
    bm_tf.word_wrap = True

    bm_metrics = [
        ("Near-100% PII Recall:", "Zero tolerance for leakage. The hybrid detector achieves ~100% recall across gold-standard Cadence test artifacts (policies, org decks, training docs)."),
        ("Structure Retention (>80%):", "Verified via structural similarity metrics. Tables, headers, and bullet structures remain intact so the LLM agent retains complete reasoning capacity."),
        ("Context Preservation:", "Entities are replaced with semantic tokens (e.g. [PERSON_1], [SSN_1]) rather than blank spaces, allowing the LLM to understand subject-object relationships."),
        ("Encrypted Token Vault (.vault.json):", "Maintains an air-gapped cryptographic reverse lookup map, enabling authorized compliance auditors to unmask entities during legal discovery.")
    ]
    for h, d in bm_metrics:
        bp = bm_tf.add_paragraph() if bm_tf.paragraphs[0].text else bm_tf.paragraphs[0]
        bp.text = f"🎯 {h}\n{d}"
        bp.font.size = Pt(11)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(8)

    # Card 2: Interactive Demo UI
    add_card(s20, 6.8, 1.8, 5.7, 4.9, "CADENCE REDACTION DEMO UI", ORANGE)
    ui2_box = s20.shapes.add_textbox(Inches(7.05), Inches(2.35), Inches(5.2), Inches(4.1))
    ui2_tf = ui2_box.text_frame
    ui2_tf.word_wrap = True

    ui2_features = [
        ("Multi-Format Artifact Upload:", "Interactive drag-and-drop support for Cadence PDF, DOCX, and PPTX files directly in the browser."),
        ("Interactive Entity Tag Viewer:", "Color-coded findings map (🔴 PERSON, 🟡 GOV_ID, 🔵 FINANCIAL, 🟢 CONTACT) showing character offsets and confidence scores."),
        ("Side-by-Side Dual Preview:", "Split-screen viewer comparing raw unredacted artifact text against the sanitized downstream LLM payload in real time."),
        ("Audit Export & Download:", "One-click download of pixel-redacted PDFs, masked DOCX/PPTX files, and comprehensive PII audit telemetry logs.")
    ]
    for h, d in ui2_features:
        bp = ui2_tf.add_paragraph() if ui2_tf.paragraphs[0].text else ui2_tf.paragraphs[0]
        bp.text = f"💻 {h}\n{d}"
        bp.font.size = Pt(11)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(8)

    s20.notes_slide.notes_text_frame.text = (
        "Expected Outcomes 4 & 5: Notice that we replace PII with typed tokens like [PERSON_1] instead of deleting text. "
        "This is critical: the downstream LLM still understands that 'a person' approved the trade, but never learns "
        "their actual name or contact details."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 21: Strategic Synthesis: Holistic Enterprise AI Governance Model
    # ──────────────────────────────────────────────────────────────────────────
    s21 = setup_slide(is_dark=False)
    add_header(s21, "Strategic Synthesis: Enterprise AI Risk Framework", "Unifying Inbound Third-Party AI Profiling with Outbound Pre-LLM Data Sanitization")

    # Card 1: Dual Shield
    add_card(s21, 0.8, 1.8, 5.7, 4.9, "THE DUAL-SHIELD AI OPERATING MODEL", NAVY_HEADER)
    s_box = s21.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(5.2), Inches(4.1))
    s_tf = s_box.text_frame
    s_tf.word_wrap = True

    s_items = [
        ("Shield 1: Inbound Vendor Governance (Meridian)", "Automated OSINT profiling inspects whether partners introduce unvetted models into bank data flows. Enforces contractual safeguards, audit rights, and TPRC escalation."),
        ("Shield 2: Outbound Data Sanitization (Cadence)", "Local pre-LLM redaction firewall strips sensitive PII, customer account numbers, and employee credentials before internal agents ingest organizational artifacts."),
        ("Defense-in-Depth AI Architecture:", "Together, these two capabilities provide end-to-end protection: knowing where AI touches your data externally, and preventing confidential data from leaking internally.")
    ]
    for h, d in s_items:
        bp = s_tf.add_paragraph() if s_tf.paragraphs[0].text else s_tf.paragraphs[0]
        bp.text = f"🛡️ {h}\n{d}"
        bp.font.size = Pt(12)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(10)

    # Card 2: 90-Day Implementation Roadmap
    add_card(s21, 6.8, 1.8, 5.7, 4.9, "RECOMMENDED 90-DAY ROLLOUT ROADMAP", ORANGE)
    rm_box = s21.shapes.add_textbox(Inches(7.05), Inches(2.35), Inches(5.2), Inches(4.1))
    rm_tf = rm_box.text_frame
    rm_tf.word_wrap = True

    rm_items = [
        ("Days 1–30: Vendor Baseline & Questionnaires", "Deploy OSINT tool across Tier 1 vendor portfolio; issue targeted 15-day questionnaires to Fiserv, BNY, and AutomWorx; populate AI sub-processor register."),
        ("Days 31–60: Pre-LLM Redaction Gateway", "Integrate local PII redaction pipeline as a mandatory middleware proxy in front of Cadence's enterprise LLM document analysis agents."),
        ("Days 61–90: Continuous Governance & Audit", "Establish automated quarterly OSINT re-crawls, integrate token vault audits with Enterprise TPRM, and present finalized risk posture to the Board.")
    ]
    for h, d in rm_items:
        bp = rm_tf.add_paragraph() if rm_tf.paragraphs[0].text else rm_tf.paragraphs[0]
        bp.text = f"🗓️ {h}\n{d}"
        bp.font.size = Pt(12)
        bp.font.color.rgb = TEXT_DARK
        bp.space_before = Pt(10)

    s21.notes_slide.notes_text_frame.text = (
        "Strategic synthesis slide: This slide unifies both case studies into a single compelling narrative for "
        "Optiv leadership. We demonstrate that we don't just solve isolated exercises—we architect an enterprise-wide "
        "AI governance operating model."
    )

    # ──────────────────────────────────────────────────────────────────────────
    # SLIDE 22: Q&A, Discussion & Thank You (Dark Navy)
    # ──────────────────────────────────────────────────────────────────────────
    s22 = setup_slide(is_dark=True, category_text="DISCUSSION & NEXT STEPS  |  OPTIV CONSULTING")

    qbox = s22.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(4.8))
    qtf = qbox.text_frame
    qtf.word_wrap = True

    qp1 = qtf.paragraphs[0]
    qp1.text = "Discussion & Questions"
    qp1.font.size = Pt(36)
    qp1.font.bold = True
    qp1.font.color.rgb = ORANGE

    qp2 = qtf.add_paragraph()
    qp2.text = "Thank you for your time. We invite stakeholder feedback and questions on:\n\n" \
               "1. Vendor AI Scoring Thresholds & 15-Day Questionnaire Prioritization for Meridian Financial\n" \
               "2. Token Vault Security & ORG_CONTACT Redaction Policy Configuration for Cadence Financial Group\n" \
               "3. Integration Pathways into Existing GRC & Third-Party Risk Management Workflows\n\n"
    qp2.font.size = Pt(16)
    qp2.font.color.rgb = RGBColor(226, 232, 240)
    qp2.space_before = Pt(20)

    # Closing card
    q_card = s22.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(4.8), Inches(11.7), Inches(1.6))
    q_card.fill.solid()
    q_card.fill.fore_color.rgb = RGBColor(17, 34, 64)
    q_card.line.color.rgb = RGBColor(30, 58, 102)

    qc_box = s22.shapes.add_textbox(Inches(1.1), Inches(5.0), Inches(11.1), Inches(1.2))
    qc_tf = qc_box.text_frame
    qc_tf.word_wrap = True
    qc_p = qc_tf.paragraphs[0]
    qc_p.text = "OPTIV CONSULTING  •  CYBER STRATEGY ENGAGEMENT TEAM"
    qc_p.font.size = Pt(14)
    qc_p.font.bold = True
    qc_p.font.color.rgb = ORANGE

    qc_p2 = qc_tf.add_paragraph()
    qc_p2.text = "Deliverables: Meridian_Vendor_Output.xlsx  |  Cadence Redaction Engine  |  Streamlit Demonstration Apps\n" \
                 "Contact: optiv-strategy-team@optiv.com  |  Confidential  |  September 2026"
    qc_p2.font.size = Pt(13)
    qc_p2.font.color.rgb = RGBColor(203, 213, 225)
    qc_p2.space_before = Pt(8)

    s22.notes_slide.notes_text_frame.text = (
        "Thank you. We are now open for stakeholder questions and deep-dives into any technical or governance component."
    )

    output_path = "/Users/aimandhanani/Desktop/Useful/Optiv_proj/cs1/Optiv_Master_Case_Study_Presentation.pptx"
    prs.save(output_path)
    print(f"✅ Presentation successfully generated and saved to: {output_path}")

if __name__ == '__main__':
    create_deck()
