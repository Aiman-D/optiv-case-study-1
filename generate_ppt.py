from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def add_detailed_slide(prs, title_text, subtitle_text, body_text):
    slide_layout = prs.slide_layouts[1] # Title and Content
    slide = prs.slides.add_slide(slide_layout)
    
    # Title
    title = slide.shapes.title
    title.text = title_text
    title.text_frame.paragraphs[0].font.color.rgb = RGBColor(237, 107, 45) # Optiv Orange
    title.text_frame.paragraphs[0].font.size = Pt(36)
    
    # Body
    content = slide.placeholders[1]
    tf = content.text_frame
    tf.clear() # Clear default
    
    # Add subtitle
    p_sub = tf.paragraphs[0]
    p_sub.text = subtitle_text
    p_sub.font.bold = True
    p_sub.font.size = Pt(24)
    p_sub.font.color.rgb = RGBColor(20, 43, 76) # Navy
    
    # Add body paragraphs
    for idx, paragraph_text in enumerate(body_text.split('\n\n')):
        p = tf.add_paragraph()
        p.text = paragraph_text
        p.font.size = Pt(18)
        p.space_after = Pt(14)
        p.level = 0
        p.font.color.rgb = RGBColor(0, 0, 0)
        
    return slide

def create_presentation():
    prs = Presentation()

    NAVY_BLUE = RGBColor(20, 43, 76)
    ORANGE = RGBColor(237, 107, 45)

    # Title Slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "Vendor AI Usage Analysis Tool"
    title.text_frame.paragraphs[0].font.color.rgb = NAVY_BLUE
    subtitle.text = "Comprehensive Project Overview & Outcomes\nDesigned for Standalone Review\nOptiv Consulting"

    # Executive Summary
    add_detailed_slide(prs, "Executive Summary", 
        "Automated OSINT-Driven Vendor Assessment",
        "This project delivers a fully automated, evidence-driven tool to determine whether third-party vendors are utilizing Artificial Intelligence within the services they provide to Meridian Financial.\n\n"
        "To ensure compliance and respect vendor boundaries, the tool relies 100% on Open Source Intelligence (OSINT). We do not contact the vendors directly, bypass paywalls, or scrape authenticated platforms.\n\n"
        "The system reads a list of vendors, determines their criticality to the business, automatically scours the internet for relevant documents, uses an LLM to identify verbatim proof of AI usage, and outputs a classified risk report."
    )

    # 1. Determine Criticality & Ratings
    add_detailed_slide(prs, "01. Determining Vendor Criticality",
        "Role: Judge how critical each vendor is based on strict, repeatable criteria.",
        "Not all vendors require the same level of scrutiny. We built a 'Criticality Engine' (Module 1) that evaluates vendors using a deterministic scoring rubric (0-3 scales). This ensures another analyst would always arrive at the same rating.\n\n"
        "The criteria evaluated include:\n"
        "• Data Sensitivity (e.g., access to full customer master vs. no data)\n"
        "• Payment-Flow Involvement (e.g., core account processing vs. indirect)\n"
        "• Regulatory Exposure & Operational Dependency\n"
        "• Data Volume\n\n"
        "Based on these scores, vendors are assigned a final Criticality Rating: Critical, High, Medium, or Low."
    )

    # 2. Set the Depth
    add_detailed_slide(prs, "02. Setting Assessment Depth",
        "Role: Decide how much assessment depth each level of criticality warrants.",
        "The Criticality Rating directly drives the 'Depth Policy'—how hard we look for evidence. This optimizes computing resources and assessment time, ensuring we don't over-investigate low-risk vendors.\n\n"
        "Depth by Tier:\n"
        "• Critical: Full internet sweep (product docs, trust centers, AI policies, security pages, technical job postings, engineering blogs, GitHub/API docs, regulatory filings, press releases, executive channels).\n"
        "• High: Product docs, trust centers, jobs, and filings.\n"
        "• Medium: General website, basic docs, and jobs.\n"
        "• Low: Trust center and AI-policy check only."
    )

    # 3. Identify Sources & Strategy
    add_detailed_slide(prs, "03. Source Identification & Strategy",
        "Role: Identify OSINT sources that reveal genuine AI usage and apply evidentiary weight.",
        "We developed a structured Source Strategy that prioritizes high-confidence evidence. A large vendor might use AI internally (e.g., HR), but we must determine if AI is used *in the specific service* Meridian buys.\n\n"
        "We categorized sources into four tiers:\n"
        "• Tier A (Direct): Product docs, sub-processor lists, AI policies (Highest Weight)\n"
        "• Tier B (Corroborating): Tech jobs, engineering blogs, SDKs\n"
        "• Tier C (Contextual): Executive posts, press releases\n"
        "• Tier D (Noise): Generic marketing buzzwords (Ignored)\n\n"
        "The system calculates a Signal Weight = Reliability × Service-Linkage × Recency × Corroboration. "
        "A critical rule is enforced: Every piece of evidence must be an exact, verbatim quote from the source document to prevent AI hallucinations."
    )

    # 4. Build the Capability: Functional Flow
    add_detailed_slide(prs, "04. Capability Build: Functional Flow",
        "Role: Scrape, extract, analyze, classify, and tag information while preserving traceability.",
        "The Functional Flow represents the journey of data through the system:\n\n"
        "1. Input: Reads 'Meridian_Vendor_Input.xlsx'.\n"
        "2. M1 (Criticality Engine): Scores the vendor to assign a Tier and Depth.\n"
        "3. M2 (OSINT Collectors): Crawls public footprints respecting robots.txt (1.2s rate limits) and saves cached snapshots of every page.\n"
        "4. M3 (AI Classifier): Uses regex pre-filtering and an LLM to classify AI signals. It enforces an 'exact substring check' to guarantee the quote is real.\n"
        "5. M4 & M5 (Aggregation & Export): Calculates the final Yes/No/Inconclusive risk score and exports to a new Excel file."
    )

    # 5. Build the Capability: Architecture
    add_detailed_slide(prs, "04. Capability Build: Architecture",
        "Technical components and system interactions",
        "Our architecture ensures modularity and testability. The system is orchestrated via a Streamlit web interface or Jupyter Notebooks.\n\n"
        "Technical Highlights:\n"
        "• Data Collection: Powered by 'trafilatura' and 'BeautifulSoup' for robust text extraction. Any unreachable site is safely logged as a 'Coverage Gap', not a false negative.\n"
        "• Classification: Uses strict JSON schemas with LLMs (e.g., OpenAI) to extract: signal type, verbatim quote, AI mechanism, and data touched.\n"
        "• Storage: All downloaded pages are cached locally by URL hash, enabling fully offline demonstrations and complete audit trails."
    )

    # 6. Surface the Risk & Demo
    add_detailed_slide(prs, "05. Surface the Risk & Working Demo",
        "Role: Surface resulting risks to support the vendor due diligence process.",
        "The final stage (M4 Risk Engine) takes the verified AI signals and maps them to Meridian's risk framework (Critical/High/Medium/None).\n\n"
        "Identified risks include:\n"
        "• Sensitive data exposed to AI models\n"
        "• Reliance on undisclosed AI sub-processors\n"
        "• Missing AI contract clauses\n\n"
        "Working Demonstration:\n"
        "We have provided a Streamlit UI (app/main.py) and a Jupyter Notebook (pipeline_demo.py). These interfaces allow users to upload the vendor list, watch the analysis run in real-time with 'evidence cards' (showing verbatim quotes, source URLs, and retrieval dates), and download the final Excel report."
    )
    
    # 7. Outputs
    add_detailed_slide(prs, "Final Deliverables",
        "What the system produces for the user",
        "1. The Output Excel File ('Meridian_Vendor_Output.xlsx'):\n"
        "The system copies the original input file byte-for-byte to preserve formatting, then safely populates columns L-V with the AI Usage Detection, Risk Classification, Rationale, and Recommended Actions.\n\n"
        "2. The Evidence Log:\n"
        "A brand new 'Evidence Log' sheet is automatically appended to the workbook. This provides a complete audit trail containing the Source URL, Vendor Tier, Retrieval Date, Verbatim Excerpt, Classification Tag, and the cryptographic hash of the cached snapshot."
    )

    prs.save('Optiv_Standalone_Presentation.pptx')
    print("Presentation created successfully as 'Optiv_Standalone_Presentation.pptx'")

if __name__ == '__main__':
    create_presentation()
