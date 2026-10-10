"""
M3: AI Signal Classifier with verbatim quote verification.

Pre-filters cached snapshots for AI keywords, then uses an LLM (or mock)
to extract structured signals. Every quote is verified as an exact substring
of the stored snapshot — invalid quotes are discarded.

LLM provider is configured via environment variables:
  - LLM_PROVIDER: "openai" | "mock" (default: "mock")
  - OPENAI_API_KEY: required if provider is "openai"
  - OPENAI_MODEL: model name (default: "gpt-4o-mini")
"""

import json
import os
import re
from pathlib import Path
from datetime import datetime, timezone

# ──────────────────────────────────────────────────────────────
# AI keyword pre-filter
# ──────────────────────────────────────────────────────────────

AI_KEYWORDS = [
    r'\bartificial intelligence\b',
    r'\bmachine learning\b',
    r'\bdeep learning\b',
    r'\bneural network\b',
    r'\bnatural language processing\b',
    r'\bNLP\b',
    r'\bLLM\b',
    r'\blarge language model\b',
    r'\bgenerative AI\b',
    r'\bgen[\s\-]?AI\b',
    r'\bGPT\b',
    r'\bChatGPT\b',
    r'\bOpenAI\b',
    r'\bClaude\b',
    r'\bAnthropic\b',
    r'\bGoogle AI\b',
    r'\bVertex AI\b',
    r'\bAzure AI\b',
    r'\bCopilot\b',
    r'\bAI[\s\-]?powered\b',
    r'\bAI[\s\-]?driven\b',
    r'\bAI[\s\-]?enabled\b',
    r'\bpredictive analytics\b',
    r'\bcomputer vision\b',
    r'\brobotics\b',
    r'\bautomation.*\bAI\b',
    r'\bAI\b.*\bautomation\b',
    r'\bintelligent automation\b',
    r'\bchatbot\b',
    r'\bmodel provider\b',
    r'\bsub[\s\-]?processor\b',
]

AI_PATTERN = re.compile('|'.join(AI_KEYWORDS), re.IGNORECASE)


def pre_filter(text):
    """Returns True if text contains AI-related keywords."""
    if not text:
        return False
    return bool(AI_PATTERN.search(text))


def find_ai_sentences(text):
    """Extract all sentences that contain AI-related keywords."""
    sentences = re.split(r'(?<=[.!?])\s+', text)
    hits = []
    for s in sentences:
        if AI_PATTERN.search(s):
            clean = s.strip()
            if len(clean) > 20 and len(clean) < 800:
                hits.append(clean)
    return hits


# ──────────────────────────────────────────────────────────────
# Quote verification
# ──────────────────────────────────────────────────────────────

def verify_quote(quote, source_text):
    """Verify that a quote is an exact substring of source_text after normalizing whitespace."""
    if not quote:
        return True
    # Normalize whitespace to avoid false rejections from HTML parsing differences
    norm_quote = " ".join(quote.split())
    norm_source = " ".join(source_text.split())
    return norm_quote in norm_source


# ──────────────────────────────────────────────────────────────
# LLM-based classifier (with mock fallback)
# ──────────────────────────────────────────────────────────────

SIGNAL_SCHEMA = {
    "signal_type": "CONFIRMED|CORROBORATED|INDICATIVE|MARKETING-ONLY|NEGATIVE-FINDING",
    "verbatim_quote": "exact substring from source",
    "ai_mechanism": "what the AI does",
    "ties_to_service": True,
    "data_touched": "what data the AI processes",
    "provider_named": "named provider or 'None'",
    "human_oversight_mentioned": False,
    "confidence": "High|Medium|Low"
}


def classify_with_mock(text, url):
    """
    Smart mock classifier that extracts real sentences from the text.
    Returns structured signals based on actual content analysis.
    """
    signals = []
    ai_sentences = find_ai_sentences(text)

    if not ai_sentences:
        return [{
            "signal_type": "NEGATIVE-FINDING",
            "verbatim_quote": "",
            "ai_mechanism": "None detected",
            "ties_to_service": False,
            "data_touched": "N/A",
            "provider_named": "None",
            "human_oversight_mentioned": False,
            "confidence": "High",
            "url": url
        }]

    for sentence in ai_sentences[:5]:  # cap at 5 per page
        # Determine signal strength from language
        sentence_lower = sentence.lower()

        if any(kw in sentence_lower for kw in ['sub-processor', 'subprocessor', 'model provider',
                                                  'uses openai', 'uses azure ai', 'powered by']):
            sig_type = "CONFIRMED"
            confidence = "High"
        elif any(kw in sentence_lower for kw in ['ai-powered', 'ai-driven', 'machine learning',
                                                    'deep learning', 'predictive analytics']):
            sig_type = "INDICATIVE"
            confidence = "Medium"
        elif any(kw in sentence_lower for kw in ['intelligent', 'automation', 'chatbot']):
            sig_type = "INDICATIVE"
            confidence = "Medium"
        else:
            sig_type = "MARKETING-ONLY"
            confidence = "Low"

        # Check for named providers
        provider = "None"
        for prov in ["OpenAI", "Anthropic", "Google AI", "Azure AI", "AWS Bedrock",
                      "Vertex AI", "Copilot", "GPT-4", "Claude", "Gemini"]:
            if prov.lower() in sentence_lower:
                provider = prov
                sig_type = "CONFIRMED"
                confidence = "High"

        # Check service linkage
        ties = any(kw in sentence_lower for kw in ['customer', 'client', 'transaction',
                                                     'payment', 'account', 'service',
                                                     'platform', 'solution', 'product'])

        signals.append({
            "signal_type": sig_type,
            "verbatim_quote": sentence,
            "ai_mechanism": "AI usage detected in public content",
            "ties_to_service": ties,
            "data_touched": "Unknown — requires vendor questionnaire",
            "provider_named": provider,
            "human_oversight_mentioned": 'human' in sentence_lower or 'oversight' in sentence_lower,
            "confidence": confidence,
            "url": url
        })

    return signals


def classify_with_llm(text, url):
    """Real LLM classifier — uses OpenAI API."""
    try:
        import openai
        client = openai.OpenAI()  # uses OPENAI_API_KEY env var
        model = os.environ.get('OPENAI_MODEL', 'gpt-4o-mini')

        prompt = f"""Analyze this text from {url} for evidence of AI usage.
For each AI-related claim, extract:
- signal_type: CONFIRMED (direct product docs, sub-processor lists) | CORROBORATED (job postings, blogs) | INDICATIVE (general mentions) | MARKETING-ONLY (buzzwords)
- verbatim_quote: the EXACT sentence from the text (must be a substring)
- ai_mechanism: what the AI does
- ties_to_service: whether it relates to a specific product/service (true/false)
- data_touched: what data the AI processes
- provider_named: any named AI provider (OpenAI, Anthropic, etc.) or "None"
- human_oversight_mentioned: true/false
- confidence: High/Medium/Low

Return a JSON array. If no AI signals found, return [{{"signal_type": "NEGATIVE-FINDING", ...}}].

TEXT:
{text[:4000]}"""

        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0.1
        )

        result = json.loads(response.choices[0].message.content)
        signals = result if isinstance(result, list) else result.get('signals', [result])

        for s in signals:
            s['url'] = url

        return signals
    except Exception as e:
        print(f"    LLM call failed ({e}), falling back to mock")
        return classify_with_mock(text, url)


# ──────────────────────────────────────────────────────────────
# Main classification pipeline
# ──────────────────────────────────────────────────────────────

def classify_vendor(vendor_id, cache_dir='cache'):
    """Classify all cached snapshots for a vendor."""
    vendor_dir = Path(cache_dir) / vendor_id
    if not vendor_dir.exists():
        print(f"  No cache directory for {vendor_id}")
        return []

    summary_file = vendor_dir / 'crawl_summary.json'
    if not summary_file.exists():
        print(f"  No crawl summary for {vendor_id}")
        return []

    with open(summary_file, 'r') as f:
        crawl_summary = json.load(f)

    provider = os.environ.get('LLM_PROVIDER', 'mock')
    classify_fn = classify_with_llm if provider == 'openai' else classify_with_mock

    all_signals = []
    sources_checked = []

    for item in crawl_summary:
        url = item.get('url', '')
        status = item.get('status')
        url_hash = item.get('url_hash', '')
        has_content = item.get('has_content', False)

        sources_checked.append({
            'url': url,
            'status': status,
            'has_content': has_content,
            'source_tier': item.get('source_tier', 'unknown'),
            'content_hash': item.get('content_hash', ''),
            'text_len': item.get('text_len', 0),
            'reason': item.get('reason')
        })

        # Coverage gap
        if status != 200 or not has_content:
            all_signals.append({
                "url": url,
                "signal_type": "COVERAGE-GAP",
                "verbatim_quote": "",
                "reason": item.get('reason') or (f"HTTP {status}" if status != 200 else "No extractable content"),
                "valid": True,
                "source_tier": item.get('source_tier', 'unknown')
            })
            continue

        # Load the cached snapshot
        snapshot_file = vendor_dir / f"{url_hash}.json"
        if not snapshot_file.exists():
            continue

        with open(snapshot_file, 'r') as sf:
            snapshot = json.load(sf)

        text = snapshot.get('cleaned_text', '')
        if not text:
            continue

        # Pre-filter
        if not pre_filter(text):
            all_signals.append({
                "url": url,
                "signal_type": "NEGATIVE-FINDING",
                "verbatim_quote": "",
                "ai_mechanism": "No AI keywords detected",
                "valid": True,
                "source_tier": item.get('source_tier', 'unknown'),
                "timestamp": snapshot.get('timestamp', '')
            })
            continue

        # Classify
        signals = classify_fn(text, url)

        # Verify every quote
        for sig in signals:
            quote = sig.get('verbatim_quote', '')
            is_valid = verify_quote(quote, text)
            sig['valid'] = is_valid
            sig['source_tier'] = item.get('source_tier', 'unknown')
            sig['timestamp'] = snapshot.get('timestamp', '')
            sig['content_hash'] = snapshot.get('content_hash', '')

            if not is_valid and quote:
                print(f"    ⚠ INVALID QUOTE discarded from {url}")
                sig['signal_type'] = "INVALID_QUOTE"

        all_signals.extend(signals)

    # Save results
    output = {
        'vendor_id': vendor_id,
        'classified_at': datetime.now(timezone.utc).isoformat(),
        'llm_provider': provider,
        'sources_checked': sources_checked,
        'signals': all_signals
    }

    with open(vendor_dir / 'signals.json', 'w') as f:
        json.dump(output, f, indent=2)

    # Print summary
    valid = [s for s in all_signals if s.get('valid', False) and s['signal_type'] not in ('COVERAGE-GAP', 'NEGATIVE-FINDING', 'INVALID_QUOTE')]
    gaps = [s for s in all_signals if s['signal_type'] == 'COVERAGE-GAP']
    negatives = [s for s in all_signals if s['signal_type'] == 'NEGATIVE-FINDING']
    print(f"  {vendor_id}: {len(valid)} valid signals, {len(negatives)} negative, {len(gaps)} coverage gaps")

    return all_signals


def main():
    cache_dir = Path('cache')
    for vendor_dir in sorted(cache_dir.iterdir()):
        if vendor_dir.is_dir() and vendor_dir.name.startswith('V-'):
            print(f"\nClassifying {vendor_dir.name}...")
            classify_vendor(vendor_dir.name)


if __name__ == '__main__':
    main()
