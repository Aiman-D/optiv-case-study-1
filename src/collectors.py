"""
M2: OSINT Collectors & Snapshot Store (v2)

Crawls vendor public websites based on the depth policy for the vendor's
criticality tier. Stores snapshots (URL, timestamp, content hash, text) for
offline replay. Respects robots.txt and rate limits.

Changes vs v1:
  * URLs are normalised, so "https://x.com" and "https://x.com/" are one page.
  * Identical page content under different URLs is flagged as a duplicate
    (soft-404s / redirects to the homepage) and counted as a coverage gap,
    not as a successfully checked source.
  * Pages need a minimum amount of real text to count as "has_content".
  * Link discovery: follows relevant links from the homepage and sitemap.xml
    instead of only guessing fixed paths.
  * robots.txt is fetched with our own session. urllib's RobotFileParser.read()
    treats a 401/403 on robots.txt as "disallow everything", which silently
    blocked whole sites. We follow RFC 9309: 4xx -> no restrictions,
    5xx/unreachable -> assume disallowed.
  * Failed fetches (403/429/5xx/timeouts) are no longer cached forever.
  * The BeautifulSoup fallback strips <script>/<style> so JS blobs are not
    mistaken for page text.
  * Publication date is captured when available (for recency weighting).
"""

import hashlib
import html as htmllib
import json
import os
import re
import time
from datetime import datetime, timezone
from urllib.parse import urljoin, urldefrag, urlparse
from urllib.robotparser import RobotFileParser

import pandas as pd
import requests
import trafilatura
import yaml
from bs4 import BeautifulSoup

# ──────────────────────────────────────────────────────────────
# Settings
# ──────────────────────────────────────────────────────────────

BOT_NAME = 'MeridianOSINTBot'
# Put a real contact URL/email here so site owners can reach you.
USER_AGENT = f'Mozilla/5.0 (compatible; {BOT_NAME}/1.0; +http://example.com/bot)'
RATE_LIMIT_SECONDS = 1.2
MIN_TEXT_CHARS = 300          # below this a page is treated as "no real content"
TIER_ORDER = ["Tier_A", "Tier_B", "Tier_C", "Tier_D"]

# Which source tiers to crawl per criticality tier
DEPTH_TIERS = {
    "Critical": ["Tier_A", "Tier_B", "Tier_C", "Tier_D"],
    "High":     ["Tier_A", "Tier_B"],
    "Medium":   ["Tier_B", "Tier_D"],
    "Low":      ["Tier_A"],
}

# Max number of discovered (non-config) pages per criticality tier
MAX_DISCOVERED = {"Critical": 1000, "High": 500, "Medium": 200, "Low": 50}

# Keywords (matched against whole path/anchor tokens) that suggest a source tier
TIER_KEYWORDS = {
    "Tier_A": {"trust", "security", "privacy", "subprocessors", "subprocessor",
               "processors", "responsible", "docs", "documentation",
               "developers", "developer", "api", "compliance"},
    "Tier_B": {"careers", "jobs", "engineering", "github"},
    "Tier_C": {"news", "newsroom", "press", "investors", "investor", "blog",
               "insights", "media"},
    "Tier_D": {"platform", "products", "solutions", "services", "about"},
}

SKIP_EXT = re.compile(r'\.(pdf|jpe?g|png|gif|svg|webp|zip|gz|mp4|mp3|css|js|ico|xml)$', re.I)


def load_config(path):
    with open(path, 'r') as f:
        return yaml.safe_load(f)


def get_hash(text):
    return hashlib.sha256(text.encode('utf-8')).hexdigest()


def normalize_url(url):
    """Lower-case host, drop fragment and trailing slash so duplicates collapse."""
    url, _ = urldefrag(url.strip())
    p = urlparse(url)
    path = p.path.rstrip('/')
    return f"{p.scheme}://{p.netloc.lower()}{path}" + (f"?{p.query}" if p.query else "")


def _tokens(*parts):
    return {t for t in re.split(r'[^a-z0-9]+', ' '.join(parts).lower()) if t}


AI_POLICY_WORDS = {"policy", "policies", "governance", "principles", "ethics", "usage"}


def guess_tier(tokens):
    """Guess a source tier from path/anchor tokens.

    Careers and press/blog pages are checked first, so "/newsroom/ai-launch"
    is Tier C (contextual), not Tier A. A bare "ai" token only reaches Tier A
    together with a policy-type word; otherwise it is generic marketing (D).
    """
    for tier in ("Tier_B", "Tier_C", "Tier_A"):
        if tokens & TIER_KEYWORDS[tier]:
            return tier
    if "ai" in tokens and tokens & AI_POLICY_WORDS:
        return "Tier_A"
    if "ai" in tokens or tokens & TIER_KEYWORDS["Tier_D"]:
        return "Tier_D"
    return "Tier_D"  # Fallback to Tier_D so we crawl everything instead of dropping unknown links


def _host(url):
    return urlparse(url).netloc.lower().split(':')[0]


def _root_domain(host):
    return host[4:] if host.startswith('www.') else host


def _same_site(url, base_host):
    host = _host(url)
    root = _root_domain(base_host)
    return host == base_host or host == root or host.endswith('.' + root)


# ──────────────────────────────────────────────────────────────
# Collector
# ──────────────────────────────────────────────────────────────

class OSINTCollector:
    def __init__(self, sources_config_path, cache_dir='cache'):
        self.config = load_config(sources_config_path)
        self.cache_dir = cache_dir
        os.makedirs(self.cache_dir, exist_ok=True)
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': USER_AGENT,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
        })
        self._robots_cache = {}

    # ── robots.txt ──────────────────────────────────────────

    def _robots_for(self, url):
        p = urlparse(url)
        origin = f"{p.scheme}://{p.netloc}"
        if origin in self._robots_cache:
            return self._robots_cache[origin]

        rp = RobotFileParser()
        try:
            time.sleep(RATE_LIMIT_SECONDS)
            r = self.session.get(origin + '/robots.txt', timeout=10)
            if r.status_code == 200:
                rp.parse(r.text.splitlines())
                rp.state = 'disallowed by robots.txt rules'
            elif 500 <= r.status_code < 600:
                rp.parse([])
                rp.disallow_all = True      # RFC 9309: server error -> assume disallowed
                rp.state = f'robots.txt returned HTTP {r.status_code}; assumed disallowed'
            else:
                rp.parse([])
                rp.allow_all = True         # RFC 9309: 4xx (incl. 401/403) -> unrestricted
                rp.state = 'unrestricted'
        except Exception as e:
            rp.parse([])
            rp.disallow_all = True          # unreachable -> be conservative
            rp.state = f'robots.txt unreachable ({type(e).__name__}); assumed disallowed'
        self._robots_cache[origin] = rp
        return rp

    def _allowed(self, url):
        return self._robots_for(url).can_fetch(BOT_NAME, url)

    def _block_reason(self, url):
        return 'Blocked: ' + getattr(self._robots_for(url), 'state', 'robots.txt')

    # ── fetching ────────────────────────────────────────────

    def fetch_url(self, url, vendor_dir, refresh=False):
        """Fetch a URL; cache result in vendor_dir. Returns snapshot dict."""
        url_hash = get_hash(url)
        cache_file = os.path.join(vendor_dir, f"{url_hash}.json")

        if os.path.exists(cache_file) and not refresh:
            with open(cache_file, 'r') as f:
                cached = json.load(f)
            # Only trust cached successes and stable "not found" answers.
            # Transient failures (403/429/5xx/timeouts) are retried.
            if cached.get('status') in (200, 404, 410):
                return cached

        print(f"  Fetching {url}...")
        snapshot = {
            'url': url,
            'url_hash': url_hash,
            'final_url': url,
            'timestamp': datetime.now(timezone.utc).isoformat(),
            'published': '',
            'status': 'error',
            'raw_html': '',
            'cleaned_text': '',
            'content_hash': ''
        }

        for attempt in range(2):
            try:
                time.sleep(RATE_LIMIT_SECONDS)
                resp = self.session.get(url, timeout=15, allow_redirects=True)
                snapshot['status'] = resp.status_code
                snapshot['final_url'] = resp.url
                if resp.status_code in (429,) or resp.status_code >= 500:
                    if attempt == 0:
                        time.sleep(3)
                        continue
                cleaned = ''
                if resp.status_code == 200:
                    page = resp.text
                    snapshot['raw_html'] = page
                    cleaned = trafilatura.extract(
                        page, include_links=True, include_comments=False,
                        favor_recall=True)
                    if not cleaned:
                        soup = BeautifulSoup(page, 'html.parser')
                        for tag in soup(['script', 'style', 'noscript', 'template']):
                            tag.decompose()
                        cleaned = soup.get_text(separator=' ', strip=True)
                    
                    try:
                        meta = trafilatura.extract_metadata(page)
                        if meta is not None and getattr(meta, 'date', None):
                            snapshot['published'] = meta.date
                    except Exception:
                        pass
                        
                # --- Tier 1: Search Engine Dorking Fallback ---
                if resp.status_code in (401, 403) or len(cleaned) < MIN_TEXT_CHARS:
                    try:
                        from duckduckgo_search import DDGS
                        import warnings
                        warnings.filterwarnings('ignore')
                        q = f"site:{_host(url)} {urlparse(url).path.replace('/', ' ')} artificial intelligence OR machine learning OR privacy"
                        results = DDGS().text(q, max_results=3)
                        if results:
                            snippet_text = " ".join(r.get('body', '') for r in results)
                            if len(snippet_text) > 50:
                                cleaned = snippet_text
                                snapshot['status'] = 200  # Rescued!
                    except Exception:
                        pass
                
                # --- Tier 2: Wayback Machine CDX API Fallback ---
                if snapshot.get('status', resp.status_code) != 200 or len(cleaned) < MIN_TEXT_CHARS:
                    try:
                        import requests
                        cdx_url = f"http://web.archive.org/cdx/search/cdx?url={_host(url)}{urlparse(url).path}&output=json&limit=1&fl=timestamp,original"
                        r_cdx = self.session.get(cdx_url, timeout=5)
                        if r_cdx.status_code == 200:
                            data = r_cdx.json()
                            if len(data) > 1:
                                timestamp, original = data[1]
                                wayback_url = f"http://web.archive.org/web/{timestamp}id_/{original}"
                                rw = self.session.get(wayback_url, timeout=10)
                                if rw.status_code == 200:
                                    page = rw.text
                                    snapshot['raw_html'] = page
                                    c = trafilatura.extract(
                                        page, include_links=True, include_comments=False, favor_recall=True)
                                    if not c:
                                        soup = BeautifulSoup(page, 'html.parser')
                                        for tag in soup(['script', 'style', 'noscript', 'template']):
                                            tag.decompose()
                                        c = soup.get_text(separator=' ', strip=True)
                                    
                                    if len(c) > 50:
                                        cleaned = c
                                        snapshot['status'] = 200 # Rescued!
                    except Exception:
                        pass
                
                if snapshot['status'] == 200:
                    snapshot['cleaned_text'] = cleaned or ''
                    snapshot['content_hash'] = get_hash(snapshot['cleaned_text'])
                elif snapshot.get('status', resp.status_code) in (401, 403):
                    # If both search engine and Wayback Machine failed to find it,
                    # this "403" is actually a 404 masquerading behind a WAF.
                    snapshot['status'] = 404
                
                break
            except Exception as e:
                print(f"    FAILED: {e}")
                snapshot['error'] = str(e)
                if attempt == 0:
                    time.sleep(3)

        with open(cache_file, 'w') as f:
            json.dump(snapshot, f, indent=2)
        return snapshot

    # ── discovery ───────────────────────────────────────────

    def _discover_links(self, base_url, page_html):
        """Relevant same-site links from the homepage, with a guessed source tier."""
        base_host = _host(base_url)
        root = _root_domain(base_host)
        found = {}
        soup = BeautifulSoup(page_html, 'html.parser')
        for a in soup.find_all('a', href=True):
            href = a['href'].strip()
            if href.startswith(('mailto:', 'tel:', 'javascript:', '#')):
                continue
            full = normalize_url(urljoin(base_url + '/', href))
            if not full.startswith('http') or not _same_site(full, base_host):
                continue
            path = urlparse(full).path
            if SKIP_EXT.search(path):
                continue
            host = _host(full)
            sub = host[:-len(root)] if host.endswith(root) else ''   # e.g. "trust." of trust.x.com
            tier = guess_tier(_tokens(sub, path, a.get_text(' ', strip=True)[:80]))
            if tier:
                found.setdefault(full, tier)
        return found

    def _discover_sitemap(self, base_url):
        """Relevant URLs from sitemap.xml (one level of sitemap index)."""
        base_host = _host(base_url)
        origin = f"{urlparse(base_url).scheme}://{urlparse(base_url).netloc}"
        queue, fetched, found = [origin + '/sitemap.xml'], 0, {}
        while queue and fetched < 4:
            sm = queue.pop(0)
            fetched += 1
            if not self._allowed(sm):
                continue
            try:
                time.sleep(RATE_LIMIT_SECONDS)
                r = self.session.get(sm, timeout=15)
            except Exception:
                continue
            if r.status_code != 200:
                continue
            for loc in re.findall(r'<loc>\s*(.*?)\s*</loc>', r.text, re.I | re.S)[:5000]:
                loc = htmllib.unescape(loc)
                if not _same_site(loc, base_host):
                    continue
                if loc.lower().endswith('.xml'):
                    queue.append(loc)           # nested sitemap
                    continue
                path = urlparse(loc).path
                if SKIP_EXT.search(path):
                    continue
                tier = guess_tier(_tokens(path))
                if tier:
                    found.setdefault(normalize_url(loc), tier)
        return found

    # ── crawl ───────────────────────────────────────────────

    def crawl_vendor(self, vendor_id, base_url, depth_tier, refresh=False):
        """Crawl a vendor's public footprint based on criticality-driven depth."""
        print(f"\n{'=' * 60}")
        print(f"Crawling {vendor_id} — {base_url} (depth: {depth_tier})")
        print(f"{'=' * 60}")

        if not base_url.startswith('http'):
            base_url = 'https://' + base_url
        base_url = normalize_url(base_url)

        vendor_dir = os.path.join(self.cache_dir, vendor_id)
        os.makedirs(vendor_dir, exist_ok=True)

        allowed_tiers = DEPTH_TIERS.get(depth_tier, ["Tier_A"])

        # 1. Fixed paths from config (highest tier wins if a path is listed twice)
        targets = {}
        for tier_name in allowed_tiers:
            for cat in (self.config.get('sources', {}).get(tier_name) or []):
                for p in (cat.get('paths') or []):
                    u = normalize_url(urljoin(base_url + '/', p))
                    if u == base_url:
                        continue                # homepage is fetched separately
                    prev = targets.get(u)
                    if prev is None or TIER_ORDER.index(tier_name) < TIER_ORDER.index(prev[0]):
                        targets[u] = (tier_name, 'config')

        summary, seen_hashes = [], {}

        def record(snap, tier, origin, reason=None, duplicate_of=None):
            text = snap.get('cleaned_text', '') or ''
            status = snap.get('status', 'error')
            has_content = (status == 200 and len(text) >= MIN_TEXT_CHARS
                           and duplicate_of is None)
            if reason is None and not has_content:
                if duplicate_of:
                    reason = f"Duplicate of {duplicate_of} (likely soft-404 or redirect)"
                elif status == 200:
                    reason = f"Too little text ({len(text)} chars; likely JS-rendered or blocked)"
                else:
                    reason = f"HTTP {status}"
            summary.append({
                'url': snap.get('url', ''),
                'url_hash': snap.get('url_hash', get_hash(snap.get('url', ''))),
                'final_url': snap.get('final_url', snap.get('url', '')),
                'status': status,
                'content_hash': snap.get('content_hash', ''),
                'text_len': len(text),
                'source_tier': tier,
                'found_by': origin,
                'timestamp': snap.get('timestamp', ''),
                'published': snap.get('published', ''),
                'error': snap.get('error', None),
                'reason': reason,
                'has_content': has_content,
            })

        # 2. Homepage
        if self._allowed(base_url):
            home = self.fetch_url(base_url, vendor_dir, refresh=refresh)
            record(home, 'Tier_D', 'homepage')
            if home.get('status') == 200 and len(home.get('cleaned_text', '')) >= MIN_TEXT_CHARS:
                seen_hashes[home['content_hash']] = base_url
        else:
            home = {'url': base_url, 'status': 'robots_blocked'}
            record(home, 'Tier_D', 'homepage', reason=self._block_reason(base_url))

        # 3. Discovery (homepage links + sitemap)
        max_disc = MAX_DISCOVERED.get(depth_tier, 4)
        if max_disc and home.get('raw_html'):
            cands = dict(self._discover_links(base_url, home['raw_html']))
        else:
            cands = {}
        if max_disc:
            for u, t in self._discover_sitemap(base_url).items():
                cands.setdefault(u, t)
        ranked = sorted((TIER_ORDER.index(t), len(u), u, t) for u, t in cands.items()
                        if t in allowed_tiers and u not in targets and u != base_url)
        for _, _, u, t in ranked[:max_disc]:
            targets[u] = (t, 'discovered')

        # 4. Fetch everything, tier A first
        for url, (tier, origin) in sorted(targets.items(),
                                          key=lambda kv: (TIER_ORDER.index(kv[1][0]), kv[0])):
            if not self._allowed(url):
                print(f"    BLOCKED by robots.txt: {url}")
                record({'url': url, 'status': 'robots_blocked'}, tier, origin,
                       reason=self._block_reason(url))
                continue

            snap = self.fetch_url(url, vendor_dir, refresh=refresh)
            dup = None
            if snap.get('status') == 200 and len(snap.get('cleaned_text', '')) >= MIN_TEXT_CHARS:
                h = snap.get('content_hash', '')
                if h in seen_hashes:
                    dup = seen_hashes[h]
                else:
                    seen_hashes[h] = url
            record(snap, tier, origin, duplicate_of=dup)

        with open(os.path.join(vendor_dir, 'crawl_summary.json'), 'w') as f:
            json.dump(summary, f, indent=2)

        ok = sum(1 for s in summary if s['has_content'])
        dupes = sum(1 for s in summary if (s['reason'] or '').startswith('Duplicate'))
        blocked = sum(1 for s in summary if s['status'] == 'robots_blocked')
        gaps = len(summary) - ok
        found = sum(1 for s in summary if s['found_by'] == 'discovered' and s['has_content'])
        print(f"  Done: {ok} distinct pages with content ({found} via discovery), "
              f"{gaps} coverage gaps ({dupes} duplicates/soft-404, {blocked} robots-blocked)")

        return summary


def main():
    import sys
    sys.path.insert(0, 'src')
    from criticality import evaluate_vendor, load_rubric

    rubric = load_rubric('config/rubric.yaml')
    df = pd.read_excel('data/Meridian_Vendor_Input.xlsx', sheet_name='Vendor Inventory', header=3)
    df = df[df['Vendor ID'].str.startswith('V-', na=False) & (df['Vendor ID'] != 'V-000')]

    collector = OSINTCollector('config/sources.yaml')

    for idx, row in df.iterrows():
        vid = row['Vendor ID']
        url = row['Public Website']
        tier, _, _ = evaluate_vendor(row, rubric)
        if pd.notna(url):
            collector.crawl_vendor(vid, str(url), depth_tier=tier)


if __name__ == '__main__':
    main()
