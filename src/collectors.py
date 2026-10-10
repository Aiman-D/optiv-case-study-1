"""
M2: OSINT Collectors & Snapshot Store

Crawls vendor public websites based on depth policy from criticality tier.
Stores snapshots with URL, timestamp, content hash for offline replay.
Respects robots.txt and rate limits.
"""

import pandas as pd
import requests
import trafilatura
from bs4 import BeautifulSoup
from urllib.robotparser import RobotFileParser
import yaml
import os
import hashlib
import time
from urllib.parse import urljoin, urlparse
import json
from datetime import datetime, timezone
from pathlib import Path

def load_config(path):
    with open(path, 'r') as f:
        return yaml.safe_load(f)

def get_hash(text):
    return hashlib.sha256(text.encode('utf-8')).hexdigest()


# Depth policy: which source tiers to crawl per criticality tier
DEPTH_TIERS = {
    "Critical": ["Tier_A", "Tier_B", "Tier_C", "Tier_D"],
    "High":     ["Tier_A", "Tier_B"],
    "Medium":   ["Tier_B", "Tier_D"],
    "Low":      ["Tier_A"],
}


class OSINTCollector:
    def __init__(self, sources_config_path, cache_dir='cache'):
        self.config = load_config(sources_config_path)
        self.cache_dir = cache_dir
        os.makedirs(self.cache_dir, exist_ok=True)
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (compatible; MeridianOSINTBot/1.0; +http://example.com/bot)',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
        })
        self._robots_cache = {}

    def _check_robots(self, base_url, target_url):
        """Respect robots.txt — returns True if allowed."""
        if base_url not in self._robots_cache:
            rp = RobotFileParser()
            robots_url = urljoin(base_url, '/robots.txt')
            try:
                rp.set_url(robots_url)
                rp.read()
            except Exception:
                rp = None  # if we can't read it, assume allowed
            self._robots_cache[base_url] = rp

        rp = self._robots_cache[base_url]
        if rp is None:
            return True
        return rp.can_fetch('*', target_url)

    def fetch_url(self, url, vendor_dir, refresh=False):
        """Fetch a URL; cache result in vendor_dir. Returns snapshot dict."""
        url_hash = get_hash(url)
        cache_file = os.path.join(vendor_dir, f"{url_hash}.json")

        if os.path.exists(cache_file) and not refresh:
            with open(cache_file, 'r') as f:
                return json.load(f)

        print(f"  Fetching {url}...")
        snapshot = {
            'url': url,
            'url_hash': url_hash,
            'timestamp': datetime.now(timezone.utc).isoformat(),
            'status': 'error',
            'raw_html': '',
            'cleaned_text': '',
            'content_hash': ''
        }

        try:
            time.sleep(1.2)  # rate limit
            resp = self.session.get(url, timeout=15, allow_redirects=True)
            snapshot['status'] = resp.status_code
            if resp.status_code == 200:
                html = resp.text
                snapshot['raw_html'] = html
                cleaned = trafilatura.extract(html, include_links=True, include_comments=False)
                if not cleaned:
                    soup = BeautifulSoup(html, 'html.parser')
                    cleaned = soup.get_text(separator=' ', strip=True)
                snapshot['cleaned_text'] = cleaned or ''
                snapshot['content_hash'] = get_hash(snapshot['cleaned_text'])
        except Exception as e:
            print(f"    FAILED: {e}")
            snapshot['error'] = str(e)

        with open(cache_file, 'w') as f:
            json.dump(snapshot, f, indent=2)

        return snapshot

    def crawl_vendor(self, vendor_id, base_url, depth_tier, refresh=False):
        """Crawl a vendor's public footprint based on criticality-driven depth."""
        print(f"\n{'='*60}")
        print(f"Crawling {vendor_id} — {base_url} (depth: {depth_tier})")
        print(f"{'='*60}")

        if not base_url.startswith('http'):
            base_url = 'https://' + base_url

        vendor_dir = os.path.join(self.cache_dir, vendor_id)
        os.makedirs(vendor_dir, exist_ok=True)

        # Determine which source tiers to crawl
        allowed_tiers = DEPTH_TIERS.get(depth_tier, ["Tier_A"])

        paths_to_crawl = set()
        for tier_name in allowed_tiers:
            if tier_name in self.config.get('sources', {}):
                for category in self.config['sources'][tier_name]:
                    for p in category.get('paths', []):
                        paths_to_crawl.add(p)

        results = []

        # Always fetch homepage
        homepage = self.fetch_url(base_url, vendor_dir, refresh=refresh)
        homepage['source_tier'] = 'Tier_D'
        results.append(homepage)

        for path in sorted(paths_to_crawl):
            target_url = urljoin(base_url, path)
            if not self._check_robots(base_url, target_url):
                print(f"    BLOCKED by robots.txt: {target_url}")
                results.append({
                    'url': target_url,
                    'status': 'robots_blocked',
                    'content_hash': '',
                    'signal_type': 'COVERAGE-GAP',
                    'reason': 'Blocked by robots.txt'
                })
                continue

            snap = self.fetch_url(target_url, vendor_dir, refresh=refresh)
            # Determine source tier for this path
            for tier_name in allowed_tiers:
                if tier_name in self.config.get('sources', {}):
                    for cat in self.config['sources'][tier_name]:
                        if path in cat.get('paths', []):
                            snap['source_tier'] = tier_name
            results.append(snap)

        # Write crawl summary
        summary = []
        for r in results:
            summary.append({
                'url': r.get('url', ''),
                'url_hash': r.get('url_hash', get_hash(r.get('url', ''))),
                'status': r.get('status', 'error'),
                'content_hash': r.get('content_hash', ''),
                'source_tier': r.get('source_tier', 'unknown'),
                'timestamp': r.get('timestamp', ''),
                'error': r.get('error', None),
                'has_content': bool(r.get('cleaned_text', ''))
            })

        with open(os.path.join(vendor_dir, 'crawl_summary.json'), 'w') as f:
            json.dump(summary, f, indent=2)

        ok = sum(1 for s in summary if s['status'] == 200 and s['has_content'])
        gaps = sum(1 for s in summary if s['status'] != 200 or not s['has_content'])
        print(f"  Done: {ok} pages with content, {gaps} coverage gaps")

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
