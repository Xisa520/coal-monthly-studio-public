"""GitHub Actions public NBS article candidate checker.
This script publishes ONLY information extracted from public stats.gov.cn pages.
Never place internal workbooks, user uploads, API keys or review data here.
"""
import datetime
import json
from pathlib import Path
from nbs import discover, read_energy

OUT = Path(__file__).resolve().parents[1] / 'data' / 'nbs-public.json'

def main():
    latest = discover()
    if not latest:
        raise SystemExit('No official energy-production article found. Existing file left untouched.')
    results = []
    for article in latest[:3]:
        try:
            r = read_energy(article)
            results.append({
                'title': r['title'],
                'url': r['url'],
                'month': r['month'],
                'candidates': [{
                    'indicator': c['indicator'],
                    'value': c['value'],
                    'unit': c['unit'],
                    'year_on_year': c['year_on_year'],
                    'excerpt': c['raw_excerpt']
                } for c in r.get('candidates', [])]
            })
        except Exception as e:
            print(f'Warning: {article["month"]} not read: {e}')
    if not results:
        raise SystemExit('No readable articles. Existing file left untouched.')
    payload = {
        'source': '国家统计局',
        'source_index': 'https://www.stats.gov.cn/sj/zxfb/',
        'generated_at': datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%d %H:%M UTC'),
        'notice': '自动识别的候选文章和候选数据，并非人工审核通过的统计结果；请对照原文核实。',
        'articles': results
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(f'Wrote {OUT} with {len(results)} public articles.')

if __name__ == '__main__':
    main()
