"""
Build the static site published to GitHub Pages.

Scrapes every URL in urls.txt (or the SCRAPE_URLS env var, set by the manual
Action run), then writes into docs/:

    docs/index.html        overview of everything scraped
    docs/pages/<slug>.html readable view of one page
    docs/report.md         the SEO markdown report
    docs/scraped.docx      the Word export

Pages serves docs/ on main, so committing the output publishes it.

Usage:
    python build_site.py                 # URLs from urls.txt
    SCRAPE_URLS="https://a\nhttps://b" python build_site.py
"""
from __future__ import annotations

import concurrent.futures
import os
import re
import sys
from datetime import datetime, timezone
from urllib.parse import urlparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from url_scraper import (            # noqa: E402
    _READABLE_CSS,
    _clean_url_input,
    _esc,
    build_readable_html,
    build_seo_docx,
    build_seo_text_report,
    scrape_url,
)

ROOT = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(ROOT, "docs")
PAGES = os.path.join(DOCS, "pages")
URLS_FILE = os.path.join(ROOT, "urls.txt")

TIMEOUT = int(os.environ.get("SCRAPE_TIMEOUT", "30"))
MAX_WORKERS = int(os.environ.get("SCRAPE_WORKERS", "4"))


def read_url_list() -> list[str]:
    """URLs from the SCRAPE_URLS env var if set, else from urls.txt."""
    raw = os.environ.get("SCRAPE_URLS", "").strip()
    if not raw:
        if not os.path.exists(URLS_FILE):
            return []
        with open(URLS_FILE, encoding="utf-8") as f:
            raw = f.read()

    urls: list[str] = []
    seen: set[str] = set()
    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        cleaned = _clean_url_input(line)
        if cleaned and cleaned not in seen:
            seen.add(cleaned)
            urls.append(cleaned)
    return urls


def slugify(url: str) -> str:
    parsed = urlparse(url if "//" in url else "https://" + url)
    raw = (parsed.netloc + parsed.path).replace("www.", "")
    slug = re.sub(r"[^a-z0-9]+", "-", raw.lower()).strip("-")
    slug = re.sub(r"-(html?|php|aspx)$", "", slug)
    return (slug or "page")[:80]


def scrape_all(urls: list[str]) -> list[dict]:
    results: list[dict | None] = [None] * len(urls)

    def one(item):
        i, url = item
        print(f"[{i + 1}/{len(urls)}] {url}", flush=True)
        resolved, title, text, error, blocks, details = scrape_url(url, timeout=TIMEOUT)
        if error:
            print(f"    !! {error}", flush=True)
        else:
            print(f"    ok  {len(text.split())} words", flush=True)
        results[i] = {"url": resolved, "title": title, "text": text,
                      "error": error, "content_blocks": blocks, "details": details}

    workers = max(1, min(MAX_WORKERS, len(urls)))
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        list(pool.map(one, enumerate(urls)))
    return [r for r in results if r is not None]


def build_index(results: list[dict], slugs: list[str]) -> str:
    generated = datetime.now(timezone.utc).strftime("%d %b %Y, %H:%M UTC")
    ok = [r for r in results if not r.get("error")]
    failed = [r for r in results if r.get("error")]
    total_words = sum(len(r["text"].split()) for r in ok)

    rows = []
    for result, slug in zip(results, slugs):
        title = _esc(result.get("title") or "(No title)")
        url = _esc(result.get("url", ""))
        if result.get("error"):
            rows.append(
                '<tr class="failed"><td><span class="t">%s</span>'
                '<a class="u" href="%s">%s</a></td>'
                '<td colspan="2" class="err">%s</td></tr>'
                % (title, url, url, _esc(result["error"])))
            continue
        details = result.get("details") or {}
        rows.append(
            '<tr><td><a class="t" href="pages/%s.html">%s</a>'
            '<a class="u" href="%s">%s</a></td>'
            "<td>%s</td><td>%s</td></tr>"
            % (_esc(slug), title, url, url,
               f"{len(result['text'].split()):,}",
               details.get("total_images", 0)))

    cards = [("Pages", str(len(results))),
             ("Scraped", str(len(ok))),
             ("Words", f"{total_words:,}")]
    if failed:
        cards.append(("Failed", str(len(failed))))

    return (
        "<!DOCTYPE html>\n"
        '<html lang="en"><head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
        "<title>URL Scraper</title>\n<style>%s\n"
        ".idx{max-width:960px;margin:0 auto;padding:28px 20px 80px}\n"
        ".cards{display:flex;gap:12px;flex-wrap:wrap;margin:18px 0 24px}\n"
        ".card{background:var(--card);border:1px solid var(--line);border-radius:10px;"
        "padding:14px 20px;min-width:120px}\n"
        ".card .k{font-size:.72rem;text-transform:uppercase;letter-spacing:.06em;"
        "color:var(--muted)}\n.card .v{font-size:1.5rem;font-weight:700;margin-top:2px}\n"
        "table.list{width:100%%;border-collapse:collapse;background:var(--card);"
        "border:1px solid var(--line);border-radius:10px;overflow:hidden}\n"
        "table.list th{background:var(--mark);text-align:left;padding:10px 14px;"
        "font-size:.72rem;text-transform:uppercase;letter-spacing:.06em;color:var(--muted)}\n"
        "table.list td{padding:12px 14px;border-top:1px solid var(--line);"
        "vertical-align:top;font-size:.88rem}\n"
        "table.list td:nth-child(2),table.list td:nth-child(3),"
        "table.list th:nth-child(2),table.list th:nth-child(3)"
        "{text-align:right;white-space:nowrap;width:90px}\n"
        ".t{display:block;font-weight:600;margin-bottom:3px}\n"
        ".u{display:block;font-size:.76rem;color:var(--muted);word-break:break-all}\n"
        ".err{color:var(--amber)}\ntr.failed .t{color:var(--muted)}\n"
        ".dl{display:flex;gap:10px;flex-wrap:wrap;margin:22px 0 0}\n"
        ".dl a{display:inline-block;padding:8px 16px;border-radius:7px;"
        "background:var(--card);border:1px solid var(--line);text-decoration:none;"
        "font-size:.85rem;font-weight:600}\n</style>\n</head>\n"
        '<body id="top"><div class="topbar"><div class="inner">'
        "<div><h1>URL Scraper</h1>"
        '<div class="sub">%d page%s &middot; rebuilt %s</div></div>'
        "</div></div>\n"
        '<div class="idx">\n<div class="cards">%s</div>\n'
        '<table class="list"><thead><tr><th>Page</th><th>Words</th><th>Images</th>'
        "</tr></thead><tbody>%s</tbody></table>\n"
        '<div class="dl"><a href="report.md">SEO report (.md)</a>'
        '<a href="scraped.docx">Word export (.docx)</a></div>\n'
        '<p class="foot">Built by the Scrape and publish Action from urls.txt</p>\n'
        "</div></body></html>\n"
        % (_READABLE_CSS, len(results), "" if len(results) == 1 else "s",
           _esc(generated),
           "".join('<div class="card"><div class="k">%s</div>'
                   '<div class="v">%s</div></div>' % (k, v) for k, v in cards),
           "".join(rows))
    )


def main() -> int:
    urls = read_url_list()
    if not urls:
        print("No URLs found in urls.txt or SCRAPE_URLS.", file=sys.stderr)
        return 1

    print(f"Scraping {len(urls)} URL(s) with {MAX_WORKERS} workers\n", flush=True)
    results = scrape_all(urls)

    os.makedirs(PAGES, exist_ok=True)
    # Pages would otherwise run the output through Jekyll.
    open(os.path.join(DOCS, ".nojekyll"), "w").close()

    # Drop readable pages from previous runs so removed URLs don't linger.
    for stale in os.listdir(PAGES):
        if stale.endswith(".html"):
            os.remove(os.path.join(PAGES, stale))

    slugs, used = [], set()
    for result in results:
        slug = slugify(result.get("url", ""))
        candidate, n = slug, 2
        while candidate in used:
            candidate, n = f"{slug}-{n}", n + 1
        used.add(candidate)
        slugs.append(candidate)

        if result.get("error"):
            continue
        page = build_readable_html([result], doc_title=result.get("title") or candidate)
        with open(os.path.join(PAGES, candidate + ".html"), "w", encoding="utf-8") as f:
            f.write(page)

    with open(os.path.join(DOCS, "index.html"), "w", encoding="utf-8") as f:
        f.write(build_index(results, slugs))

    with open(os.path.join(DOCS, "report.md"), "w", encoding="utf-8") as f:
        f.write(build_seo_text_report(results))

    try:
        build_seo_docx(results, os.path.join(DOCS, "scraped.docx"))
    except Exception as e:                      # never fail the build on the Word export
        print(f"Word export skipped: {e}", file=sys.stderr)

    ok = sum(1 for r in results if not r.get("error"))
    words = sum(len(r["text"].split()) for r in results if not r.get("error"))
    print(f"\nBuilt docs/ — {ok}/{len(results)} scraped, {words:,} words")
    return 0


if __name__ == "__main__":
    sys.exit(main())
