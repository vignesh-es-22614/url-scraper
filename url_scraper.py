"""
URL Scraper Tool
----------------
Reads URLs from an Excel (.xlsx) or CSV file, scrapes text content
from each URL, and saves results into a formatted Word document.

Usage:
    python url_scraper.py urls.xlsx             # Excel input
    python url_scraper.py urls.csv              # CSV input
    python url_scraper.py urls.xlsx -o out.docx # Custom output name

Excel/CSV format:
    - Must have a column named 'URL' or 'url' (or the first column is used)
    - One URL per row
"""
from __future__ import annotations

import sys
import os
import time
import argparse
import re
import html as _html
import requests
from bs4 import BeautifulSoup
from bs4.element import NavigableString, Tag
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openpyxl
import csv
from datetime import datetime
from urllib.parse import urljoin, urlparse


# ── Config ──────────────────────────────────────────────────────────────────

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}
REQUEST_TIMEOUT = 15       # seconds (default)
DELAY_BETWEEN_REQUESTS = 1 # seconds — be polite
MAX_TEXT_LENGTH = 0      # characters per URL (0 = unlimited, default)
MAX_HTML_EXPORT_CHARS = 30000

_INVALID_XML_RE = re.compile(
    r"[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]"
)


def _sanitize_xml_text(value: str | None) -> str:
    if not value:
        return ""
    return _INVALID_XML_RE.sub("", value)


# Pasted URLs often arrive wrapped in a label, list marker or punctuation:
# "page: https://x", "- https://x", "1. https://x", "<https://x>", "https://x,".
_URL_LABEL_RE = re.compile(
    r"^(?:[-*•>]+\s*|\d+[.)]\s*)*(?:(?:page|url|link|site|address)\s*[:=]\s*)?",
    re.IGNORECASE)
_EMBEDDED_URL_RE = re.compile(r"https?://\S+", re.IGNORECASE)
_URL_TRIM_CHARS = "<>\"'`()[] \t"


def _clean_url_input(raw: str) -> str:
    """Pull the actual URL out of a pasted line."""
    value = _sanitize_xml_text(raw).strip().strip(_URL_TRIM_CHARS)
    if not value:
        return ""

    embedded = _EMBEDDED_URL_RE.search(value)
    if embedded:
        value = embedded.group(0)
    else:
        value = _URL_LABEL_RE.sub("", value, count=1).strip()
        value = value.split()[0] if value.split() else ""

    return value.strip(_URL_TRIM_CHARS).rstrip(",;.")


def _url_candidates(url: str) -> list[str]:
    cleaned = _clean_url_input(url)
    if not cleaned:
        return []
    if cleaned.startswith(("http://", "https://")):
        return [cleaned]
    # Bare host: must at least look like a domain before we prepend a scheme,
    # otherwise "http://" + junk produces an unhelpful parse error downstream.
    host = cleaned.split("/")[0]
    if " " in cleaned or "." not in host or host.startswith(".") or host.endswith("."):
        return []
    return [f"https://{cleaned}", f"http://{cleaned}"]


def _add_hyperlink(paragraph, text: str, url: str):
    text = _sanitize_xml_text(text)
    url = _sanitize_xml_text(url).strip()

    if not text:
        return None
    if not url:
        paragraph.add_run(text)
        return None

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), paragraph.part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True))

    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    r_pr.append(color)

    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(underline)

    new_run.append(r_pr)

    text_element = OxmlElement("w:t")
    text_element.text = text
    new_run.append(text_element)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


def _resolve_url(base_url: str, value: str) -> str:
    value = _sanitize_xml_text(value).strip()
    if not value:
        return ""
    return _sanitize_xml_text(urljoin(base_url, value)).strip()


# Lazy-loading markup keeps the real image in a data-* attribute and leaves a
# tiny placeholder in src, so prefer the data-* sources before falling back.
_IMG_SRC_ATTRS = ("data-src", "data-lazy-src", "data-original", "data-image", "src")


def _best_img_src(img: Tag, base_url: str) -> str:
    """Pick the highest-fidelity source for an <img>, resolved to an absolute URL."""
    for attr in _IMG_SRC_ATTRS:
        value = str(img.get(attr) or "").strip()
        if value and not value.startswith("data:"):
            return _resolve_url(base_url, value)

    for attr in ("data-srcset", "srcset"):
        raw = str(img.get(attr) or "").strip()
        if raw:
            candidates = [c.strip().split()[0] for c in raw.split(",") if c.strip()]
            if candidates:
                return _resolve_url(base_url, candidates[-1])

    value = str(img.get("src") or "").strip()
    return _resolve_url(base_url, value) if value else ""


def _is_internal_link(base_url: str, href: str) -> bool:
    base_host = urlparse(base_url).netloc.lower()
    href_host = urlparse(href).netloc.lower()
    return bool(base_host) and href_host == base_host


# ── Sitemap discovery ───────────────────────────────────────────────────────

SITEMAP_MAX_URLS = 5000     # cap on URLs returned to the UI
SITEMAP_MAX_FILES = 50      # cap on sitemap files read per request
SITEMAP_COMMON_PATHS = ("/sitemap.xml", "/sitemap_index.xml",
                        "/sitemap-index.xml", "/sitemap/sitemap.xml")


def _fetch_sitemap_bytes(url: str, timeout: int) -> bytes:
    resp = requests.get(url, headers=HEADERS, timeout=timeout)
    resp.raise_for_status()
    data = resp.content
    # .xml.gz sitemaps, and servers that gzip without a Content-Encoding header.
    if data[:2] == b"\x1f\x8b":
        import gzip
        data = gzip.decompress(data)
    return data


def _parse_sitemap(data: bytes, base_url: str) -> tuple[list[str], list[str]]:
    """Return (child sitemap URLs, page URLs) from one sitemap document."""
    text = data.decode("utf-8", "ignore").strip()

    # Plain-text sitemaps: one URL per line.
    if not text.startswith("<") and "<urlset" not in text and "<sitemapindex" not in text:
        pages = [ln.strip() for ln in text.splitlines()
                 if ln.strip().lower().startswith(("http://", "https://"))]
        return [], pages

    try:
        soup = BeautifulSoup(data, "xml")
    except Exception:
        soup = BeautifulSoup(data, "html.parser")

    children: list[str] = []
    pages: list[str] = []

    for node in soup.find_all("sitemap"):
        loc = node.find("loc")
        if loc and loc.get_text(strip=True):
            children.append(_resolve_url(base_url, loc.get_text(strip=True)))

    for node in soup.find_all("url"):
        loc = node.find("loc")
        if loc and loc.get_text(strip=True):
            pages.append(_resolve_url(base_url, loc.get_text(strip=True)))

    # Some sitemaps use <loc> without the wrapper elements above.
    if not children and not pages:
        for loc in soup.find_all("loc"):
            value = loc.get_text(strip=True)
            if value:
                pages.append(_resolve_url(base_url, value))

    return children, pages


def _sitemap_candidates(source: str) -> list[str]:
    """Turn user input into sitemap URLs to try, in order."""
    cleaned = _clean_url_input(source)
    if not cleaned:
        return []
    if not cleaned.startswith(("http://", "https://")):
        cleaned = "https://" + cleaned

    parsed = urlparse(cleaned)
    if "." not in parsed.netloc or parsed.netloc.startswith(".") or parsed.netloc.endswith("."):
        return []
    path = (parsed.path or "").lower()
    # Already points at a sitemap - use it as given.
    if path.endswith((".xml", ".xml.gz", ".txt")) or "sitemap" in path:
        return [cleaned]

    root = f"{parsed.scheme}://{parsed.netloc}"
    return [root + suffix for suffix in SITEMAP_COMMON_PATHS]


def _sitemaps_from_robots(source: str, timeout: int) -> list[str]:
    cleaned = _clean_url_input(source)
    if not cleaned:
        return []
    if not cleaned.startswith(("http://", "https://")):
        cleaned = "https://" + cleaned
    parsed = urlparse(cleaned)
    try:
        resp = requests.get(f"{parsed.scheme}://{parsed.netloc}/robots.txt",
                            headers=HEADERS, timeout=timeout)
        resp.raise_for_status()
    except Exception:
        return []
    found = []
    for line in resp.text.splitlines():
        if line.lower().startswith("sitemap:"):
            value = line.split(":", 1)[1].strip()
            if value:
                found.append(value)
    return found


def fetch_sitemap_urls(source: str, timeout: int = 20,
                       max_urls: int = SITEMAP_MAX_URLS) -> dict:
    """Collect page URLs from a sitemap, sitemap index, or a bare domain.

    Accepts a direct sitemap URL, or a domain - in which case robots.txt and
    the usual sitemap paths are tried. Nested sitemap indexes are followed.
    """
    candidates = _sitemap_candidates(source)
    if not candidates:
        return {"urls": [], "sitemaps": [], "truncated": False,
                "error": f"Not a valid URL or domain: {source.strip()!r}"}

    # robots.txt is authoritative when the user gave a bare domain.
    if len(candidates) > 1:
        candidates = _sitemaps_from_robots(source, timeout) + candidates

    queue = list(dict.fromkeys(candidates))
    seen_files: set[str] = set()
    read_files: list[str] = []
    urls: list[str] = []
    seen_urls: set[str] = set()
    last_error = ""
    truncated = False

    while queue and len(read_files) < SITEMAP_MAX_FILES:
        current = queue.pop(0)
        if current in seen_files:
            continue
        seen_files.add(current)

        try:
            data = _fetch_sitemap_bytes(current, timeout)
        except Exception as e:
            last_error = f"{current}: {e}"
            continue

        children, pages = _parse_sitemap(data, current)
        if not children and not pages:
            continue

        read_files.append(current)
        for child in children:
            if child not in seen_files:
                queue.append(child)
        for page in pages:
            if page in seen_urls:
                continue
            seen_urls.add(page)
            urls.append(page)
            if len(urls) >= max_urls:
                truncated = True
                queue = []
                break

    error = ""
    if not urls:
        error = (f"No sitemap found. Last error - {last_error}" if last_error
                 else "No URLs found in the sitemap.")

    return {"urls": urls, "sitemaps": read_files,
            "truncated": truncated, "error": error}


# ── Template / chrome stripping ─────────────────────────────────────────────

_TMPL_TAGS = ["script", "style", "nav", "footer", "header",
              "aside", "noscript", "iframe", "svg"]

# Unmistakable chrome — removed even when it contains a heading.
_CHROME_STRONG = ["cookie", "popup", "modal", "overlay", "advertisement",
                  "breadcrumb", "social", "share", "sticky-", "floating"]

# Usually chrome, but marketing pages routinely wrap the <h1> in a "banner" or
# "promo" block, so these are only removed when they hold no heading.
_CHROME_WEAK = ["sidebar", "side-bar", "widget", "banner", "promo",
                "related", "recommended", "ads", "toc", "table-of-content"]

_CHROME_IDS_STRONG = ["cookie", "popup", "modal", "breadcrumb"]
_CHROME_IDS_WEAK = ["sidebar", "nav", "menu", "navigation", "banner",
                    "toc", "header", "footer"]

_CONTENT_ROOT_HINTS = ("content", "main", "main-content", "page-content")

# Tags that yield a block of their own, so a container holding any of them must
# be recursed into rather than flattened into a single text block.
_BLOCK_LEVEL = [
    "address", "article", "aside", "blockquote", "button", "dd", "details",
    "div", "dl", "dt", "fieldset", "figcaption", "figure", "footer", "form",
    "h1", "h2", "h3", "h4", "h5", "h6", "header", "hr", "img", "li", "main",
    "nav", "ol", "p", "picture", "pre", "section", "summary", "table",
    "tbody", "td", "tfoot", "th", "thead", "tr", "ul", "video",
]


def _is_text_node(node) -> bool:
    """True for real text only.

    Comment/CData/Doctype all subclass NavigableString, so an isinstance check
    would pull HTML comments such as <!-- END --> into the extracted text.
    """
    return type(node) is NavigableString


def _attr_matches(value: str, pattern: str) -> bool:
    """Token-aware attribute match.

    Plain substring matching was too blunt: 'ads' matched 'downloads' and
    'content' matched 'tab-content-container'. Single-word patterns must match
    a whole token; hyphenated patterns still match as substrings.
    """
    value = (value or "").lower()
    if "-" in pattern:
        return pattern in value
    return pattern in re.split(r"[^a-z0-9]+", value)


def _el_attr_text(el: Tag, attr: str) -> str:
    value = el.get(attr)
    if isinstance(value, list):
        return " ".join(value)
    return str(value or "")


def _strip_template_chrome(soup: BeautifulSoup) -> None:
    """Remove navigation/cookie/ad furniture, keeping anything that carries a heading."""
    for t in soup(_TMPL_TAGS):
        t.decompose()

    for el in soup.find_all(True):
        try:
            if el.decomposed or el.name in ("html", "body"):
                continue
            cls = _el_attr_text(el, "class")
            eid = _el_attr_text(el, "id")
        except AttributeError:      # parent was decomposed mid-iteration
            continue

        strong = (any(_attr_matches(cls, p) for p in _CHROME_STRONG)
                  or any(_attr_matches(eid, p) for p in _CHROME_IDS_STRONG))
        weak = (any(_attr_matches(cls, p) for p in _CHROME_WEAK)
                or any(_attr_matches(eid, p) for p in _CHROME_IDS_WEAK))

        if not strong and not weak:
            continue
        if weak and not strong and el.find(["h1", "h2"]):
            continue                # hero/banner block that holds the headline
        el.decompose()


def _pick_content_root(soup: BeautifulSoup) -> Tag:
    """Pick the element holding the page's real content.

    Candidates are scored by how much of the body text they actually contain;
    a narrow widget such as <div class="tab-content-container"> no longer wins
    just because its class happens to contain the word 'content'.
    """
    body = soup.find("body") or soup
    body_len = len(body.get_text(" ", strip=True))
    if not body_len:
        return body

    candidates: list[Tag] = []
    candidates.extend(soup.find_all(["main", "article"]))
    candidates.extend(soup.find_all(attrs={"role": "main"}))
    for attr in ("id", "class"):
        for el in soup.find_all(attrs={attr: True}):
            value = _el_attr_text(el, attr)
            if any(_attr_matches(value, p) for p in _CONTENT_ROOT_HINTS):
                candidates.append(el)

    best, best_len = None, 0
    for el in candidates:
        length = len(el.get_text(" ", strip=True))
        if length > best_len:
            best, best_len = el, length

    # Only trust a candidate that holds the bulk of the page; otherwise the
    # whole body is safer than silently dropping most of the content.
    if best is not None and best_len >= body_len * 0.6:
        return best
    return body


def _walk_content(elem: Tag, base_url: str) -> list[dict[str, object]]:
    """
    Walk DOM in document order and return structured blocks.
    Captures headings, paragraphs, lists, inline images, tables, code, CTA links.
    Links rendered as [text] (url) inline in text.
    """
    blocks: list[dict[str, object]] = []

    _SKIP = {"script", "style", "noscript", "meta", "link",
             "head", "template", "canvas", "select", "textarea", "iframe"}

    def _inline(node: Tag) -> str:
        parts: list[str] = []
        for child in node.children:
            if _is_text_node(child):
                txt = " ".join(_sanitize_xml_text(str(child)).split())
                if txt:
                    parts.append(txt)
            elif isinstance(child, Tag):
                cn = (child.name or "").lower()
                if cn in _SKIP:
                    continue
                if cn == "a":
                    lt = " ".join(_sanitize_xml_text(
                        child.get_text(" ", strip=True)).split())
                    href = _resolve_url(base_url, child.get("href", ""))
                    if lt and href:
                        parts.append(f"[{lt}] ({href})")
                    elif lt:
                        parts.append(lt)
                elif cn == "img":
                    pass  # captured as block
                elif cn in {"code", "kbd", "samp", "var", "tt"}:
                    ct = _sanitize_xml_text(child.get_text("", strip=True))
                    if ct:
                        parts.append(f"`{ct}`")
                else:
                    parts.append(_inline(child))
        return " ".join(parts).strip()

    def _table_rows(tbl: Tag) -> list[list[str]] | None:
        rows: list[list[str]] = []
        for tr in tbl.find_all("tr"):
            cells = [_inline(c) for c in tr.find_all(["th", "td"])]
            if any(c.strip() for c in cells):
                rows.append(cells)
        return rows if rows else None

    def walk(node: Tag) -> None:
        if not isinstance(node, Tag):
            return
        tag = (node.name or "").lower()
        if not tag or tag in _SKIP:
            return

        if tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            text = _inline(node)
            if text:
                blocks.append({"tag": tag, "type": "heading", "text": text})
            return

        if tag == "p":
            text = _inline(node)
            if text and len(text) > 3:
                blocks.append({"tag": "p", "type": "paragraph", "text": text})
            for img in node.find_all("img", recursive=True):
                _add_img_block(img)
            return

        if tag == "li":
            text = _inline(node)
            if text:
                blocks.append({"tag": "li", "type": "list_item", "text": text})
            return

        if tag == "img":
            _add_img_block(node)
            return

        if tag == "pre":
            code_node = node.find("code")
            raw = _sanitize_xml_text((code_node or node).get_text("", strip=False))
            if raw.strip():
                lang = ""
                for cls in (code_node or node).get("class", []):
                    if "language-" in cls or "lang-" in cls:
                        lang = cls.split("-", 1)[-1]; break
                blocks.append({"tag": "pre", "type": "code",
                               "text": raw, "language": lang})
            return

        if tag == "code" and not node.find_parent("pre"):
            raw = _sanitize_xml_text(node.get_text("", strip=True))
            if raw.strip():
                blocks.append({"tag": "code", "type": "code",
                               "text": raw, "language": ""})
            return

        if tag == "table":
            rows = _table_rows(node)
            if rows:
                blocks.append({"tag": "table", "type": "table", "rows": rows})
            return

        if tag == "a":
            cls = " ".join(node.get("class", [])).lower()
            if any(k in cls for k in ["btn", "cta", "button"]):
                lt = " ".join(_sanitize_xml_text(
                    node.get_text(" ", strip=True)).split())
                href = _resolve_url(base_url, node.get("href", ""))
                if lt and href and not lt.lower().startswith("skip"):
                    blocks.append({"tag": "cta", "type": "cta",
                                   "text": lt, "href": href})
            return

        # FAQ accordion: <details><summary>Q</summary>…answer…</details>
        if tag == "details":
            summary = node.find("summary")
            if summary:
                q = _inline(summary)
                if q:
                    blocks.append({"tag": "summary", "type": "faq_question",
                                   "text": q})
            # walk remaining children (the answer body)
            for child in node.children:
                if isinstance(child, Tag) and child.name != "summary":
                    walk(child)
            return

        if tag == "summary":
            text = _inline(node)
            if text:
                blocks.append({"tag": "summary", "type": "faq_question",
                               "text": text})
            return

        # Definition lists: <dl><dt>term</dt><dd>definition</dd></dl>
        if tag == "dt":
            text = _inline(node)
            if text:
                blocks.append({"tag": "dt", "type": "faq_question",
                               "text": text})
            return

        if tag == "dd":
            text = _inline(node)
            if text:
                blocks.append({"tag": "dd", "type": "paragraph", "text": text})
            return

        if tag == "blockquote":
            text = _inline(node)
            if text:
                blocks.append({"tag": "blockquote", "type": "paragraph",
                               "text": f"> {text}"})
            return

        # Detect FAQ questions by class/id patterns on any element.
        # Common patterns: <button class="accordion-button">, <div class="faq-question">,
        # <span class="question">, <li class="faq-item"> with child button, etc.
        _FAQ_Q_CLS = {
            "question", "faq-q", "faq-question", "faq_question",
            "accordion-button", "accordion-title", "accordion-header",
            "collapse-title", "collapse-button",
            "panel-title", "card-header",
        }
        node_cls  = " ".join(node.get("class", [])).lower()
        node_id   = (node.get("id") or "").lower()

        # Direct class/id match on current element
        if any(p in node_cls or p in node_id for p in _FAQ_Q_CLS):
            text = _inline(node)
            if text:
                blocks.append({"tag": tag, "type": "faq_question", "text": text})
            # Still recurse for answer children if it's a container (not a button/span)
            if tag not in {"button", "span", "a", "strong", "em", "b", "i"}:
                for child in node.children:
                    if isinstance(child, Tag):
                        walk(child)
            return

        # <button> anywhere inside an accordion/faq container
        if tag == "button":
            parent = node.parent
            if isinstance(parent, Tag):
                p_cls = " ".join(parent.get("class", [])).lower()
                p_id  = (parent.get("id") or "").lower()
                _FAQ_CONTAINERS = {"faq", "accordion", "collapse", "panel"}
                if any(k in p_cls or k in p_id for k in _FAQ_CONTAINERS):
                    text = _inline(node)
                    if text and not text.lower().startswith("skip"):
                        blocks.append({"tag": "button", "type": "faq_question",
                                       "text": text})
                    return
            # Non-FAQ button — skip
            return

        # Generic container. When it holds no block-level descendant it *is* a
        # text block — e.g. <div class="tab-button"><span>Label</span></div>,
        # whose text was previously dropped because it sits in no <p>/<li>/<h*>.
        if node.find(_BLOCK_LEVEL) is None:
            text = _inline(node)
            if text and len(text) > 1:
                blocks.append({"tag": tag, "type": "paragraph", "text": text})
            return

        for child in node.children:
            if _is_text_node(child):
                # Loose text sitting in a wrapper alongside block children.
                txt = " ".join(_sanitize_xml_text(str(child)).split())
                if len(txt) > 1 and re.search(r"\w", txt):
                    blocks.append({"tag": tag, "type": "paragraph", "text": txt})
            elif isinstance(child, Tag):
                walk(child)

    def _add_img_block(img: Tag) -> None:
        src = _best_img_src(img, base_url)
        alt = " ".join(_sanitize_xml_text(img.get("alt", "")).split())
        if not src and not alt:
            return
        blocks.append({
            "tag": "img", "type": "image",
            "src": src, "alt": alt or "N/A",
            "title": " ".join(_sanitize_xml_text(img.get("title", "")).split()),
            "width": _sanitize_xml_text(str(img.get("width", ""))).strip(),
            "height": _sanitize_xml_text(str(img.get("height", ""))).strip(),
        })

    for child in elem.children:
        if isinstance(child, Tag):
            walk(child)

    return blocks


# Keep old name as alias so CLI code still works
_extract_content_blocks = _walk_content


def _collect_images(soup: BeautifulSoup, base_url: str) -> list[dict[str, str]]:
    images: list[dict[str, str]] = []
    seen: set[tuple[str, str, str]] = set()

    for img in soup.find_all("img"):
        src = _best_img_src(img, base_url)
        alt = " ".join(_sanitize_xml_text(img.get("alt", "")).split())
        title = " ".join(_sanitize_xml_text(img.get("title", "")).split())
        if not src and not alt and not title:
            continue
        key = (src, alt, title)
        if key in seen:
            continue
        seen.add(key)
        images.append({
            "src": src,
            "alt": alt,
            "title": title,
            "width": _sanitize_xml_text(str(img.get("width", ""))).strip(),
            "height": _sanitize_xml_text(str(img.get("height", ""))).strip(),
        })

    return images


def _collect_metadata(soup: BeautifulSoup, resolved_url: str) -> dict[str, str | list[dict[str, str]]]:
    metadata: dict[str, str | list[dict[str, str]]] = {
        "url": _sanitize_xml_text(resolved_url),
        "language": _sanitize_xml_text((soup.html or {}).get("lang", "") if soup.html else ""),
    }

    title_tag = soup.find("title")
    metadata["title"] = _sanitize_xml_text(title_tag.get_text(strip=True) if title_tag else "")

    canonical = soup.find("link", rel=lambda value: value and "canonical" in str(value).lower())
    metadata["canonical"] = _sanitize_xml_text(canonical.get("href", "") if canonical else "")

    meta_pairs: list[dict[str, str]] = []
    for m in soup.find_all("meta"):
        key = _sanitize_xml_text(m.get("name") or m.get("property") or m.get("http-equiv") or "").strip()
        value = _sanitize_xml_text(m.get("content") or "").strip()
        if key and value:
            meta_pairs.append({"key": key, "value": value})

    metadata["meta_tags"] = meta_pairs

    def _first_meta(*keys: str) -> str:
        key_set = {k.lower() for k in keys}
        for item in meta_pairs:
            if item["key"].lower() in key_set:
                return item["value"]
        return ""

    metadata["description"] = _first_meta("description", "og:description", "twitter:description")
    metadata["keywords"] = _first_meta("keywords")
    metadata["author"] = _first_meta("author")
    return metadata


def _collect_headings(soup: BeautifulSoup) -> dict[str, list[str]]:
    headings: dict[str, list[str]] = {}
    for level in ["h1", "h2", "h3", "h4", "h5", "h6"]:
        vals: list[str] = []
        for h in soup.find_all(level):
            txt = _sanitize_xml_text(h.get_text(" ", strip=True))
            if txt:
                vals.append(txt)
        headings[level] = vals
    return headings


def _collect_links(soup: BeautifulSoup, base_url: str) -> list[dict[str, object]]:
    links: list[dict[str, object]] = []
    seen: set[tuple[str, str]] = set()

    for a in soup.find_all("a"):
        href = _sanitize_xml_text(a.get("href", "")).strip()
        if not href:
            continue
        absolute = _sanitize_xml_text(urljoin(base_url, href)).strip()
        if not absolute:
            continue
        text = _sanitize_xml_text(a.get_text(" ", strip=True))
        key = (absolute, text)
        if key in seen:
            continue
        seen.add(key)
        links.append({
            "href": absolute,
            "text": text,
            "internal": _is_internal_link(base_url, absolute),
            "external": not _is_internal_link(base_url, absolute),
        })

    return links


def _collect_breadcrumb(soup: BeautifulSoup, base_url: str) -> list[dict[str, str]]:
    breadcrumb_nodes = []
    breadcrumb_nodes.extend(soup.select("nav[aria-label*='breadcrumb' i] a"))
    breadcrumb_nodes.extend(soup.select("[class*='breadcrumb' i] a"))
    breadcrumb_nodes.extend(soup.select("[id*='breadcrumb' i] a"))

    crumbs: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for node in breadcrumb_nodes:
        text = _sanitize_xml_text(node.get_text(" ", strip=True))
        href = _resolve_url(base_url, node.get("href", ""))
        if not text and not href:
            continue
        key = (text, href)
        if key in seen:
            continue
        seen.add(key)
        crumbs.append({"text": text, "href": href})
    return crumbs


def _collect_cta_buttons(soup: BeautifulSoup, base_url: str) -> list[dict[str, str]]:
    nodes = soup.find_all(["a", "button"])
    ctas: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    keywords = {
        "sign up", "signup", "download", "start", "try", "trial", "contact",
        "book demo", "get started", "learn more", "request", "buy now"
    }

    for node in nodes:
        text = _sanitize_xml_text(node.get_text(" ", strip=True))
        href = ""
        if node.name == "a":
            href = _resolve_url(base_url, node.get("href", ""))

        cls = _sanitize_xml_text(" ".join(node.get("class", []))).lower()
        text_l = text.lower()
        is_cta = any(k in text_l for k in keywords) or any(k in cls for k in ["btn", "cta", "button", "download"])
        if not is_cta:
            continue
        if len(text) > 90 or len(text) < 2:
            continue
        if text_l.startswith("skip ") or text_l == "skip to content":
            continue
        if href.startswith(base_url + "#"):
            continue

        key = (text, href)
        if key in seen:
            continue
        seen.add(key)
        ctas.append({"text": text, "href": href})

    return ctas[:20]


def _image_variant_summary(src: str) -> str:
    src_l = src.lower()
    if src_l.endswith(".webp"):
        return "WebP"
    if src_l.endswith(".gif"):
        return "GIF"
    if src_l.endswith(".png"):
        return "PNG"
    if src_l.endswith(".jpg") or src_l.endswith(".jpeg"):
        return "JPG"
    return "Image"


def _format_url_section(index: int, result: dict) -> str:
    safe_url = _sanitize_xml_text(str(result.get("url", "")))
    safe_title = _sanitize_xml_text(str(result.get("title", "")))
    safe_error = _sanitize_xml_text(str(result.get("error", "")))
    details = result.get("details") if isinstance(result.get("details"), dict) else {}

    lines: list[str] = []
    lines.append(f"**URL {index}: {safe_url}**")
    lines.append("")

    if safe_error:
        lines.append(f"Could not scrape: {safe_error}")
        lines.append("")
        return "\n".join(lines)

    lines.append(f"*Page Title: {safe_title}*")
    lines.append("")
    lines.append("**Metadata**")
    lines.append("")

    metadata = details.get("metadata") if isinstance(details, dict) else {}
    metadata = metadata if isinstance(metadata, dict) else {}
    fields = [
        ("Final URL", details.get("final_url", safe_url)),
        ("Domain", details.get("domain", "")),
        ("Status", details.get("status_code", "")),
        ("Language", metadata.get("language", "")),
        ("Canonical", metadata.get("canonical", "")),
        ("Description", metadata.get("description", "")),
        ("Keywords", metadata.get("keywords", "")),
        ("Author", metadata.get("author", "")),
    ]
    for key, value in fields:
        clean = _sanitize_xml_text(str(value or "")).strip()
        lines.append(f"**{key}:** {clean if clean else 'N/A'}")

    lines.append("")
    lines.append("**Meta Tags**")
    meta_tags = metadata.get("meta_tags") if isinstance(metadata, dict) else []
    if isinstance(meta_tags, list) and meta_tags:
        for item in meta_tags[:80]:
            if not isinstance(item, dict):
                continue
            key = _sanitize_xml_text(str(item.get("key", ""))).strip()
            val = _sanitize_xml_text(str(item.get("value", ""))).strip()
            if key and val:
                lines.append(f"- {key}: {val}")
    else:
        lines.append("- N/A")

    breadcrumb = details.get("breadcrumb") if isinstance(details, dict) else []
    breadcrumb_labels: list[str] = []
    if isinstance(breadcrumb, list) and breadcrumb:
        crumbs = []
        for c in breadcrumb:
            if not isinstance(c, dict):
                continue
            txt = _sanitize_xml_text(str(c.get("text", ""))).strip() or "Link"
            href = _sanitize_xml_text(str(c.get("href", ""))).strip()
            breadcrumb_labels.append(txt.lower())
            crumbs.append(f"[{txt}]({href})" if href else txt)
        if crumbs:
            lines.append("")
            lines.append("**Breadcrumb:** " + " > ".join(crumbs))

    lines.append("")
    lines.append("---")
    lines.append("")

    blocks = details.get("content_blocks") if isinstance(details, dict) else []
    seen_img_srcs: set[str] = set()
    if isinstance(blocks, list):
        for block in blocks:
            if not isinstance(block, dict):
                continue
            btype = block.get("type", "")
            tag   = _sanitize_xml_text(str(block.get("tag", "")).lower())

            if btype == "heading":
                level = int(tag[1]) if len(tag) == 2 and tag[1].isdigit() else 2
                text  = _sanitize_xml_text(str(block.get("text", ""))).strip()
                if text:
                    lines.append(f"[H{level}] {text}")
                    lines.append("")

            elif btype == "faq_question":
                text = _sanitize_xml_text(str(block.get("text", ""))).strip()
                if text:
                    lines.append(f"**Q: {text}**")
                    lines.append("")

            elif btype == "paragraph":
                text = _sanitize_xml_text(str(block.get("text", ""))).strip()
                if text:
                    lines.append(text)
                    lines.append("")

            elif btype == "list_item":
                text = _sanitize_xml_text(str(block.get("text", ""))).strip()
                if text:
                    lines.append(f"- {text}")

            elif btype == "image":
                src    = _sanitize_xml_text(str(block.get("src", ""))).strip()
                alt    = _sanitize_xml_text(str(block.get("alt", ""))).strip() or "N/A"
                title  = _sanitize_xml_text(str(block.get("title", ""))).strip()
                width  = _sanitize_xml_text(str(block.get("width", ""))).strip()
                height = _sanitize_xml_text(str(block.get("height", ""))).strip()
                if src and src in seen_img_srcs:
                    continue
                if src:
                    seen_img_srcs.add(src)
                dims = f" ({width}x{height})" if width and height else ""
                lines.append("**Image:**")
                lines.append(f"alt text: {alt}")
                if title:
                    lines.append(f"title: {title}")
                if src:
                    lines.append(f"- {_image_variant_summary(src)}: {src}{dims}")
                lines.append("")

            elif btype == "table":
                rows = block.get("rows", [])
                if rows:
                    max_cols = max(len(r) for r in rows)
                    norm = [r + [""] * (max_cols - len(r)) for r in rows]
                    # sanitize pipes in cell values
                    def _cell(v: str) -> str:
                        return _sanitize_xml_text(str(v)).replace("|", "\\|").strip()
                    lines.append("| " + " | ".join(_cell(c) for c in norm[0]) + " |")
                    lines.append("| " + " | ".join(["---"] * max_cols) + " |")
                    for row in norm[1:]:
                        lines.append("| " + " | ".join(_cell(c) for c in row) + " |")
                    lines.append("")

            elif btype == "code":
                lang = _sanitize_xml_text(str(block.get("language", ""))).strip()
                code = str(block.get("text", "")).rstrip()
                if code.strip():
                    lines.append(f"```{lang}")
                    lines.append(code)
                    lines.append("```")
                    lines.append("")

            elif btype == "cta":
                txt  = _sanitize_xml_text(str(block.get("text", ""))).strip()
                href = _sanitize_xml_text(str(block.get("href", ""))).strip()
                if txt and href:
                    lines.append(f"[CTA: {txt}] ({href})")
                    lines.append("")

    return "\n".join(lines).strip() + "\n"


def build_seo_text_report(results: list[dict]) -> str:
    sections = []
    for idx, result in enumerate(results, 1):
        sections.append(_format_url_section(idx, result))
    return "\n\n".join(section.strip() for section in sections if section.strip()) + "\n"


def build_seo_docx(results: list[dict], output_path: str) -> None:
    """Write the SEO-format report into a styled Word document."""
    doc = Document()

    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.2)
        section.right_margin  = Inches(1.2)

    # Cover
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cr = cp.add_run("Web Content SEO Report")
    cr.bold = True; cr.font.size = Pt(22)
    cr.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    sp = doc.add_paragraph()
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sp.add_run(
        f"Generated: {datetime.now().strftime('%B %d, %Y %H:%M')}  ·  "
        f"{len(results)} URL(s) processed"
    )
    sr.font.size = Pt(10); sr.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    doc.add_paragraph()

    def _bold_kv(para, key: str, value: str):
        para.add_run(f"{key}: ").bold = True
        para.add_run(value)

    def _add_hyperlink_run(para, text: str, url: str):
        """Inline clickable hyperlink."""
        try:
            hyperlink = OxmlElement("w:hyperlink")
            hyperlink.set(qn("r:id"),
                para.part.relate_to(
                    url,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                    is_external=True))
            r = OxmlElement("w:r")
            rpr = OxmlElement("w:rPr")
            col = OxmlElement("w:color"); col.set(qn("w:val"), "0563C1"); rpr.append(col)
            ul  = OxmlElement("w:u");    ul.set(qn("w:val"), "single");  rpr.append(ul)
            r.append(rpr)
            te = OxmlElement("w:t"); te.text = _sanitize_xml_text(text); r.append(te)
            hyperlink.append(r)
            para._p.append(hyperlink)
        except Exception:
            para.add_run(_sanitize_xml_text(text))

    for idx, result in enumerate(results, 1):
        safe_url   = _sanitize_xml_text(str(result.get("url", "")))
        safe_title = _sanitize_xml_text(str(result.get("title", "")))
        safe_error = _sanitize_xml_text(str(result.get("error", "")))
        details    = result.get("details") or {}

        # ── URL heading ──────────────────────────────────────────────
        p = doc.add_paragraph()
        r = p.add_run(f"URL {idx}: {safe_url}")
        r.bold = True; r.font.size = Pt(13)
        r.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D); r.underline = True

        if safe_error:
            ep = doc.add_paragraph()
            ep.add_run(f"Could not scrape: {safe_error}").font.color.rgb = RGBColor(0xCC, 0, 0)
            doc.add_paragraph("─" * 80); doc.add_paragraph(); continue

        # Page title
        tp = doc.add_paragraph()
        tr2 = tp.add_run(f"Page Title: {safe_title}")
        tr2.italic = True; tr2.font.size = Pt(10)
        tr2.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        # ── Metadata ──────────────────────────────────────────────────
        doc.add_paragraph().add_run("Metadata").bold = True
        metadata = details.get("metadata") or {}
        for key, value in [
            ("Final URL",   details.get("final_url", safe_url)),
            ("Domain",      details.get("domain", "")),
            ("Status",      details.get("status_code", "")),
            ("Language",    metadata.get("language", "")),
            ("Canonical",   metadata.get("canonical", "")),
            ("Description", metadata.get("description", "")),
            ("Keywords",    metadata.get("keywords", "")),
            ("Author",      metadata.get("author", "")),
        ]:
            clean = _sanitize_xml_text(str(value or "")).strip()
            if not clean:
                continue
            _bold_kv(doc.add_paragraph(), key, clean)

        meta_tags = metadata.get("meta_tags") or []
        if meta_tags:
            doc.add_paragraph().add_run("Meta Tags").bold = True
            for item in meta_tags[:80]:
                mk = _sanitize_xml_text(str(item.get("key", ""))).strip()
                mv = _sanitize_xml_text(str(item.get("value", ""))).strip()
                if mk and mv:
                    doc.add_paragraph(f"- {mk}: {mv}")

        # ── Breadcrumb ────────────────────────────────────────────────
        breadcrumb = details.get("breadcrumb") or []
        if breadcrumb:
            bp = doc.add_paragraph()
            bp.add_run("Breadcrumb: ").bold = True
            for ci, c in enumerate(breadcrumb):
                if ci:
                    bp.add_run(" > ")
                bt = _sanitize_xml_text(str(c.get("text", "")))
                bh = _sanitize_xml_text(str(c.get("href", "")))
                if bh:
                    _add_hyperlink_run(bp, bt, bh)
                else:
                    bp.add_run(bt)

        doc.add_paragraph("─" * 60)

        # ── Content blocks in document order ─────────────────────────
        seen_img_srcs: set[str] = set()
        blocks = details.get("content_blocks") or []
        for block in blocks:
            if not isinstance(block, dict):
                continue
            btype = block.get("type", "")
            text  = _sanitize_xml_text(str(block.get("text", ""))).strip()

            if btype == "heading":
                tag   = block.get("tag", "h2")
                level = int(tag[1]) if len(tag) == 2 and tag[1].isdigit() else 2
                if text:
                    hp = doc.add_paragraph()
                    hr = hp.add_run(f"[H{level}] {text}")
                    hr.bold = True
                    hr.font.size = Pt(max(9, 14 - level))
                    hr.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

            elif btype in {"paragraph", "faq_question"}:
                if text:
                    pp = doc.add_paragraph()
                    if btype == "faq_question":
                        pp.add_run("Q: ").bold = True
                    pp.add_run(text).font.size = Pt(11)

            elif btype == "list_item":
                if text:
                    li_p = doc.add_paragraph(style="List Bullet")
                    li_p.add_run(text).font.size = Pt(11)

            elif btype == "image":
                src    = _sanitize_xml_text(str(block.get("src", ""))).strip()
                alt    = _sanitize_xml_text(str(block.get("alt", ""))).strip() or "N/A"
                title  = _sanitize_xml_text(str(block.get("title", ""))).strip()
                width  = _sanitize_xml_text(str(block.get("width", ""))).strip()
                height = _sanitize_xml_text(str(block.get("height", ""))).strip()
                if src and src in seen_img_srcs:
                    continue
                if src:
                    seen_img_srcs.add(src)
                dims = f" ({width}x{height})" if width and height else ""
                ip = doc.add_paragraph()
                ip.add_run("Image – ").bold = True
                ip.add_run(f"alt: {alt}")
                if title:
                    ip.add_run(f" | title: {title}")
                if src:
                    doc.add_paragraph(f"{_image_variant_summary(src)}: {src}{dims}")

            elif btype == "table":
                rows = block.get("rows", [])
                if not rows:
                    continue
                max_cols = max(len(r) for r in rows)
                norm = [r + [""] * (max_cols - len(r)) for r in rows]
                tbl = doc.add_table(rows=len(norm), cols=max_cols)
                tbl.style = "Table Grid"
                for ri, row in enumerate(norm):
                    for ci, cell in enumerate(row):
                        tbl.cell(ri, ci).text = _sanitize_xml_text(str(cell))
                        if ri == 0:
                            for run in tbl.cell(0, ci).paragraphs[0].runs:
                                run.bold = True
                doc.add_paragraph()

            elif btype == "code":
                lang = _sanitize_xml_text(str(block.get("language", ""))).strip()
                code = _sanitize_xml_text(str(block.get("text", ""))).rstrip()
                if code.strip():
                    lbl = doc.add_paragraph()
                    lbl.add_run(f"Code{(' (' + lang + ')') if lang else ''}:").bold = True
                    cp2 = doc.add_paragraph(code)
                    for run in cp2.runs:
                        run.font.name = "Courier New"
                        run.font.size = Pt(9)

            elif btype == "cta":
                lt   = _sanitize_xml_text(str(block.get("text", ""))).strip()
                href = _sanitize_xml_text(str(block.get("href", ""))).strip()
                if lt and href:
                    ctap = doc.add_paragraph()
                    ctap.add_run("CTA: ").bold = True
                    _add_hyperlink_run(ctap, lt, href)

        doc.add_paragraph("─" * 80)
        doc.add_paragraph()

    doc.save(output_path)


# ── Input readers ────────────────────────────────────────────────────────────

def read_urls_from_excel(path: str) -> list[str]:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    headers = [str(c.value).strip().lower() if c.value else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
    
    # Find the URL column
    url_col = None
    for i, h in enumerate(headers):
        if h in ("url", "urls", "link", "links", "website"):
            url_col = i
            break
    if url_col is None:
        url_col = 0  # fallback: first column
        start_row = 1
    else:
        start_row = 2  # skip header row

    urls = []
    for row in ws.iter_rows(min_row=start_row, values_only=True):
        val = row[url_col] if len(row) > url_col else None
        if val:
            url = str(val).strip()
            if url:
                urls.append(url)
    return urls


def read_urls_from_csv(path: str) -> list[str]:
    urls = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        headers = next(reader, [])
        url_col = None
        for i, h in enumerate(headers):
            if h.strip().lower() in ("url", "urls", "link", "links", "website"):
                url_col = i
                break
        if url_col is None:
            url_col = 0
            # If first row had no header matching, treat it as a URL
            if headers:
                urls.append(headers[0].strip())

        for row in reader:
            val = row[url_col].strip() if len(row) > url_col else ""
            if val:
                urls.append(val)
    return urls


def read_urls(path: str) -> list[str]:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return read_urls_from_excel(path)
    elif ext == ".csv":
        return read_urls_from_csv(path)
    else:
        # Try plain text — one URL per line
        with open(path, encoding="utf-8") as f:
            return [line.strip() for line in f if line.strip()]


# ── Scraper ──────────────────────────────────────────────────────────────────

def scrape_url(url: str, timeout: int = None, max_text: int = None) -> tuple[str, str, str, str, list[dict[str, object]], dict]:
    """
    Returns (resolved_url, title, text, error_message, content_blocks, details).
    error_message is empty string on success.
    timeout and max_text override the module-level defaults when provided.
    """
    _timeout  = timeout  if timeout  is not None else REQUEST_TIMEOUT
    _max_text = max_text if max_text is not None else MAX_TEXT_LENGTH
    candidates = _url_candidates(url)
    if not candidates:
        return url, "", "", f"Not a valid URL: {url.strip()!r}", [], {}
    last_error = ""

    for candidate in candidates:
        try:
            resp = requests.get(candidate, headers=HEADERS, timeout=_timeout)
            resp.raise_for_status()

            # Parse raw bytes, not resp.text: requests falls back to ISO-8859-1
            # for text/html without a charset, which mangles UTF-8 punctuation
            # into mojibake. BeautifulSoup sniffs <meta charset> / BOM instead.
            declared_charset = None
            ctype = resp.headers.get("Content-Type", "")
            if "charset=" in ctype.lower():
                declared_charset = ctype.lower().split("charset=", 1)[1].split(";")[0].strip() or None

            soup = BeautifulSoup(resp.content, "html.parser", from_encoding=declared_charset)
            content_soup = BeautifulSoup(resp.content, "html.parser", from_encoding=declared_charset)

            title_tag = soup.find("title")
            title = _sanitize_xml_text(title_tag.get_text(strip=True) if title_tag else "(No title)")

            metadata = _collect_metadata(soup, resp.url)
            headings = _collect_headings(soup)
            links = _collect_links(soup, resp.url)
            images = _collect_images(soup, resp.url)
            breadcrumb = _collect_breadcrumb(soup, resp.url)
            cta_buttons = _collect_cta_buttons(soup, resp.url)

            _strip_template_chrome(content_soup)
            main = _pick_content_root(content_soup)

            blocks = _walk_content(main, resp.url)

            text = " ".join(
                _sanitize_xml_text(str(b.get("text", ""))).strip()
                for b in blocks
                if b.get("type") in {"heading", "paragraph", "list_item"}
            )
            if not text.strip():
                text = "(No readable text content found on this page.)"
            if _max_text and len(text) > _max_text:
                text = text[:_max_text] + "\n\n[... content truncated ...]"

            parsed_page = urlparse(resp.url)
            details = {
                "final_url": _sanitize_xml_text(resp.url),
                "status_code": str(resp.status_code),
                "domain": _sanitize_xml_text(parsed_page.netloc),
                "metadata": metadata,
                "headings": headings,
                "links": links,
                "images": images,
                "breadcrumb": breadcrumb,
                "cta_buttons": cta_buttons,
                "content_blocks": blocks,
                "total_links": len(links),
                "internal_links": sum(1 for l in links if l.get("internal")),
                "external_links": sum(1 for l in links if l.get("external")),
                "total_images": len(images),
                "full_text": text,
            }

            return resp.url, title, text, "", blocks, details

        except requests.exceptions.Timeout:
            last_error = f"Request timed out after {_timeout}s"
        except requests.exceptions.ConnectionError as e:
            last_error = f"Connection error: {e}"
        except requests.exceptions.HTTPError as e:
            last_error = f"HTTP {e.response.status_code}: {e}"
        except Exception as e:
            last_error = f"Unexpected error: {e}"

    return (candidates[0] if candidates else url), "", "", last_error or "Unexpected error", [], {}


# ── Word doc builder ─────────────────────────────────────────────────────────

def build_docx(results: list[dict], output_path: str) -> None:
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.2)
        section.right_margin  = Inches(1.2)

    # ── Cover title ────────────────────────────────────────────────────────
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.add_run("Web Content Scrape Report")
    run.bold = True
    run.font.size = Pt(22)
    run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    sub_para = doc.add_paragraph()
    sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_para.add_run(
        f"Generated: {datetime.now().strftime('%B %d, %Y %H:%M')}  ·  "
        f"{len(results)} URL(s) processed"
    )
    sub_run.font.size = Pt(10)
    sub_run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    doc.add_paragraph()  # spacer

    # ── Each URL ──────────────────────────────────────────────────────────
    for i, r in enumerate(results, 1):
        safe_url = _sanitize_xml_text(r.get("url", ""))
        safe_title = _sanitize_xml_text(r.get("title", ""))
        safe_error = _sanitize_xml_text(r.get("error", ""))
        details = r.get("details") or {}

        # URL heading
        url_heading = doc.add_paragraph()
        idx_run = url_heading.add_run(f"URL {i}: ")
        idx_run.bold = True
        idx_run.font.size = Pt(13)
        idx_run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

        url_run = url_heading.add_run(safe_url)
        url_run.bold = True
        url_run.font.size = Pt(13)
        url_run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        url_run.underline = True

        # Page title (if available)
        if safe_title:
            tp = doc.add_paragraph()
            t = tp.add_run(f"Page Title: {safe_title}")
            t.italic = True
            t.font.size = Pt(10)
            t.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        # Error or content
        if safe_error:
            ep = doc.add_paragraph()
            er = ep.add_run(f"⚠ Could not scrape: {safe_error}")
            er.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
            er.font.size = Pt(11)
        else:
            metadata = details.get("metadata") if isinstance(details, dict) else None
            if isinstance(metadata, dict):
                meta_heading = doc.add_paragraph()
                meta_run = meta_heading.add_run("Metadata")
                meta_run.bold = True
                meta_run.font.size = Pt(11)

                summary_fields = [
                    ("Final URL", details.get("final_url", safe_url)),
                    ("Domain", details.get("domain", "")),
                    ("Status", details.get("status_code", "")),
                    ("Language", metadata.get("language", "")),
                    ("Canonical", metadata.get("canonical", "")),
                    ("Description", metadata.get("description", "")),
                    ("Keywords", metadata.get("keywords", "")),
                    ("Author", metadata.get("author", "")),
                ]
                for key, value in summary_fields:
                    clean = _sanitize_xml_text(str(value or "")).strip()
                    if not clean:
                        continue
                    p = doc.add_paragraph()
                    p.add_run(f"{key}: ").bold = True
                    p.add_run(clean)

                meta_tags = metadata.get("meta_tags") or []
                if isinstance(meta_tags, list) and meta_tags:
                    doc.add_paragraph().add_run("Meta Tags").bold = True
                    for item in meta_tags[:80]:
                        if not isinstance(item, dict):
                            continue
                        m_key = _sanitize_xml_text(str(item.get("key", ""))).strip()
                        m_val = _sanitize_xml_text(str(item.get("value", ""))).strip()
                        if not m_key or not m_val:
                            continue
                        doc.add_paragraph(f"- {m_key}: {m_val}")

            images = details.get("images") if isinstance(details, dict) else None
            if isinstance(images, list) and images:
                doc.add_paragraph().add_run(f"Images ({len(images)})").bold = True
                for image in images[:150]:
                    if not isinstance(image, dict):
                        continue
                    src = _sanitize_xml_text(str(image.get("src", ""))).strip()
                    alt = _sanitize_xml_text(str(image.get("alt", ""))).strip()
                    title_text = _sanitize_xml_text(str(image.get("title", ""))).strip()
                    if not src and not alt and not title_text:
                        continue
                    details_line = []
                    if alt:
                        details_line.append(f"alt: {alt}")
                    if title_text:
                        details_line.append(f"title: {title_text}")
                    if src:
                        details_line.append(f"src: {src}")
                    doc.add_paragraph("- " + " | ".join(details_line))

            headings = details.get("headings") if isinstance(details, dict) else None
            if isinstance(headings, dict):
                doc.add_paragraph().add_run("Headings").bold = True
                for level in ["h1", "h2", "h3", "h4", "h5", "h6"]:
                    vals = headings.get(level) or []
                    if not vals:
                        continue
                    doc.add_paragraph().add_run(f"{level.upper()} ({len(vals)})").bold = True
                    for txt in vals[:25]:
                        clean_txt = _sanitize_xml_text(str(txt)).strip()
                        if clean_txt:
                            doc.add_paragraph(f"- {clean_txt}")

            links = details.get("links") if isinstance(details, dict) else None
            if isinstance(links, list):
                doc.add_paragraph().add_run(f"Links ({len(links)})").bold = True
                for link in links[:150]:
                    if not isinstance(link, dict):
                        continue
                    href = _sanitize_xml_text(str(link.get("href", ""))).strip()
                    text = _sanitize_xml_text(str(link.get("text", "")).strip() or href)
                    if not href:
                        continue
                    lp = doc.add_paragraph()
                    _add_hyperlink(lp, text, href)

            doc.add_paragraph().add_run("Readable Page Content").bold = True
            # Content paragraphs
            blocks = r.get("content_blocks") or []
            if blocks:
                for block in blocks:
                    if not isinstance(block, dict):
                        continue
                    cp = doc.add_paragraph()
                    if block.get("tag") in {"h1", "h2", "h3", "h4", "h5", "h6"}:
                        prefix = f"{str(block.get('tag')).upper()}: "
                        header_run = cp.add_run(prefix)
                        header_run.bold = True
                    for segment in block.get("segments", []):
                        if not isinstance(segment, dict):
                            continue
                        segment_type = segment.get("type")
                        text = _sanitize_xml_text(str(segment.get("text", ""))).strip()
                        href = _sanitize_xml_text(str(segment.get("href", ""))).strip() or None
                        if not text:
                            if segment_type == "image":
                                src = _sanitize_xml_text(str(segment.get("src", ""))).strip()
                                alt = _sanitize_xml_text(str(segment.get("alt", ""))).strip()
                                image_text = alt or src
                                if image_text:
                                    cp.add_run(f"[Image: {image_text}" + (f" | {src}" if src else "") + "]")
                            continue
                        if segment_type == "link" and href:
                            _add_hyperlink(cp, text, href)
                        elif segment_type == "image":
                            src = _sanitize_xml_text(str(segment.get("src", ""))).strip()
                            alt = _sanitize_xml_text(str(segment.get("alt", ""))).strip()
                            image_text = alt or text or src
                            cp.add_run(f"[Image: {image_text}" + (f" | {src}" if src else "") + "]")
                        else:
                            cp.add_run(text).font.size = Pt(11)
            else:
                for line in _sanitize_xml_text(r.get("text", "")).split("\n\n"):
                    line = line.strip()
                    if not line:
                        continue
                    cp = doc.add_paragraph()
                    cr = cp.add_run(line)
                    cr.font.size = Pt(11)

            raw_html = _sanitize_xml_text(str(details.get("raw_html", ""))) if isinstance(details, dict) else ""
            if raw_html:
                doc.add_paragraph().add_run("Raw HTML (truncated)").bold = True
                html_excerpt = raw_html[:MAX_HTML_EXPORT_CHARS]
                if len(raw_html) > MAX_HTML_EXPORT_CHARS:
                    html_excerpt += "\n\n[... raw HTML truncated ...]"
                doc.add_paragraph(html_excerpt)

        # Divider
        doc.add_paragraph("─" * 80)
        doc.add_paragraph()

    doc.save(output_path)


# ── Main ──────────────────────────────────────────────────────────────────────

# ── Readable HTML builder ────────────────────────────────────────────────────
#
# Renders the same structured blocks the .md / .docx exports use, but as a
# self-contained HTML page meant to be *read* in a browser: real headings,
# real links, real images, real tables — no external CSS/JS/fonts.

# Inline markers produced by _walk_content: "[label] (https://…)" and `code`.
_MD_LINK_RE = re.compile(r"\[([^\[\]]+?)\]\s*\((https?://[^\s()]+)\)")
_MD_CODE_RE = re.compile(r"`([^`\n]+)`")

_READABLE_CSS = """
*,*::before,*::after{box-sizing:border-box}
:root{
  --bg:#f6f7f9; --card:#fff; --ink:#1a1d21; --muted:#6b7280; --line:#e3e6ea;
  --blue:#1b4f8a; --blue-lt:#2f6fb8; --green:#166534; --amber:#92400e;
  --amber-bg:#fef3c7; --code-bg:#f2f4f7; --mark:#eef4fb;
}
@media (prefers-color-scheme:dark){
  :root{
    --bg:#14171a; --card:#1c2025; --ink:#e7e9ec; --muted:#9aa3ad; --line:#2c323a;
    --blue:#7fb2e8; --blue-lt:#9cc6f2; --green:#6ee7a8; --amber:#fbbf24;
    --amber-bg:#3b2f12; --code-bg:#22272e; --mark:#1f2732;
  }
}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--ink);
  font:16px/1.7 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif}
a{color:var(--blue-lt)}
a:hover{color:var(--blue)}
.wrap{max-width:860px;margin:0 auto;padding:0 20px 80px}

/* top bar */
.topbar{position:sticky;top:0;z-index:10;background:var(--card);
  border-bottom:1px solid var(--line);padding:12px 20px}
.topbar .inner{max-width:860px;margin:0 auto;display:flex;align-items:center;
  gap:14px;flex-wrap:wrap}
.topbar h1{margin:0;font-size:1rem;font-weight:700}
.topbar .sub{color:var(--muted);font-size:.78rem}
.tools{margin-left:auto;display:flex;gap:14px;flex-wrap:wrap}
.tools label{font-size:.78rem;color:var(--muted);cursor:pointer;user-select:none;
  display:inline-flex;align-items:center;gap:5px}

/* table of contents */
.toc{background:var(--card);border:1px solid var(--line);border-radius:10px;
  padding:16px 20px;margin:22px 0}
.toc h2{margin:0 0 8px;font-size:.8rem;text-transform:uppercase;
  letter-spacing:.06em;color:var(--muted)}
.toc ol{margin:0;padding-left:20px}
.toc li{margin:4px 0;font-size:.9rem}

/* page section */
.page{background:var(--card);border:1px solid var(--line);border-radius:10px;
  margin:22px 0;overflow:hidden}
.pg-head{padding:20px 24px;border-bottom:1px solid var(--line);background:var(--mark)}
.pg-num{font-size:.72rem;font-weight:700;letter-spacing:.08em;
  text-transform:uppercase;color:var(--muted)}
.pg-head h2{margin:4px 0 6px;font-size:1.3rem;line-height:1.35}
.src{font-size:.83rem;word-break:break-all}
.chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
.chip{font-size:.72rem;padding:3px 9px;border-radius:20px;
  background:var(--card);border:1px solid var(--line);color:var(--muted)}
.crumb{margin-top:10px;font-size:.78rem;color:var(--muted)}
.desc{margin:10px 0 0;font-size:.88rem;color:var(--muted)}

/* metadata drawer */
.meta{border-bottom:1px solid var(--line)}
.meta>summary{cursor:pointer;padding:11px 24px;font-size:.82rem;
  font-weight:600;color:var(--muted)}
.meta>summary:hover{color:var(--ink)}
.meta .body{padding:4px 24px 18px}
.meta table{width:100%;border-collapse:collapse;font-size:.8rem}
.meta td{padding:5px 8px;border-bottom:1px solid var(--line);vertical-align:top}
.meta td:first-child{width:180px;color:var(--muted);font-weight:600}

/* article content */
.content{padding:10px 24px 28px}
.content h1,.content h2,.content h3,
.content h4,.content h5,.content h6{line-height:1.3;margin:1.5em 0 .5em}
.content h1{font-size:1.6rem}
.content h2{font-size:1.32rem}
.content h3{font-size:1.12rem}
.content h4,.content h5,.content h6{font-size:1rem}
.content p{margin:0 0 1em}
.content ul{margin:0 0 1em;padding-left:24px}
.content li{margin:.3em 0}
.content blockquote{margin:0 0 1em;padding:2px 0 2px 16px;
  border-left:3px solid var(--line);color:var(--muted)}
.content code{background:var(--code-bg);padding:1px 5px;border-radius:4px;
  font-size:.88em;font-family:ui-monospace,SFMono-Regular,Consolas,monospace}
.content pre{background:var(--code-bg);border:1px solid var(--line);
  border-radius:8px;padding:14px 16px;overflow-x:auto;margin:0 0 1em}
.content pre code{background:none;padding:0;font-size:.84rem;line-height:1.55}
.hlvl{display:none;font-size:.6em;font-weight:700;vertical-align:middle;
  margin-right:8px;padding:2px 6px;border-radius:4px;
  background:var(--mark);border:1px solid var(--line);color:var(--muted)}
body.show-levels .hlvl{display:inline-block}
.faq-q{font-weight:700;margin:1.2em 0 .4em!important}
.faq-q::before{content:"Q: ";color:var(--muted)}
.cta{display:inline-block;margin:0 8px 1em 0;padding:8px 16px;border-radius:7px;
  background:var(--blue);color:#fff!important;font-size:.85rem;font-weight:600;
  text-decoration:none}
.cta:hover{background:var(--blue-lt);color:#fff!important}

/* tables */
.tbl-wrap{overflow-x:auto;margin:0 0 1.2em}
.content table{border-collapse:collapse;width:100%;font-size:.86rem}
.content th,.content td{border:1px solid var(--line);padding:7px 10px;
  text-align:left;vertical-align:top}
.content th{background:var(--mark);font-weight:700}

/* images */
figure{margin:0 0 1.4em}
figure img{max-width:100%;height:auto;display:block;border-radius:8px;
  border:1px solid var(--line);background:var(--mark)}
figcaption{font-size:.76rem;color:var(--muted);margin-top:6px;word-break:break-word}
figcaption .altline{color:var(--ink)}
body.no-images figure img{display:none}
.img-broken{font-size:.78rem;color:var(--amber);background:var(--amber-bg);
  border:1px solid var(--line);border-radius:8px;padding:10px 12px}

/* errors + footer */
.err{margin:0;padding:20px 24px;color:var(--amber);background:var(--amber-bg);
  font-size:.9rem}
.totop{display:block;padding:10px 24px;font-size:.75rem;color:var(--muted);
  border-top:1px solid var(--line);text-decoration:none}
.totop:hover{color:var(--blue)}
.foot{text-align:center;color:var(--muted);font-size:.75rem;margin-top:26px}

@media print{
  .topbar,.tools,.totop,.toc{display:none}
  body{background:#fff}
  .page{border:none;margin:0 0 24px;page-break-inside:avoid}
  .meta[open]>summary{display:none}
}
"""

_READABLE_JS = """
(function(){
  var body = document.body;
  // invert=true -> the class is applied when the box is UNchecked.
  function bind(id, cls, invert){
    var el = document.getElementById(id);
    if(!el) return;
    function apply(){ body.classList.toggle(cls, invert ? !el.checked : el.checked); }
    apply();
    el.addEventListener('change', apply);
  }
  bind('t-levels', 'show-levels', false);
  bind('t-images', 'no-images', true);

  // Flag images that fail to load so the caption still carries the alt text.
  Array.prototype.forEach.call(document.images, function(img){
    img.addEventListener('error', function(){
      img.style.display = 'none';
      var cap = img.parentNode.querySelector('figcaption');
      if (cap) cap.classList.add('img-broken');
    });
  });
})();
"""


def _esc(value) -> str:
    """Sanitize then HTML-escape any scraped value."""
    return _html.escape(_sanitize_xml_text(str(value if value is not None else "")))


def _rich(text: str) -> str:
    """Escape text, then turn the inline `[label] (url)` and `code` markers
    produced by _walk_content back into real <a> and <code> elements."""
    out = _esc(text)
    out = _MD_LINK_RE.sub(
        lambda m: '<a href="%s" target="_blank" rel="noopener noreferrer">%s</a>'
                  % (m.group(2), m.group(1)),
        out)
    out = _MD_CODE_RE.sub(lambda m: "<code>%s</code>" % m.group(1), out)
    return out


def _render_blocks_html(blocks: list) -> str:
    """Turn structured content blocks into readable article HTML."""
    parts: list[str] = []
    pending_list: list[str] = []
    seen_img_srcs: set[str] = set()

    def flush_list() -> None:
        if pending_list:
            parts.append("<ul>" + "".join("<li>%s</li>" % i for i in pending_list) + "</ul>")
            pending_list.clear()

    for block in blocks if isinstance(blocks, list) else []:
        if not isinstance(block, dict):
            continue
        btype = block.get("type", "")
        tag = str(block.get("tag", "")).lower()
        text = str(block.get("text", "")).strip()

        if btype == "list_item":
            if text:
                pending_list.append(_rich(text))
            continue

        flush_list()

        if btype == "heading":
            level = int(tag[1]) if len(tag) == 2 and tag[1].isdigit() else 2
            if text:
                parts.append('<h%d><span class="hlvl">H%d</span>%s</h%d>'
                             % (level, level, _rich(text), level))

        elif btype == "faq_question":
            if text:
                parts.append('<p class="faq-q">%s</p>' % _rich(text))

        elif btype == "paragraph":
            if not text:
                continue
            if text.startswith(">"):
                parts.append("<blockquote>%s</blockquote>" % _rich(text.lstrip("> ").strip()))
            else:
                parts.append("<p>%s</p>" % _rich(text))

        elif btype == "image":
            src = _sanitize_xml_text(str(block.get("src", ""))).strip()
            alt = _sanitize_xml_text(str(block.get("alt", ""))).strip()
            img_title = _sanitize_xml_text(str(block.get("title", ""))).strip()
            width = _sanitize_xml_text(str(block.get("width", ""))).strip()
            height = _sanitize_xml_text(str(block.get("height", ""))).strip()
            if src and src in seen_img_srcs:
                continue
            if src:
                seen_img_srcs.add(src)
            cap = ['<span class="altline">alt: %s</span>' % (_esc(alt) if alt else "<em>(missing)</em>")]
            if img_title:
                cap.append("title: %s" % _esc(img_title))
            if width and height:
                cap.append("%s&times;%s" % (_esc(width), _esc(height)))
            if src:
                cap.append('%s &middot; <a href="%s" target="_blank" rel="noopener noreferrer">source</a>'
                           % (_esc(_image_variant_summary(src)), _esc(src)))
            img_html = ('<img src="%s" alt="%s" loading="lazy">' % (_esc(src), _esc(alt))
                        if src else "")
            parts.append("<figure>%s<figcaption>%s</figcaption></figure>"
                         % (img_html, " &nbsp;&middot;&nbsp; ".join(cap)))

        elif btype == "table":
            rows = block.get("rows") or []
            if not rows:
                continue
            max_cols = max(len(r) for r in rows)
            norm = [list(r) + [""] * (max_cols - len(r)) for r in rows]
            head = "".join("<th>%s</th>" % _rich(str(c)) for c in norm[0])
            body = "".join(
                "<tr>%s</tr>" % "".join("<td>%s</td>" % _rich(str(c)) for c in row)
                for row in norm[1:]
            )
            parts.append('<div class="tbl-wrap"><table><thead><tr>%s</tr></thead>'
                         "<tbody>%s</tbody></table></div>" % (head, body))

        elif btype == "code":
            code = str(block.get("text", "")).rstrip()
            if code.strip():
                lang = _esc(block.get("language", ""))
                cls = ' class="lang-%s"' % lang if lang else ""
                parts.append("<pre><code%s>%s</code></pre>" % (cls, _esc(code)))

        elif btype == "cta":
            href = _sanitize_xml_text(str(block.get("href", ""))).strip()
            if text and href:
                parts.append('<a class="cta" href="%s" target="_blank" rel="noopener noreferrer">%s</a>'
                             % (_esc(href), _esc(text)))

    flush_list()
    return "\n".join(parts) if parts else "<p><em>No readable content was extracted.</em></p>"


def _render_result_html(index: int, result: dict) -> str:
    url = _sanitize_xml_text(str(result.get("url", "")))
    title = _sanitize_xml_text(str(result.get("title", ""))) or "(No title)"
    error = _sanitize_xml_text(str(result.get("error", "")))
    details = result.get("details") if isinstance(result.get("details"), dict) else {}

    head = ['<section class="page" id="p%d">' % index,
            '<div class="pg-head">',
            '<div class="pg-num">URL %d</div>' % index,
            "<h2>%s</h2>" % _esc(title)]
    if url:
        head.append('<a class="src" href="%s" target="_blank" rel="noopener noreferrer">%s</a>'
                    % (_esc(url), _esc(url)))

    if error:
        head.append("</div>")
        head.append('<p class="err">Could not scrape: %s</p>' % _esc(error))
        head.append("</section>")
        return "\n".join(head)

    metadata = details.get("metadata") if isinstance(details.get("metadata"), dict) else {}
    blocks = details.get("content_blocks") or result.get("content_blocks") or []
    words = len(str(result.get("text", "")).split())

    chips: list[str] = []
    if details.get("domain"):
        chips.append(str(details["domain"]))
    if details.get("status_code"):
        chips.append("HTTP %s" % details["status_code"])
    if metadata.get("language"):
        chips.append("Lang: %s" % metadata["language"])
    if words:
        chips.append(f"{words:,} words")
    if details.get("total_images"):
        chips.append("%s images" % details["total_images"])
    if details.get("total_links"):
        chips.append("%s links (%s int / %s ext)"
                     % (details["total_links"], details.get("internal_links", 0),
                        details.get("external_links", 0)))
    chip_html = "".join('<span class="chip">%s</span>' % _esc(c) for c in chips)
    if chip_html:
        head.append('<div class="chips">%s</div>' % chip_html)

    breadcrumb = details.get("breadcrumb") if isinstance(details.get("breadcrumb"), list) else []
    crumbs = []
    for c in breadcrumb:
        if not isinstance(c, dict):
            continue
        label = _esc(c.get("text", "")) or "Link"
        href = _sanitize_xml_text(str(c.get("href", ""))).strip()
        crumbs.append('<a href="%s" target="_blank" rel="noopener noreferrer">%s</a>'
                      % (_esc(href), label) if href else label)
    if crumbs:
        head.append('<div class="crumb">%s</div>' % " &rsaquo; ".join(crumbs))

    description = _sanitize_xml_text(str(metadata.get("description", ""))).strip()
    if description:
        head.append('<p class="desc">%s</p>' % _esc(description))
    head.append("</div>")

    # Metadata drawer
    rows = [("Final URL", details.get("final_url", url)),
            ("Domain", details.get("domain", "")),
            ("Status", details.get("status_code", "")),
            ("Language", metadata.get("language", "")),
            ("Canonical", metadata.get("canonical", "")),
            ("Description", description),
            ("Keywords", metadata.get("keywords", "")),
            ("Author", metadata.get("author", ""))]
    meta_tags = metadata.get("meta_tags") if isinstance(metadata.get("meta_tags"), list) else []
    for item in meta_tags[:80]:
        if isinstance(item, dict) and item.get("key") and item.get("value"):
            rows.append((item.get("key"), item.get("value")))
    row_html = "".join("<tr><td>%s</td><td>%s</td></tr>"
                       % (_esc(k), _esc(v) if str(v or "").strip() else "N/A")
                       for k, v in rows)
    head.append('<details class="meta"><summary>Metadata &amp; meta tags (%d)</summary>'
                '<div class="body"><table><tbody>%s</tbody></table></div></details>'
                % (len(rows), row_html))

    head.append('<article class="content">%s</article>' % _render_blocks_html(blocks))
    head.append('<a class="totop" href="#top">&uarr; Back to top</a>')
    head.append("</section>")
    return "\n".join(head)


def build_readable_html(results: list[dict], doc_title: str = "Scraped Content") -> str:
    """Build a self-contained, readable HTML page from scraped results."""
    results = results or []
    generated = datetime.now().strftime("%d %b %Y, %H:%M")
    ok = sum(1 for r in results if not r.get("error"))
    failed = len(results) - ok

    toc = ""
    if len(results) > 1:
        items = "".join(
            '<li><a href="#p%d">%s</a></li>'
            % (i, _esc(r.get("title") or r.get("url") or "Untitled"))
            for i, r in enumerate(results, 1))
        toc = '<nav class="toc"><h2>Pages in this report</h2><ol>%s</ol></nav>' % items

    sections = "\n".join(_render_result_html(i, r) for i, r in enumerate(results, 1))
    summary = "%d page%s &middot; %d scraped%s &middot; %s" % (
        len(results), "" if len(results) == 1 else "s", ok,
        ", %d failed" % failed if failed else "", _esc(generated))

    return (
        "<!DOCTYPE html>\n"
        '<html lang="en"><head>\n'
        '<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
        '<meta name="robots" content="noindex">\n'
        "<title>%s</title>\n<style>%s</style>\n</head>\n"
        '<body id="top">\n'
        '<div class="topbar"><div class="inner">'
        "<div><h1>%s</h1><div class=\"sub\">%s</div></div>"
        '<div class="tools">'
        '<label><input type="checkbox" id="t-images" checked> Images</label>'
        '<label><input type="checkbox" id="t-levels"> Heading tags</label>'
        "</div></div></div>\n"
        '<div class="wrap">%s\n%s\n'
        '<p class="foot">Generated by URL Scraper</p></div>\n'
        "<script>%s</script>\n</body></html>\n"
        % (_esc(doc_title), _READABLE_CSS, _esc(doc_title), summary,
           toc, sections, _READABLE_JS)
    )


def build_readable_html_file(results: list[dict], output_path: str,
                             doc_title: str = "Scraped Content") -> None:
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(build_readable_html(results, doc_title))


def main():
    parser = argparse.ArgumentParser(description="Scrape URLs and save to Word doc.")
    parser.add_argument("input", help="Excel (.xlsx) or CSV file with URLs")
    parser.add_argument("-o", "--output", default="scraped_content.docx",
                        help="Output .docx filename (default: scraped_content.docx)")
    parser.add_argument("--html", nargs="?", const="scraped_content.html", default=None,
                        metavar="PATH",
                        help="Also write a readable, self-contained HTML page "
                             "(default: scraped_content.html)")
    parser.add_argument("--md", nargs="?", const="scraped_content.md", default=None,
                        metavar="PATH",
                        help="Also write the SEO markdown report "
                             "(default: scraped_content.md)")
    parser.add_argument("--no-docx", action="store_true",
                        help="Skip the Word output (use with --html / --md)")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"❌ File not found: {args.input}")
        sys.exit(1)

    print(f"📂 Reading URLs from: {args.input}")
    urls = read_urls(args.input)

    if not urls:
        print("❌ No URLs found in the file.")
        sys.exit(1)

    print(f"🔗 Found {len(urls)} URL(s)\n")

    results = []
    for i, url in enumerate(urls, 1):
        print(f"[{i}/{len(urls)}] Scraping: {url}")
        resolved_url, title, text, error, content_blocks, details = scrape_url(url)

        if error:
            print(f"    ⚠  {error}")
        else:
            words = len(text.split())
            print(f"    ✅ Title: {title[:60]}  |  ~{words} words")

        results.append({
            "url": resolved_url,
            "title": title,
            "text": text,
            "error": error,
            "content_blocks": content_blocks,
            "details": details,
        })

        if i < len(urls):
            time.sleep(DELAY_BETWEEN_REQUESTS)

    if not args.no_docx:
        print(f"\n💾 Saving to: {args.output}")
        build_docx(results, args.output)
        print(f"✅ Done! File saved: {args.output}")

    if args.md:
        with open(args.md, "w", encoding="utf-8") as f:
            f.write(build_seo_text_report(results))
        print(f"✅ Markdown saved: {args.md}")

    if args.html:
        build_readable_html_file(results, args.html)
        print(f"✅ Readable HTML saved: {args.html}  (open it in a browser)")


if __name__ == "__main__":
    main()
