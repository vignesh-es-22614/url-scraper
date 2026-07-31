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

import sys
import os
import time
import argparse
import re
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


def _url_candidates(url: str) -> list[str]:
    url = _sanitize_xml_text(url).strip()
    if not url:
        return []
    if url.startswith(("http://", "https://")):
        return [url]
    return [f"https://{url}", f"http://{url}"]


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


def _is_internal_link(base_url: str, href: str) -> bool:
    base_host = urlparse(base_url).netloc.lower()
    href_host = urlparse(href).netloc.lower()
    return bool(base_host) and href_host == base_host


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
            if isinstance(child, NavigableString):
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

        for child in node.children:
            if isinstance(child, Tag):
                walk(child)

    def _add_img_block(img: Tag) -> None:
        src = _resolve_url(base_url, img.get("src", ""))
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
        src = _resolve_url(base_url, img.get("src", ""))
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
    last_error = ""

    for candidate in candidates:
        try:
            resp = requests.get(candidate, headers=HEADERS, timeout=_timeout)
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "html.parser")
            content_soup = BeautifulSoup(resp.text, "html.parser")

            title_tag = soup.find("title")
            title = _sanitize_xml_text(title_tag.get_text(strip=True) if title_tag else "(No title)")

            metadata = _collect_metadata(soup, resp.url)
            headings = _collect_headings(soup)
            links = _collect_links(soup, resp.url)
            images = _collect_images(soup, resp.url)
            breadcrumb = _collect_breadcrumb(soup, resp.url)
            cta_buttons = _collect_cta_buttons(soup, resp.url)

            # ── Strip chrome / template areas ───────────────────────
            _TMPL_TAGS = ["script", "style", "nav", "footer", "header",
                          "aside", "noscript", "iframe", "svg"]
            _TMPL_CLS  = ["sidebar", "side-bar", "widget", "banner",
                          "popup", "modal", "cookie", "promo", "overlay",
                          "breadcrumb", "related", "social", "share",
                          "advertisement", "ads", "recommended", "toc",
                          "table-of-content", "floating", "sticky-"]
            _TMPL_IDS  = ["sidebar", "nav", "menu", "navigation",
                          "breadcrumb", "cookie", "banner", "popup",
                          "modal", "footer", "header", "toc"]

            for t in content_soup(_TMPL_TAGS):
                t.decompose()
            for pattern in _TMPL_CLS:
                for el in content_soup.find_all(
                        class_=lambda c, p=pattern:
                        c and any(p in v.lower() for v in (c if isinstance(c, list) else [c]))):
                    el.decompose()
            for pattern in _TMPL_IDS:
                for el in content_soup.find_all(
                        id=lambda i, p=pattern: i and p in i.lower()):
                    el.decompose()

            main = (content_soup.find("main") or content_soup.find("article") or
                    content_soup.find(id=lambda i: i and "content" in i.lower()) or
                    content_soup.find(class_=lambda c: c and "content" in
                        " ".join(c if isinstance(c, list) else [c]).lower()) or
                    content_soup.find("body") or content_soup)

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

def main():
    parser = argparse.ArgumentParser(description="Scrape URLs and save to Word doc.")
    parser.add_argument("input", help="Excel (.xlsx) or CSV file with URLs")
    parser.add_argument("-o", "--output", default="scraped_content.docx",
                        help="Output .docx filename (default: scraped_content.docx)")
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

    print(f"\n💾 Saving to: {args.output}")
    build_docx(results, args.output)
    print(f"✅ Done! File saved: {args.output}")


if __name__ == "__main__":
    main()
