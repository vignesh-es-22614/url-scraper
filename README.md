# URL Scraper

Extracts the readable content of a web page — headings, paragraphs, lists,
tables, images, CTAs and FAQ questions — and publishes it as a readable HTML
page, an SEO markdown report, or a Word document.

**Published site — GitHub Pages, served from `docs/`:**
https://vignesh-es-22614.github.io/url-scraper/

## Two ways to run it

| | Static site (GitHub Pages) | Web app (local / Render) |
|---|---|---|
| Scrapes | the URLs in `urls.txt` | anything you paste or upload |
| Runs on | GitHub Actions runners | your machine, or a Python host |
| Output | published at the Pages URL | downloaded in the browser |
| Speed | ~1 min per Action run | seconds |

GitHub Pages cannot run Python, so the scraping happens in the Action and Pages
serves what it produced. Live, type-a-URL-and-go scraping needs the web app.

## The published site

`urls.txt` holds the URL list, one per line. Edit it and push — the
**Scrape and publish** Action rebuilds `docs/` and commits the result:

```
docs/index.html          overview of everything scraped
docs/pages/<slug>.html   readable view of one page
docs/report.md           SEO markdown report
docs/scraped.docx        Word export
```

To scrape something ad hoc without editing `urls.txt`, open **Actions → Scrape
and publish → Run workflow** and paste URLs into the `urls` box.

The Action also runs weekly (Mondays, 05:00 UTC) so published content stays
fresh.

## The web app

```bash
pip install -r requirements.txt
python app.py            # http://localhost:5000
```

Upload an `.xlsx`/`.csv` of URLs or paste them in, then take the output as
**Read page** (readable HTML), **Download HTML**, **SEO (.md)** or
**Word (.docx)**. Deploys to Render from `render.yaml`, or to Zoho Catalyst
AppSail — see `CATALYST_DEPLOY.md`.

Note: job state lives in process memory, so run a single worker. `Procfile` and
`render.yaml` already use one threaded `gthread` worker for this reason.

## Command line

```bash
python url_scraper.py sample_urls.xlsx                      # -> scraped_content.docx
python url_scraper.py sample_urls.xlsx --html --md --no-docx
python build_site.py                                        # rebuild docs/ locally
```

Input is forgiving: `page: https://x`, `- https://x`, `1. https://x`,
`<https://x>` and bare hosts like `example.com` all resolve to the same URL.

## How extraction works

`scrape_url()` fetches the page, strips navigation/cookie/ad chrome, picks the
element holding the bulk of the body text, then walks it in document order into
typed blocks (`heading`, `paragraph`, `list_item`, `image`, `table`, `code`,
`cta`, `faq_question`). Every exporter renders those same blocks, so the HTML,
markdown and Word outputs always agree.

Two rules keep content from being lost, both learned from real pages:

- Chrome is removed by **whole-token** match, never substring — `ads` must not
  match `downloads` — and a block containing an `h1`/`h2` is never treated as
  chrome, because marketing pages wrap the headline in `class="banner-content"`.
- The content root is chosen by **share of body text** (≥60%), not by a class
  name containing `content`, which used to select a single tab panel and throw
  away two thirds of the page.
