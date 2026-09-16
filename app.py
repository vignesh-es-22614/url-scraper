"""
Flask web frontend for the URL Scraper tool.
"""
from __future__ import annotations

import os, sys, json, queue, threading, tempfile, time, uuid, concurrent.futures
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, Response, stream_with_context

sys.path.insert(0, os.path.dirname(__file__))
from url_scraper import (
    read_urls, scrape_url, build_seo_text_report, build_seo_docx,
    build_readable_html_file,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"}
STREAM_HEARTBEAT_SECONDS = 3   # keep Render proxy alive
SCRAPE_MAX_WORKERS = 2         # keep low on free tier (0.1 CPU) to avoid starving heartbeat thread
ARTIFACTS_DIR = os.path.join(tempfile.gettempdir(), "url_scraper_artifacts")
jobs: dict[str, dict] = {}

os.makedirs(ARTIFACTS_DIR, exist_ok=True)


def _safe_int(raw, default=0):
    try:
        return int(raw)
    except (TypeError, ValueError):
        return default


def allowed_file(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS


def _artifact_path(job_id: str, ext: str) -> str:
    safe_job = "".join(ch for ch in job_id if ch.isalnum() or ch in {"-", "_"})
    return os.path.join(ARTIFACTS_DIR, f"{safe_job}.{ext}")


def _build_markdown_artifact(job_id: str, results: list[dict]) -> str:
    md_path = _artifact_path(job_id, "md")
    report_text = build_seo_text_report(results)
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(report_text)
    return md_path


def _build_docx_artifact(job_id: str, results: list[dict]) -> str:
    docx_path = _artifact_path(job_id, "docx")
    build_seo_docx(results, docx_path)
    return docx_path


def _build_html_artifact(job_id: str, results: list[dict]) -> str:
    html_path = _artifact_path(job_id, "html")
    build_readable_html_file(results, html_path)
    return html_path


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/scrape", methods=["POST"])
def start_scrape():
    urls = []
    settings = {}

    if "file" in request.files and request.files["file"].filename:
        f = request.files["file"]
        if not allowed_file(f.filename):
            return jsonify({"error": "Only .xlsx, .xls, or .csv files are allowed."}), 400
        ext = os.path.splitext(f.filename)[1].lower()
        tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
        f.save(tmp.name)
        tmp.close()
        try:
            urls = read_urls(tmp.name)
        except Exception as e:
            return jsonify({"error": f"Could not read file: {e}"}), 400
        finally:
            try:
                os.unlink(tmp.name)
            except Exception:
                pass
        settings = {
            "timeout":  int(request.form.get("timeout", 10)),
            "delay":    float(request.form.get("delay", 0)),
            "max_text": int(request.form.get("max_text", 0)),
        }
    else:
        data = request.get_json(silent=True) or {}
        raw = data.get("urls", [])
        urls = [u.strip() for u in raw if isinstance(u, str) and u.strip()]
        settings = {
            "timeout":  int(data.get("timeout", 10)),
            "delay":    float(data.get("delay", 0)),
            "max_text": int(data.get("max_text", 0)),
        }

    if not urls:
        return jsonify({"error": "No URLs provided."}), 400

    settings["timeout"]  = max(5, min(settings["timeout"],  120))
    settings["delay"]    = max(0, min(settings["delay"],    30))
    settings["max_text"] = max(0, min(settings["max_text"], 500000))

    job_id = str(uuid.uuid4())
    jobs[job_id] = {
        "status": "running",
        "total": len(urls),
        "done": 0,
        "results": [],
        "artifact_error": None,
        "events": queue.Queue(),
        "replay": [],
        "next_event_id": 1,
    }

    threading.Thread(target=_run_job, args=(job_id, urls, settings), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(urls)})


def _emit(job: dict, payload: dict):
    event_id = job["next_event_id"]
    job["next_event_id"] += 1
    msg = json.dumps(payload)
    event = (event_id, msg)
    job["events"].put(event)
    job["replay"].append(event)


def _run_job(job_id: str, urls: list[str], settings: dict):
    job = jobs[job_id]
    timeout  = settings["timeout"]
    delay    = settings["delay"]
    max_text = settings["max_text"]
    total    = len(urls)

    # Pre-allocate ordered results so output order matches input order.
    ordered_results: list[dict | None] = [None] * total
    emit_lock = threading.Lock()

    def _scrape_one(idx_url):
        i, url = idx_url          # 0-based index
        display_i = i + 1         # 1-based for UI

        # Stagger workers slightly when delay requested so we don't hit every host at t=0.
        if delay > 0:
            time.sleep(delay * i)

        with emit_lock:
            _emit(job, {"type": "progress", "index": display_i,
                        "total": total, "url": url, "status": "scraping"})

        resolved_url, title, text, error, content_blocks, details = scrape_url(
            url, timeout=timeout, max_text=max_text
        )

        result = {
            "url": resolved_url, "title": title, "text": text,
            "error": error, "content_blocks": content_blocks, "details": details,
        }
        ordered_results[i] = result

        with emit_lock:
            job["done"] += 1
            job["results"] = [r for r in ordered_results if r is not None]
            _emit(job, {
                "type": "progress", "index": display_i, "total": total,
                "url": url, "title": title,
                "status": "error" if error else "done",
                "error": error,
                "words": len(text.split()) if text else 0,
            })

    workers = min(SCRAPE_MAX_WORKERS, total)
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        pool.map(_scrape_one, enumerate(urls))

    final_results = [r for r in ordered_results if r is not None]

    try:
        # Build downloadable artifacts before marking complete so download links work immediately.
        _build_markdown_artifact(job_id, final_results)
        _build_docx_artifact(job_id, final_results)
        _build_html_artifact(job_id, final_results)
        job["status"] = "complete"
        ok  = sum(1 for r in final_results if not r["error"])
        err = sum(1 for r in final_results if r["error"])
        total_words = sum(len(r["text"].split()) for r in final_results if not r["error"])
        _emit(job, {"type": "complete", "ok": ok, "errors": err, "total_words": total_words})
    except Exception as e:
        job["artifact_error"] = str(e)
        job["status"] = "error"
        _emit(job, {"type": "error", "message": str(e)})


def _download_seo_response(job_id: str):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"scraped_content_{timestamp}.md"
    md_path = _artifact_path(job_id, "md")

    if os.path.exists(md_path):
        return send_file(md_path, as_attachment=True, download_name=out_name, mimetype="text/markdown")

    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job. The job may have expired or run on a different instance."}), 404

    if job.get("artifact_error"):
        return jsonify({"error": f"Could not prepare markdown output: {job['artifact_error']}"}), 500

    results = job.get("results") or []
    if not results:
        return jsonify({"error": "Output file not ready"}), 404

    try:
        md_path = _build_markdown_artifact(job_id, results)
    except Exception as e:
        job["artifact_error"] = str(e)
        return jsonify({"error": f"Could not build markdown output: {e}"}), 500

    return send_file(md_path, as_attachment=True, download_name=out_name, mimetype="text/markdown")


@app.route("/download-seo/<job_id>")
def download_seo(job_id: str):
    return _download_seo_response(job_id)


@app.route("/stream/<job_id>")
def stream(job_id: str):
    if job_id not in jobs:
        return jsonify({"error": "Unknown job"}), 404

    last_event_id = _safe_int(request.headers.get("Last-Event-ID"), 0)

    def _sse_event(event_id: int, msg: str) -> str:
        return f"id: {event_id}\ndata: {msg}\n\n"

    def generate():
        job = jobs[job_id]
        seen_event_id = last_event_id

        for event_id, msg in list(job["replay"]):
            if event_id <= seen_event_id:
                continue
            seen_event_id = event_id
            yield _sse_event(event_id, msg)
            if json.loads(msg).get("type") in ("complete", "error"):
                return

        while True:
            try:
                event_id, msg = job["events"].get(timeout=STREAM_HEARTBEAT_SECONDS)
                if event_id <= seen_event_id:
                    continue
                seen_event_id = event_id
                yield _sse_event(event_id, msg)
                if json.loads(msg).get("type") in ("complete", "error"):
                    break
            except queue.Empty:
                if job["status"] != "running":
                    break
                yield 'data: {"type":"heartbeat"}\n\n'

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        },
    )


@app.route("/download/<job_id>")
def download(job_id: str):
    return _download_seo_response(job_id)


def _readable_html_path(job_id: str) -> tuple[str | None, tuple]:
    """Return (path, error_response) for a job's readable HTML artifact."""
    out_path = _artifact_path(job_id, "html")
    if os.path.exists(out_path):
        return out_path, ()

    job = jobs.get(job_id)
    if not job:
        return None, (jsonify({"error": "Unknown job. The job may have expired "
                                        "or run on a different instance."}), 404)
    if job.get("artifact_error"):
        return None, (jsonify({"error": f"Could not prepare readable page: "
                                        f"{job['artifact_error']}"}), 500)

    results = job.get("results") or []
    if not results:
        return None, (jsonify({"error": "Output file not ready"}), 404)

    try:
        return _build_html_artifact(job_id, results), ()
    except Exception as e:
        job["artifact_error"] = str(e)
        return None, (jsonify({"error": f"Could not build readable page: {e}"}), 500)


@app.route("/view/<job_id>")
def view_readable(job_id: str):
    """Open the scraped content as a readable page in the browser."""
    out_path, err = _readable_html_path(job_id)
    if out_path is None:
        return err
    resp = send_file(out_path, mimetype="text/html")
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/download-html/<job_id>")
def download_html(job_id: str):
    out_path, err = _readable_html_path(job_id)
    if out_path is None:
        return err
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(out_path, as_attachment=True,
                     download_name=f"scraped_content_{timestamp}.html",
                     mimetype="text/html")


@app.route("/download-docx/<job_id>")
def download_docx(job_id: str):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"scraped_content_{timestamp}.docx"
    out_path = _artifact_path(job_id, "docx")

    if os.path.exists(out_path):
        return send_file(
            out_path, as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job. The job may have expired or run on a different instance."}), 404

    if job.get("artifact_error"):
        return jsonify({"error": f"Could not prepare Word output: {job['artifact_error']}"}), 500

    results = job.get("results") or []
    if not results:
        return jsonify({"error": "Output file not ready"}), 404

    try:
        out_path = _build_docx_artifact(job_id, results)
    except Exception as e:
        job["artifact_error"] = str(e)
        return jsonify({"error": str(e)}), 500

    return send_file(
        out_path, as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    print("  URL Scraper - Web UI")
    print("  Open http://localhost:5000 in your browser")
    print("  Press Ctrl+C to stop\n")
    # PORT (Render/Heroku) or X_ZOHO_CATALYST_LISTEN_PORT (Catalyst AppSail), else 5000.
    port = int(os.environ.get("PORT") or os.environ.get("X_ZOHO_CATALYST_LISTEN_PORT") or 5000)
    app.run(debug=False, host="0.0.0.0", port=port, threaded=True)
