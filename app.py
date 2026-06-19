"""
Flask web frontend for the URL Scraper tool.
"""
import os, sys, json, queue, threading, tempfile, time, uuid, atexit
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, Response, stream_with_context

sys.path.insert(0, os.path.dirname(__file__))
from url_scraper import read_urls, scrape_url, build_docx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"}
STREAM_HEARTBEAT_SECONDS = 10
jobs: dict[str, dict] = {}
_output_files: list[str] = []


def _safe_int(raw, default=0):
    try:
        return int(raw)
    except (TypeError, ValueError):
        return default


def _cleanup():
    for f in _output_files:
        try:
            if os.path.exists(f):
                os.unlink(f)
        except Exception:
            pass

atexit.register(_cleanup)


def allowed_file(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS


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
        "events": queue.Queue(),
        "replay": [],
        "next_event_id": 1,
        "output_path": None,
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
    results = []
    timeout  = settings["timeout"]
    delay    = settings["delay"]
    max_text = settings["max_text"]

    for i, url in enumerate(urls, 1):
        _emit(job, {"type": "progress", "index": i, "total": len(urls),
                    "url": url, "status": "scraping"})

        resolved_url, title, text, error, content_blocks = scrape_url(url, timeout=timeout, max_text=max_text)
        results.append({"url": resolved_url, "title": title, "text": text, "error": error, "content_blocks": content_blocks})
        job["done"] = i

        _emit(job, {
            "type": "progress", "index": i, "total": len(urls),
            "url": url, "title": title,
            "status": "error" if error else "done",
            "error": error,
            "words": len(text.split()) if text else 0,
        })

        if i < len(urls):
            time.sleep(delay)

    out_path = os.path.join(tempfile.gettempdir(), f"scraped_{job_id}.docx")
    try:
        build_docx(results, out_path)
        _output_files.append(out_path)
        job["output_path"] = out_path
        job["status"] = "complete"
        ok  = sum(1 for r in results if not r["error"])
        err = sum(1 for r in results if r["error"])
        total_words = sum(len(r["text"].split()) for r in results if not r["error"])
        _emit(job, {"type": "complete", "ok": ok, "errors": err, "total_words": total_words})
    except Exception as e:
        job["status"] = "error"
        _emit(job, {"type": "error", "message": str(e)})


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
    if job_id not in jobs:
        return jsonify({"error": "Unknown job"}), 404
    path = jobs[job_id].get("output_path")
    if not path or not os.path.exists(path):
        return jsonify({"error": "Output file not ready"}), 404
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        path, as_attachment=True,
        download_name=f"scraped_content_{timestamp}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    print("  URL Scraper - Web UI")
    print("  Open http://localhost:5000 in your browser")
    print("  Press Ctrl+C to stop\n")
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)
