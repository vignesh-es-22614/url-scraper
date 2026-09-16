# Deploying the URL Scraper to Zoho Catalyst (AppSail)

This app is a long-running Flask process (SSE streaming, background threads, in-memory
job state), so it belongs on **Catalyst AppSail** — a Catalyst-managed runtime container —
**not** Catalyst Functions (those are stateless and time-limited).

Everything in the repo is already prepped. You only need to run the login + deploy steps.

---

## What was already done for you

- `from __future__ import annotations` added to `app.py` and `url_scraper.py`
  (Catalyst's managed Python runtime is **3.9**; the code used `X | None` syntax that 3.9
  would otherwise reject at import time).
- `app.py` now reads the port from `X_ZOHO_CATALYST_LISTEN_PORT` (falling back to `PORT`,
  then `5000`) and runs threaded.
- `build_catalyst.ps1` assembles a clean `catalyst_build/` folder with ONLY the runtime
  files — so `.venv/`, `.git/`, and the sample `.docx`/`.xlsx` files are never uploaded.
- `render.yaml` and `Procfile` are untouched, so the Render deployment still works.

---

## One-time deploy

```powershell
# 1. Rebuild the clean deploy folder (run again whenever you change the app)
.\build_catalyst.ps1

# 2. Log in (opens your browser for Zoho OAuth)
catalyst login

# 3. Initialize AppSail in this project
catalyst init
#    Choose:
#      - AppSail
#      - Catalyst-Managed Runtime  (not Docker Image)
#      - Your own app
#      - Stack:        Python
#      - Source / build path: point it at the catalyst_build folder
#      - App name:     url-scraper (or anything)

# 4. Set the start command (see below), then deploy
catalyst deploy
```

### Start command (put this in the generated `app-config.json` -> `command`)

AppSail runs the command **without a shell**, so wrap it in `sh -c` so the port variable
expands:

```
sh -c 'python3 -m gunicorn app:app --bind 0.0.0.0:$X_ZOHO_CATALYST_LISTEN_PORT --workers 1 --threads 8 --worker-class gthread --timeout 120'
```

- Keep `--workers 1` — see "Single instance" below.
- `python3 -m gunicorn` works whether gunicorn is installed by the build or vendored.

---

## Two things to keep in mind

### 1. Keep it to a SINGLE instance
Job state lives in this process's memory (`jobs` dict) and downloadable artifacts live in
the container's `/tmp`. If AppSail scales to 2+ instances, a `/stream` or `/download`
request can hit an instance that never ran the job -> "Unknown job." Do not enable
horizontal scaling until that state is externalized (Catalyst Cache/Data Store for jobs,
Catalyst Stratus/File Store for artifacts).

### 2. If the build does NOT auto-install requirements.txt
Some managed-runtime builds install `requirements.txt`; if yours doesn't and the app fails
to start with a missing-module error, vendor the deps into the build folder and redeploy:

```powershell
.\build_catalyst.ps1 -Vendor
catalyst deploy
```

---

## Building more tools in the same project later

Catalyst is a good base for a tool suite. Once this is live you can add, in the same
project: more AppSail apps, Catalyst **Functions** (for short stateless jobs / webhooks /
cron), **Data Store** (relational), **Cache**, **Stratus/File Store** (object storage),
and **Catalyst Scheduler** (cron). Share them behind one project and one auth domain.
