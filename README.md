# AHUSE / AUHSE — Autonomous HSE Intelligence System

Flask web app for **AI-assisted construction safety inspections**: multi-image and optional video upload, **YOLO** hazard detection via a remote API, **OpenRouter** vision-LM for structured HSE reports (JSON + DOCX), user accounts with optional **Google OAuth**, and an inspection history dashboard.

**Repository:** [github.com/MahmoudOmar03/AHUSE](https://github.com/MahmoudOmar03/AHUSE)

## Stack

- **Backend:** Python 3, Flask 3, SQLAlchemy (SQLite), Flask-Login, Authlib (Google OAuth)
- **Frontend:** Jinja2 templates, vanilla CSS (`static/css/style.css`), no npm build
- **ML / APIs:** `yolo_api_client` + OpenCV (video frames) → your YOLO HTTP endpoint; `LLM_VLM.py` → OpenRouter chat completions with vision

## Quick start

```bash
git clone https://github.com/MahmoudOmar03/AHUSE.git
cd AHUSE
python -m pip install -r requirements.txt
```

Create a **`.env`** file in the project root (never commit it). Minimum for image inspections with cloud YOLO + reports:

```env
FLASK_SECRET_KEY=your-long-random-secret
OPENROUTER_API_KEY=sk-or-v1-...
YOLO_API_URL=https://your-yolo-service/predict
YOLO_API_KEY=your-bearer-token
```

Optional — **Google sign-in** (all three required for the button to enable):

```env
GOOGLE_OAUTH_CLIENT_ID=...
GOOGLE_OAUTH_CLIENT_SECRET=...
GOOGLE_OAUTH_REDIRECT_URI=http://127.0.0.1:8000/auth/google/callback
```

Run:

```bash
python app.py
```

Open `http://127.0.0.1:8000` (or the host/port in `PORT`).

## Environment variables

| Variable | Purpose |
|----------|---------|
| `FLASK_SECRET_KEY` | Session cookie signing (use a strong value in production). |
| `OPENROUTER_API_KEY` | OpenRouter API key for HSE report generation. |
| `OPENROUTER_BASE_URL` | Optional. Default `https://openrouter.ai/api/v1`. |
| `OPENROUTER_VLM_MODEL` | Optional. Default `openai/gpt-4o-mini` (vision). |
| `YOLO_API_URL` | Full URL of YOLO `POST` predict endpoint. |
| `YOLO_API_KEY` | Bearer token for YOLO API. |
| `YOLO_USE_LOCAL` | Set to `1` to use local Ultralytics + `YOLO_MODEL_PATH` instead of the API. Place your weights file locally (e.g. `best.pt`); weight files are not committed to this repo. |
| `GOOGLE_OAUTH_*` | Client ID, secret, and redirect URI (exact match to Google Cloud console). |
| `PORT` | Dev server port (default `8000`). |
| `MAX_UPLOAD_BYTES` | Max request body size (default ~350 MB). |

`python-dotenv` loads `.env` automatically when the app starts.

## Features

- **Upload modes:** images only (1–20 files, 100 MB total), video only (250 MB max), or both.
- **Reports:** DOCX + JSON in `outputs/<timestamp>/`; video runs add `video_yolo_summary.json` where applicable.
- **Auth:** Email/password register & login; Google OAuth links to the same `User` model.
- **SQLite:** Database file `auhse.db`; startup runs lightweight `ALTER TABLE` migrations in `db_schema.py` for older DBs.

## Project layout

| Path | Role |
|------|------|
| `app.py` | Flask app, routes, uploads, OAuth |
| `models.py` | SQLAlchemy models |
| `pipeline.py` | YOLO gate + HSE pipeline orchestration |
| `yolo_api_client.py` | HTTP client for YOLO predict API |
| `yolo_api_video.py` | Video frame sampling + same API |
| `LLM_VLM.py` | OpenRouter multimodal call + JSON/DOCX generation |
| `templates/` | Jinja HTML |
| `static/` | CSS, JS |

## Production notes

- Serve with **HTTPS** in production; set `GOOGLE_OAUTH_REDIRECT_URI` to your public callback URL.
- Put the app behind a reverse proxy that forwards `X-Forwarded-Proto` (and prefix if mounted on a subpath).
- Do not commit `.env`, API keys, or customer uploads.

## License

This project is intended to align with the repository license on GitHub ([Apache-2.0](https://github.com/MahmoudOmar03/AHUSE/blob/main/LICENSE) where applicable). Add or adjust a `LICENSE` file in this repo if you maintain a separate copy.
