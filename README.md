# Kill the Quote — AI Procurement Analyst

Prototype for the Aerchain product take-home. One end-to-end flow: a buyer **talks an RFx
into existence**, (simulated) sends it to five vendors, the system **reads whatever they send
back** — Excel, PDF, Word, a phone photo of a rate card, and an email — lands every quote in
**one comparable basis**, and lets the buyer interrogate the result in natural language toward
a **defensible award**.

Only the plumbing (SMTP) is faked. The AI loops are real: extraction and analyst reasoning
both call Google Gemini on the actual source artifacts. Nothing is hardcoded.

## Stack
- Python + FastAPI, single-file static SPA (no build step)
- Google **Gemini** (`gemini-3.5-flash-lite`, free tier) for extraction + analyst reasoning
- One deterministic normalization engine (currency, tax, per-100 / per-pack) — the single
  source of truth used by the comparison, CSV export and the analyst

## Key capabilities
- **Voice-driven RFx co-pilot** — dictate the sourcing need (Web Speech API) and have Gemini
  draft scope, evaluation criteria, questionnaire, terms and guardrails.
- **Simulated email send** — an outbox with delivery states; SMTP is stubbed.
- **Real multimodal extraction** — PDFs and the angled photo are sent to Gemini natively;
  xlsx/docx/txt are parsed locally. Each of the 30 lines gets evidence, confidence and a
  QUOTED / MISSING / AMBIGUOUS status.
- **Trustworthy comparison** — lowest-price highlighting is restricted to quoted, comparable
  lines from qualified suppliers; USD, per-100, per-box, tax-in and freight-extra edges are
  surfaced, not silently normalized away.
- **Grounded analyst** — answers use the pre-computed normalized numbers, cite line IDs and
  vendors, and state assumptions, exclusions and what still needs a human.

## Run locally (Windows)

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

> If your network blocks the default PyPI CDN (TLS handshake failure on
> `files.pythonhosted.org`), install via a mirror:
> ```powershell
> .\.venv\Scripts\python.exe -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
> ```

Provide the Gemini key (any one of these):
- set `GEMINI_API_KEY` in the environment, **or**
- place `Gemini API Key.txt` in this folder or its parent (already present in this workspace).

```powershell
$env:GEMINI_API_KEY = "YOUR_KEY"   # optional if the .txt file is present
.\.venv\Scripts\python.exe -m uvicorn app:app --reload
```

Open http://127.0.0.1:8000

## Deploy (free tier)
- `render.yaml` — Render Blueprint. Set `GEMINI_API_KEY` in the dashboard (never commit it).
- `Dockerfile` — works on any container host (Fly, Railway, Cloud Run).

## API
| Endpoint | Purpose |
|---|---|
| `POST /api/rfx-copilot` | Draft an RFx from a (possibly dictated) prompt |
| `POST /api/send-rfx` | Simulated dispatch to vendors (returns outbox) |
| `POST /api/extract-all` | Live Gemini extraction of all 5 vendor artifacts |
| `POST /api/ask` | Analyst over the normalized dataset + historical context |
| `POST /api/export-csv` | Normalized comparison as CSV |
| `GET /api/source/{vendor}` | Original vendor artifact |

## Note on trust
The prototype prefers an explicit review state over a confident guess. Missing, ambiguous,
alternate-configuration and unqualified responses stay visible and are excluded from automatic
lowest-price selection — the buyer keeps the judgment, the tool removes the re-keying.
