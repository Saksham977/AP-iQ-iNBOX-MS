# AP iQ iNBOX — Accounts Payable Working Dashboard

A local, single-page browser tool for reconciling AP invoice status, matching payments, and managing vendor ledger data. Everything runs on your own machine — no cloud database, no external services except three explicit on-click Claude call sites.

---

## What it does

| Panel | Name | What happens |
|---|---|---|
| 1 | **Run Bulk Status** | Match a vendor invoice file (or paste invoice numbers) against the Business Central Vendor Ledger export. Results shown inline; Excel saved only on explicit Download. |
| 2 | **Detect Excel Corrections** | Diff the downloaded output against a hand-edited version to log every change automatically. |
| 3 | **Payment Category** | Match a screenshot, pasted email, PDF, or Excel invoice against the ledger. Extracts fields via Claude (image/text) or directly (PDF/Excel). |
| 4 | **Payment Insights** | Aggregate views over the whole ledger — due, overdue, ready-for-payment, cash-outflow — with filters. No AI call. |
| 5 | **Draft Reply** | Draft a short AP reply email for a single invoice row via Claude. |
| 6 | **Corrections Log** | Lifetime history of every hand-correction, downloadable as Excel. |
| 7 | **Update AI Learning** | Batch unreviewed corrections → Claude proposes a skill update → you approve or reject. |

---

## Prerequisites

- **Python 3.9 or newer**
- **pip** (comes with Python)
- **Claude** — one of:
  - [Claude CLI](https://claude.ai/code) installed and logged in *(free, no API key needed)*
  - `ANTHROPIC_API_KEY` set in your environment *(Anthropic API)*
  - Any OpenAI-compatible provider key *(OpenAI, Azure, Gemini, OpenRouter, Ollama…)*

---

## Quick start

```bash
# 1. Clone or download the repo
git clone https://github.com/<your-username>/ap-dashboard.git
cd ap-dashboard

# 2. Install dependencies
pip install -r requirements.txt

# 3. Start the server
#    Windows:
run.bat
#    Mac / Linux:
bash run.sh
#    Or directly:
python server.py

# 4. Open in browser
#    http://localhost:5000
```

The `data/` directories are created automatically on first run — nothing to set up manually.

---

## Claude connection — three-tier fallback, no API key required

All three Claude call sites (`call_claude` / `call_claude_extract` in `server.py`) try, in order:

1. **`ANTHROPIC_API_KEY` set** → Anthropic API directly (Haiku for draft/learning, Sonnet for extraction).
2. **`OTHER_MODEL_API_KEY` set** → any OpenAI-compatible endpoint.
3. **Neither set** → the locally installed Claude CLI (`claude -p`), which uses your logged-in session. This is how the app runs with zero configuration.

The header shows a live status chip — **Claude: API**, **Claude: CLI**, or a red/amber error — so you know immediately which tier is active.

To configure Tier 2, copy `.env.example` to `.env` and fill in the values:

```bash
OTHER_MODEL_API_KEY=sk-...
OTHER_MODEL_BASE_URL=https://api.openai.com/v1   # or Azure / Ollama endpoint
OTHER_MODEL_NAME=gpt-4o-mini
```

---

## Ledger cache — fast after the first load

The BC Vendor Ledger export (50 MB+) is slow to parse from Excel. After the first load the server saves a binary cache to `data/ledger_cache/`. Every subsequent start — including after a server restart — loads from cache in ~1–2 seconds.

The cache is keyed to the file's absolute path and modification time. When you replace the ledger with a new daily export, the old cache is discarded automatically and a fresh one is built.

`data/ledger_cache/` is in `.gitignore` — each machine builds its own cache from its own ledger file.

---

## Project structure

```
ap-dashboard/
├── server.py                          # Flask server — all routes, Claude call sites, ledger cache
├── payment_category.py                # Ledger loading, invoice matching, confidence scoring
├── payment_insights.py                # Aggregate views — due/overdue/payment-ready/cash-outflow
├── run.bat / run.sh                   # One-click startup (Windows / Mac & Linux)
├── requirements.txt                   # pip dependencies
├── .env.example                       # Optional environment variables (copy → .env)
│
├── static/
│   ├── dashboard.html                 # Single-page UI — all panels, all JS
│   └── logo.png
│
├── skills_mirror/
│   ├── bulk-status/
│   │   ├── SKILL.md                   # Skill definition (reference)
│   │   └── scripts/match_status.py   # Bulk invoice matcher + vendor column detection
│   └── ap-category-parameters/
│       ├── SKILL.md
│       └── references/               # category-map.md, house-format.md
│
├── watcher/
│   └── detect_changes.py             # Excel diff — compares snapshot vs edited file
│
├── tests/
│   └── test_core.py                  # 62 unit tests (pytest)
│
└── data/                             # Auto-created, git-ignored
    ├── ledger_cache/                 # Binary pickle — rebuilt per machine
    ├── snapshots/                    # Bulk status outputs
    ├── payment_category/             # Payment category results
    ├── payment_insights/             # Insights exports
    └── proposals/                    # AI learning proposals (pending review)
```

---

## Running tests

```bash
pytest tests/
# Expected: 62 passed
```

---

## What is NOT in this repo

The `.gitignore` excludes everything in `data/` — ledger cache, snapshots, payment outputs, correction logs, and audit logs. These contain real financial and vendor data and must never be committed.

The `.env` file (your actual API keys) is also excluded. Use `.env.example` as the template.

---

## Security note

The Flask server binds to `localhost:5000` with no authentication. It is designed for single-user local use only. Do not expose it to a network or the internet.

---

*Designed and built by Saksham Setia · MeetingSelect*
