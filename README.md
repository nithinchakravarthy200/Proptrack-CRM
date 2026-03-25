# PropTrack CRM — Real Estate Intelligence Platform

Full-stack CRM for 108,429 Telangana real estate prospects.

---

## Quick Start

### Step 1 — Install Python 3
Download from https://python.org (if not already installed)

### Step 2 — Run

**Mac / Linux:**
```bash
bash run.sh
```

**Windows:**
```
Double-click run.bat
```

**Manual (any OS):**
```bash
pip install flask flask-cors pandas openpyxl
cd backend
python app.py
```

### Step 3 — Open browser
→ **http://localhost:5000**

> ⏱️ First launch takes ~30–60 seconds to load 108K records into the database.
> Subsequent launches are instant.

---

## Project Structure

```
proptrack/
├── backend/
│   ├── app.py              ← Flask API server
│   ├── requirements.txt    ← Python dependencies
│   └── proptrack.db        ← SQLite DB (auto-created on first run)
├── frontend/
│   └── index.html          ← Full CRM UI
├── data/
│   └── consolidated_data.xlsx   ← Your 108K records
├── run.sh                  ← Mac/Linux launcher
├── run.bat                 ← Windows launcher
└── README.md
```

---

## Features

| Page | Description |
|------|-------------|
| **Dashboard** | KPIs, segment bars, top districts, companies, car brands, top scored leads |
| **AI Assistant** | Live Claude-powered chat — pitch scripts, outreach templates, strategy |
| **Contacts** | Full 108K database, real-time search, filter by segment + district, pagination |
| **Pipeline** | Kanban deal tracker — add, move between stages, delete, all saved to DB |
| **Segments** | Deep-dive per segment with AI pitch generator buttons |
| **Analytics** | Income distribution, car brands, GST commissionerates, priority table |

---

## API Reference

```
GET  /api/stats
GET  /api/contacts?segment=it&district=Hyderabad&q=rahul&page=1&per_page=48&sort=score
GET  /api/contacts/:id
POST /api/contacts
PUT  /api/contacts/:id
DEL  /api/contacts/:id
GET  /api/pipeline
POST /api/pipeline
PUT  /api/pipeline/:id
DEL  /api/pipeline/:id
GET  /api/districts
GET  /api/analytics
GET  /api/search?q=rahul
```

---

## Data Loaded from Excel

| Sheet | Records | Key Fields |
|-------|---------|------------|
| IT EMPLOYEES | 20,338 | Name, Mobile, Company, Annual Income |
| GST TAX PAYERS | 32,476 | Name, GSTIN, Email, Commissionerate, Division |
| VEHICLE OWNERS | 30,261 | Name, Car Brand/Variant, Vehicle No, Finance, Insurance |
| HDFC CUSTOMERS | 16,595 | Name, Mobile (×2), Company |
| TEACHERS | 8,759 | Name, School, District, Category, Management |

---

## Tech Stack

- **Backend**: Python 3 + Flask + SQLite (WAL mode, indexed)
- **Frontend**: Vanilla HTML/CSS/JS — no framework, no build step
- **AI**: Claude Sonnet via Anthropic API
- **Database**: SQLite auto-created at `backend/proptrack.db`

