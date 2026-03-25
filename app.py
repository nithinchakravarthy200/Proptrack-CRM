"""
PropTrack CRM — Flask Backend API
Deployment-ready: auto-downloads Excel from Google Drive if not found locally.
Set env var: GDRIVE_FILE_ID=your_google_drive_file_id
"""

from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
import sqlite3, os, math

BASE  = os.path.dirname(os.path.abspath(__file__))
DB    = os.path.join(BASE, "proptrack.db")
DATA  = os.path.join(BASE, "data", "consolidated_data.xlsx")
FRONT = BASE  # frontend index.html is in same folder on Railway

# Google Drive File ID — set this as environment variable on Railway/Render
# or paste your file ID directly here as fallback
GDRIVE_FILE_ID = os.environ.get("GDRIVE_FILE_ID", "")

app = Flask(__name__, static_folder=FRONT, static_url_path="")
CORS(app)

# ─── DB helpers ───────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn

def query(sql, params=(), one=False):
    with get_db() as conn:
        cur = conn.execute(sql, params)
        return cur.fetchone() if one else cur.fetchall()

def execute(sql, params=()):
    with get_db() as conn:
        cur = conn.execute(sql, params)
        conn.commit()
        return cur.lastrowid

# ─── Schema ───────────────────────────────────────────────────────────────────
SCHEMA = """
CREATE TABLE IF NOT EXISTS contacts (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    segment     TEXT    NOT NULL,
    name        TEXT,
    mobile      TEXT,
    mobile_alt  TEXT,
    email       TEXT,
    company     TEXT,
    income      REAL,
    address     TEXT,
    village     TEXT,
    mandal      TEXT,
    district    TEXT,
    state       TEXT,
    pincode     TEXT,
    gstin       TEXT,
    gst_comm    TEXT,
    gst_div     TEXT,
    gst_range   TEXT,
    school_code TEXT,
    school_cat  TEXT,
    school_mgmt TEXT,
    car_brand   TEXT,
    car_variant TEXT,
    vehicle_no  TEXT,
    chassis_no  TEXT,
    veh_finance TEXT,
    veh_insure  TEXT,
    score       INTEGER DEFAULT 50,
    stage       TEXT    DEFAULT 'New Lead',
    notes       TEXT,
    created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS pipeline (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    contact_id  INTEGER,
    name        TEXT,
    segment     TEXT    DEFAULT 'it',
    stage       TEXT    DEFAULT 'New Lead',
    value       REAL,
    notes       TEXT,
    created_at  DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(contact_id) REFERENCES contacts(id) ON DELETE SET NULL
);
CREATE TABLE IF NOT EXISTS activities (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    contact_id  INTEGER,
    type        TEXT,
    note        TEXT,
    created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_seg      ON contacts(segment);
CREATE INDEX IF NOT EXISTS idx_dist     ON contacts(district);
CREATE INDEX IF NOT EXISTS idx_mobile   ON contacts(mobile);
CREATE INDEX IF NOT EXISTS idx_score    ON contacts(score DESC);
CREATE INDEX IF NOT EXISTS idx_company  ON contacts(company);
CREATE INDEX IF NOT EXISTS idx_car      ON contacts(car_brand);
"""

def init_db():
    with get_db() as conn:
        conn.executescript(SCHEMA)
        conn.commit()
    print("[DB] Schema ready.")

# ─── Google Drive downloader ─────────────────────────────────────────────────
def download_from_gdrive(file_id, dest_path):
    import urllib.request, urllib.error
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    print(f"[DOWNLOAD] Fetching Excel from Google Drive (file_id={file_id})...")

    # Try Google Sheets export URL first (works for files opened as Google Sheets)
    # Then fall back to regular Drive download URL
    urls_to_try = [
        f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx",
        f"https://drive.google.com/uc?export=download&id={file_id}&confirm=t",
    ]

    for url in urls_to_try:
        try:
            print(f"  Trying: {url[:60]}...")
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            })
            with urllib.request.urlopen(req, timeout=120) as resp, open(dest_path, "wb") as f:
                total = 0
                while True:
                    chunk = resp.read(1024 * 256)
                    if not chunk:
                        break
                    f.write(chunk)
                    total += len(chunk)
                    print(f"  Downloaded {total//1024//1024} MB...", end="\r")

            # Verify it's a valid xlsx (not an HTML error page)
            if os.path.getsize(dest_path) < 5000:
                print(f"\n  [WARN] File too small ({os.path.getsize(dest_path)} bytes) — might be an error page, trying next URL...")
                os.remove(dest_path)
                continue

            print(f"\n[DOWNLOAD] Done — {os.path.getsize(dest_path)//1024//1024} MB saved to {dest_path}")
            return True

        except Exception as e:
            print(f"\n  [WARN] Failed with {url[:50]}: {e}")
            if os.path.exists(dest_path):
                os.remove(dest_path)
            continue

    print("[DOWNLOAD] All URLs failed.")
    return False

# ─── Data loader ─────────────────────────────────────────────────────────────
def load_excel():
    import pandas as pd

    # If Excel not present locally, try downloading from Google Drive
    if not os.path.exists(DATA):
        if GDRIVE_FILE_ID:
            success = download_from_gdrive(GDRIVE_FILE_ID, DATA)
            if not success:
                print("[WARN] Could not download Excel — skipping data load.")
                return
        else:
            print(f"[WARN] Excel not found at {DATA} and GDRIVE_FILE_ID not set.")
            print("       Set env var: GDRIVE_FILE_ID=your_google_drive_file_id")
            return

    with get_db() as conn:
        n = conn.execute("SELECT COUNT(*) FROM contacts").fetchone()[0]
        if n > 0:
            print(f"[DB] {n:,} records already loaded.")
            return

    print("[LOAD] Reading Excel sheets...")

    def clean(v):
        if v is None: return None
        if isinstance(v, float) and math.isnan(v): return None
        s = str(v).strip()
        return None if s in ("", "nan", "NaN", "None") else s

    rows = []

    df = pd.read_excel(DATA, sheet_name="TEACHERS")
    for _, r in df.iterrows():
        rows.append(("teacher", clean(r.get("FULL NAME / ENTITY")), clean(r.get("MOBILE (PRIMARY)")),
            clean(r.get("MOBILE (ALTERNATE)")), None, None, None,
            clean(r.get("FULL ADDRESS")), clean(r.get("VILLAGE / AREA")),
            clean(r.get("MANDAL")), clean(r.get("DISTRICT")), clean(r.get("STATE")),
            None, None, None, None, None,
            clean(r.get("SCHOOL CODE")), clean(r.get("SCHOOL CATEGORY")), clean(r.get("MANAGEMENT")),
            None, None, None, None, None, None, 52))
    print(f"  TEACHERS:       {len(rows):>7,}")

    base = len(rows)
    df = pd.read_excel(DATA, sheet_name="IT EMPLOYEES")
    for _, r in df.iterrows():
        inc = r.get("ANNUAL INCOME (₹)")
        try:
            inc = float(inc) if inc is not None and not (isinstance(inc, float) and math.isnan(inc)) else None
        except Exception:
            inc = None
        score = 95 if inc and inc > 3_000_000 else 85 if inc and inc > 1_500_000 else 70 if inc and inc > 500_000 else 60
        rows.append(("it", clean(r.get("FULL NAME / ENTITY")), clean(r.get("MOBILE (PRIMARY)")),
            None, None, clean(r.get("COMPANY / ORGANISATION")), inc,
            None, None, None, None, "Telangana",
            None, None, None, None, None, None, None, None,
            None, None, None, None, None, None, score))
    print(f"  IT EMPLOYEES:   {len(rows)-base:>7,}")

    base = len(rows)
    df = pd.read_excel(DATA, sheet_name="GST TAX PAYERS")
    for _, r in df.iterrows():
        mob = r.get("MOBILE (PRIMARY)")
        try:
            mob = str(int(float(mob))) if mob is not None and not (isinstance(mob, float) and math.isnan(mob)) else None
        except Exception:
            mob = clean(mob)
        rows.append(("gst", clean(r.get("FULL NAME / ENTITY")), mob,
            None, clean(r.get("EMAIL")), None, None,
            None, None, None, None, clean(r.get("STATE")),
            clean(r.get("PINCODE")), clean(r.get("GSTIN")),
            clean(r.get("GST COMMISSIONERATE")), clean(r.get("GST DIVISION")), clean(r.get("GST RANGE")),
            None, None, None, None, None, None, None, None, None, 65))
    print(f"  GST PAYERS:     {len(rows)-base:>7,}")

    base = len(rows)
    df = pd.read_excel(DATA, sheet_name="VEHICLE OWNERS")
    for _, r in df.iterrows():
        rows.append(("vehicle", clean(r.get("FULL NAME / ENTITY")), clean(r.get("MOBILE (PRIMARY)")),
            None, None, None, None,
            clean(r.get("FULL ADDRESS")), None, None,
            clean(r.get("DISTRICT")), clean(r.get("STATE")),
            None, None, None, None, None, None, None, None,
            clean(r.get("CAR BRAND")), clean(r.get("CAR VARIANT")),
            clean(r.get("VEHICLE NUMBER")), clean(r.get("CHASSIS NUMBER")),
            clean(r.get("VEHICLE FINANCE")), clean(r.get("VEHICLE INSURANCE")), 74))
    print(f"  VEHICLE OWNERS: {len(rows)-base:>7,}")

    base = len(rows)
    df = pd.read_excel(DATA, sheet_name="HDFC CUSTOMERS")
    for _, r in df.iterrows():
        rows.append(("hdfc", clean(r.get("FULL NAME / ENTITY")), clean(r.get("MOBILE (PRIMARY)")),
            clean(r.get("MOBILE (ALTERNATE)")), None, clean(r.get("COMPANY / ORGANISATION")), None,
            None, None, None, None, "Telangana",
            None, None, None, None, None, None, None, None,
            None, None, None, None, None, None, 87))
    print(f"  HDFC CUSTOMERS: {len(rows)-base:>7,}")

    INSERT_SQL = """
        INSERT INTO contacts
        (segment,name,mobile,mobile_alt,email,company,income,
         address,village,mandal,district,state,pincode,gstin,
         gst_comm,gst_div,gst_range,school_code,school_cat,school_mgmt,
         car_brand,car_variant,vehicle_no,chassis_no,veh_finance,veh_insure,score)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """
    BATCH = 5000
    with get_db() as conn:
        for i in range(0, len(rows), BATCH):
            conn.executemany(INSERT_SQL, rows[i:i+BATCH])
            conn.commit()
            print(f"  Inserted {min(i+BATCH,len(rows)):>7,} / {len(rows):,}", end="\r")
    print(f"\n[LOAD] Done — {len(rows):,} records loaded.")

# ─── Routes ───────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory(FRONT, "index.html")

@app.route("/api/stats")
def stats():
    segs  = query("SELECT segment, COUNT(*) as cnt FROM contacts GROUP BY segment")
    pipe  = query("SELECT stage,   COUNT(*) as cnt FROM pipeline  GROUP BY stage")
    total = query("SELECT COUNT(*) as n FROM contacts", one=True)["n"]
    return jsonify({"total": total,
                    "segments": {r["segment"]: r["cnt"] for r in segs},
                    "pipeline": {r["stage"]:   r["cnt"] for r in pipe}})

@app.route("/api/contacts")
def contacts():
    segment  = request.args.get("segment",  "all")
    district = request.args.get("district", "")
    q        = request.args.get("q",        "").strip()
    page     = max(1,   int(request.args.get("page",     1)))
    per_page = min(200, int(request.args.get("per_page", 48)))
    sort     = request.args.get("sort", "score")
    where, params = ["1=1"], []
    if segment != "all": where.append("segment = ?"); params.append(segment)
    if district:         where.append("district LIKE ?"); params.append(f"%{district}%")
    if q:
        where.append("(name LIKE ? OR mobile LIKE ? OR mobile_alt LIKE ? OR company LIKE ? OR district LIKE ? OR gstin LIKE ? OR email LIKE ? OR car_brand LIKE ?)")
        s = f"%{q}%"; params.extend([s]*8)
    clause = " AND ".join(where)
    total  = query(f"SELECT COUNT(*) as n FROM contacts WHERE {clause}", params, one=True)["n"]
    sort   = sort if sort in {"score","name","income","created_at"} else "score"
    rows   = query(
        f"SELECT id,segment,name,mobile,mobile_alt,company,income,district,state,"
        f"gstin,car_brand,car_variant,email,score,stage,notes,created_at "
        f"FROM contacts WHERE {clause} ORDER BY {sort} DESC LIMIT ? OFFSET ?",
        params + [per_page, (page-1)*per_page])
    return jsonify({"total": total, "page": page, "per_page": per_page,
                    "pages": max(1, math.ceil(total/per_page)),
                    "contacts": [dict(r) for r in rows]})

@app.route("/api/contacts/<int:cid>")
def contact_detail(cid):
    row = query("SELECT * FROM contacts WHERE id = ?", (cid,), one=True)
    if not row: return jsonify({"error": "Not found"}), 404
    acts = query("SELECT * FROM activities WHERE contact_id=? ORDER BY created_at DESC LIMIT 20", (cid,))
    d = dict(row); d["activities"] = [dict(a) for a in acts]
    return jsonify(d)

@app.route("/api/contacts", methods=["POST"])
def add_contact():
    d = request.json or {}
    cid = execute("""INSERT INTO contacts
        (segment,name,mobile,mobile_alt,email,company,income,address,district,
         state,pincode,gstin,car_brand,car_variant,vehicle_no,score,stage,notes)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (d.get("segment","it"), d.get("name"), d.get("mobile"), d.get("mobile_alt"),
         d.get("email"), d.get("company"), d.get("income"), d.get("address"),
         d.get("district"), d.get("state","Telangana"), d.get("pincode"),
         d.get("gstin"), d.get("car_brand"), d.get("car_variant"),
         d.get("vehicle_no"), d.get("score",50), d.get("stage","New Lead"), d.get("notes")))
    execute("INSERT INTO activities (contact_id,type,note) VALUES (?,?,?)", (cid,"created","Contact created"))
    return jsonify({"id": cid, "message": "Contact created"}), 201

@app.route("/api/contacts/<int:cid>", methods=["PUT"])
def update_contact(cid):
    d = request.json or {}
    execute("UPDATE contacts SET stage=?,score=?,notes=?,company=?,income=?,district=?,name=?,mobile=? WHERE id=?",
            (d.get("stage"), d.get("score"), d.get("notes"), d.get("company"),
             d.get("income"), d.get("district"), d.get("name"), d.get("mobile"), cid))
    execute("INSERT INTO activities (contact_id,type,note) VALUES (?,?,?)", (cid,"updated",d.get("notes","Updated")))
    return jsonify({"message": "Updated"})

@app.route("/api/contacts/<int:cid>", methods=["DELETE"])
def delete_contact(cid):
    execute("DELETE FROM contacts WHERE id=?", (cid,))
    return jsonify({"message": "Deleted"})

@app.route("/api/pipeline")
def get_pipeline():
    rows = query("""SELECT p.id,p.name,p.segment,p.stage,p.value,p.notes,p.created_at,
                           c.mobile,c.district,c.company,c.car_brand,c.income
                    FROM pipeline p LEFT JOIN contacts c ON p.contact_id=c.id
                    ORDER BY p.created_at DESC""")
    stages = ["New Lead","Contacted","Site Visit","Negotiation","Closed"]
    result = {s: [] for s in stages}
    for r in rows:
        d = dict(r); result.setdefault(d.get("stage","New Lead"),[]).append(d)
    return jsonify(result)

@app.route("/api/pipeline", methods=["POST"])
def add_pipeline():
    d = request.json or {}
    pid = execute("INSERT INTO pipeline (contact_id,name,segment,stage,value,notes) VALUES (?,?,?,?,?,?)",
                  (d.get("contact_id"),d.get("name"),d.get("segment","it"),
                   d.get("stage","New Lead"),d.get("value"),d.get("notes")))
    return jsonify({"id": pid}), 201

@app.route("/api/pipeline/<int:pid>", methods=["PUT"])
def update_pipeline(pid):
    d = request.json or {}
    execute("UPDATE pipeline SET stage=?,value=?,notes=? WHERE id=?",
            (d.get("stage"),d.get("value"),d.get("notes"),pid))
    return jsonify({"message": "Updated"})

@app.route("/api/pipeline/<int:pid>", methods=["DELETE"])
def delete_pipeline(pid):
    execute("DELETE FROM pipeline WHERE id=?", (pid,))
    return jsonify({"message": "Deleted"})

@app.route("/api/districts")
def districts():
    rows = query("SELECT DISTINCT district FROM contacts WHERE district IS NOT NULL ORDER BY district")
    return jsonify([r["district"] for r in rows if r["district"]])

@app.route("/api/analytics")
def analytics():
    top_cos  = query("SELECT company, COUNT(*) cnt FROM contacts WHERE company IS NOT NULL AND segment='it' GROUP BY company ORDER BY cnt DESC LIMIT 12")
    top_cars = query("SELECT car_brand, COUNT(*) cnt FROM contacts WHERE car_brand IS NOT NULL GROUP BY car_brand ORDER BY cnt DESC LIMIT 10")
    inc_dist = query("""SELECT CASE WHEN income<500000 THEN '< 5L' WHEN income<1500000 THEN '5L - 15L' WHEN income<3000000 THEN '15L - 30L' ELSE '30L+' END as bracket, COUNT(*) cnt FROM contacts WHERE income IS NOT NULL GROUP BY bracket""")
    gst_comm = query("SELECT gst_comm, COUNT(*) cnt FROM contacts WHERE gst_comm IS NOT NULL GROUP BY gst_comm ORDER BY cnt DESC LIMIT 10")
    top_dist = query("SELECT district, COUNT(*) cnt FROM contacts WHERE district IS NOT NULL GROUP BY district ORDER BY cnt DESC LIMIT 12")
    return jsonify({"top_companies":[dict(r) for r in top_cos],"top_cars":[dict(r) for r in top_cars],
                    "income_distribution":[dict(r) for r in inc_dist],"gst_commissionerate":[dict(r) for r in gst_comm],
                    "top_districts":[dict(r) for r in top_dist]})

@app.route("/api/search")
def search():
    q = request.args.get("q","").strip()
    if not q or len(q)<2: return jsonify([])
    s = f"%{q}%"
    rows = query("SELECT id,segment,name,mobile,company,district,score FROM contacts WHERE name LIKE ? OR mobile LIKE ? OR company LIKE ? OR gstin LIKE ? ORDER BY score DESC LIMIT 20",(s,s,s,s))
    return jsonify([dict(r) for r in rows])

# ─── Entry ────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("\n╔══════════════════════════════════════╗")
    print("║     PropTrack CRM — Backend          ║")
    print("╚══════════════════════════════════════╝\n")
    init_db()
    load_excel()
    port = int(os.environ.get("PORT", 5000))
    print(f"\n🚀  Running at → http://localhost:{port}\n")
    app.run(debug=False, host="0.0.0.0", port=port, threaded=True)
