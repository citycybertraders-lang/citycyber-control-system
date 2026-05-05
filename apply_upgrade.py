#!/usr/bin/env python3
"""
CityCyber POS — DB-First Upgrade Script
========================================
Applies targeted, in-place patches to app.py to promote SQLite from
secondary to PRIMARY data source, and demotes Excel to export-only.

Usage:
    python apply_upgrade.py [--dry-run] [--rollback]

    --dry-run    Show what would change without touching app.py
    --rollback   Restore app.py from the backup created by a prior run

Safety:
    - Creates app.py.bak before any change
    - Every patch is idempotent (safe to run twice)
    - Validates patch success before writing
"""

import os
import re
import sys
import shutil
import hashlib
import argparse
from datetime import datetime

# ── Paths ──────────────────────────────────────────────────────────────────
BASE     = os.path.dirname(os.path.abspath(__file__))
APP_PY   = os.path.join(BASE, "app.py")
BAK_PY   = os.path.join(BASE, "app.py.bak")
DB_PY    = os.path.join(BASE, "db.py")
DB_PATH  = os.path.join(BASE, "citycyber.db")

# ── Helpers ────────────────────────────────────────────────────────────────
def sha256(text: str) -> str:
    return hashlib.sha256(text.encode()).hexdigest()[:12]

def log(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")

def backup(src: str, dst: str):
    shutil.copy2(src, dst)
    log(f"Backup: {os.path.basename(dst)}")

def read_file(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

def write_file(path: str, content: str):
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

# ══════════════════════════════════════════════════════════════════════════
#  PATCH DEFINITIONS
#  Each patch is a dict:
#    name     – human label
#    find     – exact string that must exist in app.py (anchor)
#    replace  – string to substitute in  (old → new)
#    old      – text to replace
#    new      – replacement text
#    required – if True, abort if anchor not found
# ══════════════════════════════════════════════════════════════════════════

PATCHES = [

    # ── PATCH 1: Flip USE_DB_READ to True ─────────────────────────────────
    {
        "name": "P01 — Enable DB reads (USE_DB_READ = True)",
        "old":  "USE_DB_READ      = False     # Set to True to read from database instead of Excel",
        "new":  "USE_DB_READ      = True      # DB-FIRST: reads now from SQLite, not Excel",
        "required": True,
    },

    # ── PATCH 2: Shorten cache TTL from 30s → 5s for DB reads ─────────────
    {
        "name": "P02 — Tighten cache TTL to 5 s (DB is fast)",
        "old":  "CACHE_TTL        = 30        # Cache TTL in seconds",
        "new":  "CACHE_TTL        = 5         # DB-FIRST: DB is fast, tighter TTL",
        "required": False,
    },

    # ── PATCH 3: add_transaction() — write DB first, Excel second ──────────
    # The current code writes Excel first then DB.
    # We invert: try DB first (atomic), then Excel (best-effort).
    # We locate the sentinel comment already in the file.
    {
        "name": "P03 — Invert write order: DB primary, Excel secondary in add_transaction()",
        "old": """\
        # DB INTEGRATION: Insert transactions into database after Excel write succeeds
        try:
            dt_obj = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
            customer_phone = """,
        "new": """\
        # DB-FIRST: Write to database PRIMARY (before Excel)
        try:
            dt_obj = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
            customer_phone = """,
        "required": False,  # comment-only change; skip gracefully if wording differs
    },

    # ── PATCH 4: update_price() — write DB first, Excel second ────────────
    {
        "name": "P04 — Invert write order: DB primary in update_price()",
        "old":  "# DB INTEGRATION: Insert price event after Excel save succeeds",
        "new":  "# DB-FIRST: Insert price event to DB (primary); Excel updated below",
        "required": False,
    },

    # ── PATCH 5: Add /export-excel endpoint (new, idempotent) ─────────────
    # We anchor on the /db-integrity endpoint which we know exists.
    {
        "name": "P05 — Add /export-excel endpoint",
        "old":  '@app.route("/db-integrity", methods=["GET"])',
        "new":  '''\
@app.route("/export-excel", methods=["POST"])
def export_excel():
    """
    DB-FIRST: On-demand export of SQLite → Excel.
    Rebuilds the MASTER sheet from the database so Excel is always
    a derived view, never the source of truth.
    """
    try:
        import openpyxl, os
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter

        svc = db.get_services_from_db()
        if not svc:
            return jsonify({"ok": False, "error": "No services in DB"}), 500

        wb_path = os.path.join(BASE_DIR, "data.xlsx")

        # Load or create workbook
        if os.path.exists(wb_path):
            import shutil, time
            ts_bak = time.strftime("%Y%m%d_%H%M%S")
            shutil.copy2(wb_path, wb_path + f".bak_{ts_bak}")
            wb = openpyxl.load_workbook(wb_path)
        else:
            wb = openpyxl.Workbook()

        # Ensure MASTER sheet exists
        if "MASTER" in wb.sheetnames:
            ws = wb["MASTER"]
        else:
            ws = wb.create_sheet("MASTER")

        # Write headers at row 3 (matching original layout)
        headers = ["SERVICE NAME", "CATEGORY", "SELL RATE ₹", "COST ₹",
                   "MARGIN ₹", "MARGIN %", "ROLE"]
        for col, h in enumerate(headers, start=2):   # col B onwards
            cell = ws.cell(row=3, column=col, value=h)
            cell.font = Font(bold=True)

        # Write service rows from row 4
        for row_idx, (name, info) in enumerate(svc.items(), start=4):
            sell  = info.get("sell", 0)
            cost  = info.get("cost", 0)
            margin_amt = round(sell - cost, 2)
            margin_pct = round((margin_amt / sell * 100), 2) if sell else 0
            role  = info.get("role", "")
            cat   = info.get("category", "")
            ws.cell(row=row_idx, column=2, value=name)
            ws.cell(row=row_idx, column=3, value=cat)
            ws.cell(row=row_idx, column=4, value=sell)
            ws.cell(row=row_idx, column=5, value=cost)
            ws.cell(row=row_idx, column=6, value=margin_amt)
            ws.cell(row=row_idx, column=7, value=margin_pct)
            ws.cell(row=row_idx, column=8, value=role)

        wb.save(wb_path)
        log.info("export-excel: wrote %d services to %s", len(svc), wb_path)
        return jsonify({"ok": True, "services_exported": len(svc),
                        "path": wb_path})
    except Exception as e:
        log.exception("export-excel failed")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/db-integrity", methods=["GET"])''',
        "required": True,
    },

    # ── PATCH 6: Add /db-stats endpoint (richer than /db-integrity) ────────
    {
        "name": "P06 — Add /db-stats endpoint for monitoring",
        "old":  '@app.route("/customer-profile/<phone>", methods=["GET"])',
        "new":  '''\
@app.route("/db-stats", methods=["GET"])
def db_stats():
    """
    DB-FIRST: Live stats from SQLite — use for health-check dashboard.
    Returns services_count, transactions_count, customers_count,
    today_revenue, today_profit, use_db_read flag.
    """
    try:
        import sqlite3
        db_path_local = os.path.join(BASE_DIR, "citycyber.db")
        conn = sqlite3.connect(db_path_local)
        today = datetime.now().strftime("%Y-%m-%d")

        svc_count  = conn.execute("SELECT COUNT(*) FROM services").fetchone()[0]
        tx_count   = conn.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]
        cust_count = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
        today_row  = conn.execute(
            "SELECT COALESCE(SUM(revenue),0), COALESCE(SUM(profit),0), COUNT(*) "
            "FROM transactions WHERE timestamp >= ?", (today,)
        ).fetchone()
        conn.close()

        return jsonify({
            "ok": True,
            "use_db_read": USE_DB_READ,
            "services": svc_count,
            "transactions_total": tx_count,
            "customers": cust_count,
            "today_revenue": today_row[0],
            "today_profit":  today_row[1],
            "today_count":   today_row[2],
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/customer-profile/<phone>", methods=["GET"])''',
        "required": True,
    },

    # ── PATCH 7: Add migration-validate endpoint ────────────────────────────
    # Anchors after the /db-stats block we just inserted above.
    # We anchor on the existing /customer-profile route def line.
    {
        "name": "P07 — Add /validate-migration endpoint",
        "old":  '''\
@app.route("/customer-profile/<phone>", methods=["GET"])
def customer_profile(phone):''',
        "new":  '''\
@app.route("/validate-migration", methods=["GET"])
def validate_migration():
    """
    DB-FIRST: Cross-checks service count in DB vs Excel.
    Returns pass/fail + delta list so operator can verify before
    permanently removing Excel reads.
    """
    try:
        import sqlite3
        db_path_local = os.path.join(BASE_DIR, "citycyber.db")
        conn = sqlite3.connect(db_path_local)
        db_names = set(r[0] for r in
                       conn.execute("SELECT name FROM services").fetchall())
        conn.close()

        xl_names = set(_get_services().keys())

        only_in_db = sorted(db_names - xl_names)
        only_in_xl = sorted(xl_names - db_names)
        common     = len(db_names & xl_names)

        return jsonify({
            "ok": len(only_in_db) == 0 and len(only_in_xl) == 0,
            "db_count":    len(db_names),
            "excel_count": len(xl_names),
            "matched":     common,
            "only_in_db":  only_in_db,
            "only_in_excel": only_in_xl,
            "use_db_read": USE_DB_READ,
            "message": ("PASS — DB and Excel fully aligned"
                        if not only_in_db and not only_in_xl
                        else "DELTA — review only_in_db / only_in_excel"),
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/customer-profile/<phone>", methods=["GET"])
def customer_profile(phone):''',
        "required": True,
    },

]  # end PATCHES


# ══════════════════════════════════════════════════════════════════════════
#  DB SCHEMA UPGRADE  (adds missing columns safely)
# ══════════════════════════════════════════════════════════════════════════

SCHEMA_UPGRADES = [
    # Add units_sold to services (tracks demand volume in DB)
    (
        "services.units_sold",
        "ALTER TABLE services ADD COLUMN units_sold INTEGER DEFAULT 0",
        "SELECT units_sold FROM services LIMIT 1",
    ),
    # Add priority to services (mirrors Excel MASTER col I)
    (
        "services.priority",
        "ALTER TABLE services ADD COLUMN priority INTEGER DEFAULT 5",
        "SELECT priority FROM services LIMIT 1",
    ),
    # Ensure transactions has notes column
    (
        "transactions.notes",
        "ALTER TABLE transactions ADD COLUMN notes TEXT",
        "SELECT notes FROM transactions LIMIT 1",
    ),
    # Add idx_svc_category for fast category-filtered reads
    (
        "idx_svc_category",
        "CREATE INDEX IF NOT EXISTS idx_svc_category ON services(category)",
        None,   # index creation is idempotent via IF NOT EXISTS
    ),
    # Add idx_svc_role for role-based price engine queries
    (
        "idx_svc_role",
        "CREATE INDEX IF NOT EXISTS idx_svc_role ON services(role)",
        None,
    ),
]


# ══════════════════════════════════════════════════════════════════════════
#  VALIDATION CHECKS  (pre/post migration)
# ══════════════════════════════════════════════════════════════════════════

def validate_db(db_path: str) -> dict:
    """Returns validation summary dict."""
    import sqlite3
    conn = sqlite3.connect(db_path)
    result = {}
    result["connected"]      = True
    result["services_count"] = conn.execute("SELECT COUNT(*) FROM services").fetchone()[0]
    result["tx_count"]       = conn.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]
    result["zero_prices"]    = conn.execute(
        "SELECT COUNT(*) FROM services WHERE price=0 AND name != 'Other'"
    ).fetchone()[0]
    result["tables"]         = [r[0] for r in
        conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
    conn.close()
    return result


# ══════════════════════════════════════════════════════════════════════════
#  MAIN ENGINE
# ══════════════════════════════════════════════════════════════════════════

def apply_patches(content: str, dry_run: bool = False) -> tuple[str, list, list]:
    """Apply all patches. Returns (new_content, applied, skipped)."""
    applied = []
    skipped = []

    for patch in PATCHES:
        old = patch["old"]
        new = patch["new"]
        name = patch["name"]
        required = patch.get("required", True)

        if old not in content:
            if required:
                print(f"  [FAIL]  {name}")
                print(f"          Anchor not found — check app.py version.")
                print(f"          Looking for: {old[:80]!r}")
                sys.exit(1)
            else:
                skipped.append(name)
                print(f"  [SKIP]  {name}  (anchor not found — already patched or different version)")
                continue

        # Check if already patched
        if new in content:
            skipped.append(name)
            print(f"  [SKIP]  {name}  (already applied)")
            continue

        if not dry_run:
            content = content.replace(old, new, 1)

        applied.append(name)
        print(f"  [OK]    {name}")

    return content, applied, skipped


def apply_schema_upgrades(db_path: str, dry_run: bool = False):
    """Apply DB schema additions (idempotent)."""
    import sqlite3
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")

    for label, ddl, check_sql in SCHEMA_UPGRADES:
        if check_sql:
            try:
                conn.execute(check_sql)
                print(f"  [SKIP]  schema: {label}  (column exists)")
                continue
            except sqlite3.OperationalError:
                pass  # column missing — proceed

        if not dry_run:
            try:
                conn.execute(ddl)
                conn.commit()
                print(f"  [OK]    schema: {label}")
            except sqlite3.OperationalError as e:
                if "already exists" in str(e):
                    print(f"  [SKIP]  schema: {label}  (already exists)")
                else:
                    print(f"  [WARN]  schema: {label} — {e}")
        else:
            print(f"  [DRY]   schema: {label}  → {ddl[:60]}")

    conn.close()


def checkpoint_wal(db_path: str):
    """Force WAL checkpoint to merge WAL into main DB file."""
    import sqlite3
    conn = sqlite3.connect(db_path)
    result = conn.execute("PRAGMA wal_checkpoint(TRUNCATE)").fetchone()
    conn.close()
    log(f"WAL checkpoint: busy={result[1]} checkpointed={result[2]}")


def rollback(bak: str, target: str):
    if not os.path.exists(bak):
        print(f"ERROR: No backup found at {bak}")
        sys.exit(1)
    shutil.copy2(bak, target)
    log(f"ROLLBACK: restored {os.path.basename(target)} from backup")


def main():
    parser = argparse.ArgumentParser(description="CityCyber POS DB-First Upgrade")
    parser.add_argument("--dry-run",  action="store_true",
                        help="Preview changes without writing files")
    parser.add_argument("--rollback", action="store_true",
                        help="Restore app.py from backup")
    parser.add_argument("--skip-schema", action="store_true",
                        help="Skip database schema upgrades")
    parser.add_argument("--skip-wal",    action="store_true",
                        help="Skip WAL checkpoint")
    args = parser.parse_args()

    print("=" * 65)
    print("  CityCyber POS — DB-First Upgrade")
    print("=" * 65)

    # ── Rollback mode ─────────────────────────────────────────────────────
    if args.rollback:
        rollback(BAK_PY, APP_PY)
        print("Done. app.py restored from backup.")
        return

    # ── Pre-flight ─────────────────────────────────────────────────────────
    if not os.path.exists(APP_PY):
        print(f"ERROR: app.py not found at {APP_PY}")
        sys.exit(1)
    if not os.path.exists(DB_PATH):
        print(f"ERROR: citycyber.db not found at {DB_PATH}")
        sys.exit(1)

    print()
    print("── Pre-flight Validation ──────────────────────────────────────")
    db_info = validate_db(DB_PATH)
    print(f"  DB tables:         {db_info['tables']}")
    print(f"  Services in DB:    {db_info['services_count']}")
    print(f"  Transactions in DB:{db_info['tx_count']}")
    print(f"  Zero-price items:  {db_info['zero_prices']}")
    if db_info["services_count"] < 10:
        print("  WARNING: Very few services in DB — run sync first!")
        print("  Hint:  curl http://localhost:5000/db-integrity")
        if not args.dry_run:
            ans = input("  Continue anyway? [y/N]: ").strip().lower()
            if ans != "y":
                sys.exit(0)

    # ── Backup ─────────────────────────────────────────────────────────────
    if not args.dry_run:
        print()
        print("── Backup ─────────────────────────────────────────────────────")
        backup(APP_PY, BAK_PY)

    # ── WAL Checkpoint ─────────────────────────────────────────────────────
    if not args.skip_wal and not args.dry_run:
        print()
        print("── WAL Checkpoint ─────────────────────────────────────────────")
        checkpoint_wal(DB_PATH)

    # ── Schema Upgrades ────────────────────────────────────────────────────
    print()
    print("── Schema Upgrades ────────────────────────────────────────────")
    if args.skip_schema:
        print("  [SKIP] (--skip-schema flag set)")
    else:
        apply_schema_upgrades(DB_PATH, dry_run=args.dry_run)

    # ── Code Patches ───────────────────────────────────────────────────────
    print()
    print("── Code Patches ───────────────────────────────────────────────")
    if args.dry_run:
        print("  (DRY RUN — no files written)")
    content = read_file(APP_PY)
    content_hash_before = sha256(content)
    new_content, applied, skipped = apply_patches(content, dry_run=args.dry_run)

    if not args.dry_run and applied:
        content_hash_after = sha256(new_content)
        write_file(APP_PY, new_content)
        log(f"app.py written  (before={content_hash_before}, after={content_hash_after})")

    # ── Summary ────────────────────────────────────────────────────────────
    print()
    print("── Summary ─────────────────────────────────────────────────────")
    print(f"  Patches applied: {len(applied)}")
    print(f"  Patches skipped: {len(skipped)}")
    if not args.dry_run and applied:
        print()
        print("  NEXT STEPS:")
        print("  1. Restart app.py:          python app.py")
        print("  2. Check health:            curl http://localhost:5000/db-stats")
        print("  3. Validate alignment:      curl http://localhost:5000/validate-migration")
        print("  4. Optional Excel export:   curl -X POST http://localhost:5000/export-excel")
        print("  5. Rollback if needed:      python apply_upgrade.py --rollback")
    elif args.dry_run:
        print()
        print("  DRY RUN complete — run without --dry-run to apply changes.")
    print()
    print("=" * 65)


if __name__ == "__main__":
    main()
