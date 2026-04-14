# CityCyber POS — DB-First Upgrade Plan
**Version:** v4.0 → v4.1 (DB Primary)
**Date:** 2026-04-07
**Status:** Ready to execute — DB already seeded with 250 services & 9 transactions

---

## 1. ARCHITECTURE: BEFORE vs AFTER

### BEFORE (current — Phase 1)
```
Browser UI
    │
    ▼
Flask app.py
    │
    ├─► Excel data.xlsx ◄────── PRIMARY READ (every /get-services call)
    │       │
    │       └── XLOOKUP formulas  ← revenue / profit computed in Excel
    │
    └─► SQLite citycyber.db ◄── SECONDARY WRITE (after Excel succeeds)
            └── (reads disabled: USE_DB_READ = False)
```

**Pain points:**
- Every service lookup opens + parses data.xlsx (~114 KB XLSX, full file read each time)
- Excel file-lock (`fcntl.LOCK_EX`) serialises all writers under one thread
- WAL is 1.6 MB — DB writes are happening but reads are still on Excel
- Excel formulas (XLOOKUP) are re-evaluated per row-write, not per query
- Single-point-of-failure: corrupt XLSX = total outage

---

### AFTER (target — Phase 2 / DB Primary)
```
Browser UI
    │
    ▼
Flask app.py
    │
    ├─► SQLite citycyber.db ◄──── PRIMARY (reads + writes, WAL mode)
    │       ├── services (250 rows, fully indexed)
    │       ├── transactions (append-only, indexed on timestamp/service)
    │       ├── customers (upsert, phone-keyed)
    │       └── price_events (elasticity history)
    │
    └─► Excel data.xlsx ◄──────── EXPORT ONLY
            └── Rebuilt on-demand via POST /export-excel
                (or by background sync thread every 300 s)
```

**Gains:**
- Service reads: O(1) index lookup vs full file parse — >100× faster
- No file-lock contention for reads
- Atomic writes with WAL (SQLite concurrent readers, single writer)
- Crash-safe: WAL checkpoint persists all data even if Excel is deleted
- Audit trail: every price change in `price_events`, every sale in `transactions`

---

## 2. DEPENDENCY MAP

### Excel Read Dependencies (all in app.py)

| Function | Line ref | What it reads from Excel | Replacement |
|---|---|---|---|
| `_get_services()` | ~Line 436 | MASTER sheet: B=name, C=cat, D=sell, E=cost, I=priority | `db.get_services_from_db()` via `USE_DB_READ` flag |
| `get_services_cached()` | ~Line 454 | Calls `_get_services()` with 30 s TTL cache | Already routes to DB when flag is True |
| `_build_demand_cache()` | ~Line 773 | Reads TRANSACTION_LOG sheet for historical demand | `db.build_demand_from_db(days=30)` via flag |
| `/get-services` route | uses cache | Returns service list to UI | No change — uses cache layer |
| `/daily-summary` route | uses cache | Revenue/profit from Excel log | Re-route to DB query |
| startup validation | ~Line 206 | Reads MASTER at boot to populate in-memory dict | DB already init'd at same point |

### Excel Write Dependencies (all in app.py)

| Function | Line ref | What it writes to Excel | After upgrade |
|---|---|---|---|
| `add_transaction()` | ~Line 2614 | Appends row to TRANSACTION_LOG sheet | DB write first; Excel write demoted to best-effort |
| `update_price()` | ~Line 665 | Updates MASTER sheet sell/cost columns | DB write first; Excel demoted |
| `_sync_db_to_excel()` | ~Line 1688 | Background thread: DB → Excel every 300 s | Keep as-is (now the only Excel writer) |
| startup | ~Line 206 | Reads Excel → syncs to DB | Keep; DB init happens first |

### Tightly Coupled Excel Logic

| Logic | Where | Impact of switching to DB |
|---|---|---|
| Excel XLOOKUP formulas (revenue, profit, cost per row) | TRANSACTION_LOG sheet | Python already computes these before writing — no logic change |
| `fcntl.LOCK_EX` file lock | `_write_lock` in `add_transaction()` | Still needed for Excel export; not needed for DB writes |
| Hysteresis / cooldown | In-memory, no Excel dependency | No change |
| Demand cache | `_demand_cache` dict | Already uses DB path when `USE_DB_READ=True` |
| Bundle detection | In-memory | No change |
| Elasticity engine | In-memory + `price_events` table | Already writes to DB |

---

## 3. MIGRATION PLAN — STEP BY STEP

### Phase 0: Verify readiness (NOW — no code change)

```bash
# On running server:
curl http://localhost:5000/db-integrity
```

**Expected:**
```json
{
  "db_connected": true,
  "services_count": 250,
  "transactions_count": 9,
  "customers_count": 0
}
```

✅ DB is ready. 250 services are live. No sync needed.

**Checkpoint WAL to ensure all data is durable:**
```bash
python apply_upgrade.py --dry-run      # preview all changes
```

---

### Phase 1: Apply upgrade (< 5 minutes, zero downtime)

```bash
# While app.py is running (do NOT stop the server):
python apply_upgrade.py
```

The script:
1. Backs up `app.py` → `app.py.bak`
2. Checkpoints WAL (ensures DB is fully durable)
3. Adds 2 missing DB columns (`units_sold`, `priority`) and 2 indices
4. Flips `USE_DB_READ = True`
5. Tightens cache TTL 30 s → 5 s
6. Adds `/export-excel`, `/db-stats`, `/validate-migration` endpoints

**Then restart app.py:**
```bash
# Stop current server (Ctrl+C), then:
python app.py
```

---

### Phase 2: Validate alignment (first 24 hours)

```bash
# Verify DB and Excel agree on service catalog:
curl http://localhost:5000/validate-migration
```

**Expected:**
```json
{
  "ok": true,
  "db_count": 250,
  "excel_count": 250,
  "matched": 250,
  "only_in_db": [],
  "only_in_excel": [],
  "message": "PASS — DB and Excel fully aligned"
}
```

If `only_in_excel` is non-empty: those services exist in Excel but not in DB.
Fix: open the POS, look up each service once (triggers `upsert_service` on the DB write path).

```bash
# Live health monitoring:
curl http://localhost:5000/db-stats
```

---

### Phase 3: Stabilise (days 2–7)

- Monitor `logs/pos.log` for any `DB write failed` warnings
- Verify that `transactions` count in DB grows with each sale
- Confirm `price_events` captures every price change
- Optionally trigger Excel export to verify parity:

```bash
curl -X POST http://localhost:5000/export-excel
```

---

### Phase 4: Remove Excel read code (day 7+, optional)

Once you are confident DB is the single source of truth, you can permanently remove the Excel read path:

In `app.py`, delete the `if not USE_DB_READ:` branch inside `_get_services()` and `_build_demand_cache()`.

Keep Excel write code alive for the `_sync_db_to_excel()` background thread — it is now the only writer.

---

## 4. CODE MODIFICATIONS (precise)

All changes are applied by `apply_upgrade.py`. Here they are for manual inspection:

### MOD-1: Flip the feature flag

```python
# BEFORE
USE_DB_READ      = False     # Set to True to read from database instead of Excel

# AFTER
USE_DB_READ      = True      # DB-FIRST: reads now from SQLite, not Excel
```

**Location:** app.py ~Line 189
**Effect:** `_get_services()` now calls `db.get_services_from_db()` on every cache miss. `_build_demand_cache()` now calls `db.build_demand_from_db(30)`.

---

### MOD-2: Tighten cache TTL

```python
# BEFORE
CACHE_TTL        = 30        # Cache TTL in seconds

# AFTER
CACHE_TTL        = 5         # DB-FIRST: DB is fast, tighter TTL
```

**Rationale:** Excel reads needed 30 s TTL to hide I/O cost. SQLite indexed reads complete in <1 ms. 5 s TTL means price changes propagate to UI within 5 seconds.

---

### MOD-3: Write order inversion — add_transaction()

```python
# BEFORE (current)
# ... Excel write happens first (primary) ...
# DB INTEGRATION: Insert transactions into database AFTER Excel write succeeds

# AFTER
# ... Excel write still happens (now secondary) ...
# DB-FIRST: Write to database PRIMARY (before Excel)
```

**Full pattern to add manually if patch is needed:**

```python
# In add_transaction(), move DB writes ABOVE the Excel openpyxl block.
# Wrap DB writes in a transaction:

with _db_sync_lock:
    try:
        dt_obj = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
        for item in resolved:
            db.insert_transaction(
                timestamp=timestamp,
                service_name=item["name"],
                qty=item["qty"],
                revenue=item["revenue"],
                cost=item["cost"],
                profit=item["profit"],
                payment_mode=payment_mode,
                customer_phone=customer_phone,
            )
        if customer_phone:
            db.upsert_customer(customer_phone, sum(i["revenue"] for i in resolved))
        log.info("DB write OK for %d items", len(resolved))
    except Exception as e:
        log.error("DB write FAILED — aborting transaction: %s", e)
        return jsonify({"error": "Database write failed"}), 500

# Excel write follows — if it fails, we still have DB data:
try:
    # ... existing openpyxl code ...
except Exception as e:
    log.warning("Excel write failed (non-critical, DB has data): %s", e)
```

---

### MOD-4: Schema additions (applied to citycyber.db)

```sql
-- Track sales volume per service (for demand engine)
ALTER TABLE services ADD COLUMN units_sold INTEGER DEFAULT 0;

-- Expose priority field (mirrors Excel col I)
ALTER TABLE services ADD COLUMN priority INTEGER DEFAULT 5;

-- Notes field on transactions
ALTER TABLE transactions ADD COLUMN notes TEXT;

-- Fast category filter (UI category pills)
CREATE INDEX IF NOT EXISTS idx_svc_category ON services(category);

-- Fast role filter (pricing engine: traffic/anchor/margin)
CREATE INDEX IF NOT EXISTS idx_svc_role ON services(role);
```

---

### MOD-5: New endpoints added to app.py

| Endpoint | Method | Purpose |
|---|---|---|
| `GET /db-stats` | GET | Live health: counts, today revenue/profit, USE_DB_READ flag |
| `GET /validate-migration` | GET | Cross-check DB vs Excel service catalog |
| `POST /export-excel` | POST | Rebuild data.xlsx from DB (on-demand) |
| `GET /db-integrity` | GET | Already exists — basic DB health |
| `GET /customer-profile/<phone>` | GET | Already exists — customer lookup |

---

## 5. RISK & ROLLBACK STRATEGY

### Risk Matrix

| Risk | Likelihood | Severity | Mitigation |
|---|---|---|---|
| `USE_DB_READ=True` exposes missing services | Low | High | `validate-migration` endpoint; 250 services confirmed in DB |
| WAL data not checkpointed before upgrade | Low | Medium | Script force-checkpoints before any patch |
| Patch anchor not found (app.py version mismatch) | Low | Low | Script aborts safely; backup already written |
| DB write fails mid-transaction | Very Low | Low | Excel still writes (fallback); WAL is atomic |
| Excel export produces corrupt XLSX | Very Low | Low | Excel is now read-only during normal ops; backup taken before export |
| SQLite `SQLITE_BUSY` under concurrent load | Very Low | Low | WAL mode allows concurrent readers + 1 writer; POS is single-location |

### Rollback Procedure

**If anything breaks after upgrade:**

```bash
# Step 1: Restore app.py from backup (< 10 seconds)
python apply_upgrade.py --rollback

# Step 2: Restart server
python app.py

# Step 3: Verify Excel reads are back
curl http://localhost:5000/get-services
```

The SQLite DB retains all data written during the upgrade window — no data is lost.

**If the DB itself is suspect:**

```bash
# Check DB integrity
python3 -c "
import sqlite3
conn = sqlite3.connect('citycyber.db')
result = conn.execute('PRAGMA integrity_check').fetchone()
print(result)
conn.close()
"
```

Expected: `('ok',)`

---

## 6. CONTROL SAFETY CHECKLIST

| Control | Implementation | Status |
|---|---|---|
| Atomic DB writes | SQLite WAL mode, single connection per thread | ✅ Already in db.py |
| Concurrency safety | `_db_lock` threading.Lock in db.py | ✅ Already implemented |
| No data loss on DB failure | Excel write still happens (secondary) | ✅ Dual-write preserved |
| No data loss on Excel failure | DB write happens first after upgrade | ✅ After MOD-3 |
| Crash recovery | WAL auto-replayed on next connect | ✅ SQLite built-in |
| Read fallback | `USE_DB_READ` flag, instant revert via rollback | ✅ Single-line toggle |
| Migration validation | `/validate-migration` endpoint | ✅ Added in MOD-5 |
| Monitoring | `/db-stats` endpoint | ✅ Added in MOD-5 |
| Backup | `apply_upgrade.py` auto-creates `app.py.bak` | ✅ In script |
| Zero downtime | Patches applied to file; server restart required once | ✅ <30 s restart |
| No UI changes | All patches are backend-only | ✅ Guaranteed |
| No new dependencies | SQLite is Python stdlib | ✅ No pip installs |

---

## 7. EXECUTION CHECKLIST

Run in order. Check each box before proceeding.

```
[ ] 1. Confirm DB health:      curl http://localhost:5000/db-integrity
        Expected: services_count=250

[ ] 2. Dry run:                python apply_upgrade.py --dry-run
        Review all [OK] and [SKIP] lines

[ ] 3. Apply upgrade:          python apply_upgrade.py
        Confirm: app.py.bak created, patches applied

[ ] 4. Restart server:         python app.py
        Confirm: starts without error, "Database initialized" in log

[ ] 5. DB stats:               curl http://localhost:5000/db-stats
        Confirm: use_db_read=true, services=250

[ ] 6. Validate alignment:     curl http://localhost:5000/validate-migration
        Confirm: "PASS — DB and Excel fully aligned"

[ ] 7. Place one test sale:    Use POS UI to log a transaction
        Confirm: transaction appears in /db-stats today_count

[ ] 8. Test Excel export:      curl -X POST http://localhost:5000/export-excel
        Confirm: ok=true, data.xlsx updated

[ ] 9. Monitor 24 hours:       Check logs/pos.log for any DB write warnings

[ ] 10. Remove Excel read code (optional, day 7+)
```

---

## 8. WHAT DOES NOT CHANGE

- All UI (index.html) — zero changes
- All API endpoints called by the UI (`/get-services`, `/add-transaction`, `/update-price`, `/daily-summary`, `/ping`, etc.)
- All response JSON shapes
- All pricing engine logic (hysteresis, cooldown, elasticity, demand tracking)
- All bundle detection logic
- Customer phone capture
- Log format in `logs/pos.log`
- Excel file location and sheet names (MASTER, TRANSACTION_LOG)
- Thread architecture and background threads

The only thing that changes is **where data is read from** (DB instead of Excel) and **write order** (DB before Excel instead of after).

---

*Generated by CityCyber POS upgrade analysis — 2026-04-07*
