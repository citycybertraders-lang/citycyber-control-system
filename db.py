"""
CityCyber POS — Database Layer  (db.py)
========================================
Phase 1-7 implementation per engineering brief.

Architecture:
  - SQLite with thread-safe connections (check_same_thread=False + WAL mode)
  - All writes parameterised — no string interpolation in SQL
  - Module is purely additive; app.py continues to function without it
  - USE_DB_READ flag in app.py controls whether reads come from DB or Excel

Tables:
  services        — service catalogue mirror
  transactions    — every sale, parallel to TRANSACTION_LOG sheet
  customers       — phone-keyed spend accumulator
  price_events    — full price-change audit trail
"""

import os
import sqlite3
import threading
import logging
import time as _time
from datetime import datetime

log = logging.getLogger("citycyber.db")

# ── Path ──────────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, "citycyber.db")

# ── Thread-local connection storage ───────────────────────────────────────────
_local = threading.local()

# ── Init lock: ensures CREATE TABLE runs exactly once across threads ──────────
_init_lock   = threading.Lock()
_initialized = False

# ── Retry constants ───────────────────────────────────────────────────────────
_RETRY_MAX   = 3
_RETRY_DELAY = 0.05   # 50ms, doubles per attempt


def _retry_execute(conn, sql, params=()):
    """Execute with exponential-backoff retry on SQLITE_BUSY."""
    delay = _RETRY_DELAY
    for attempt in range(_RETRY_MAX):
        try:
            return conn.execute(sql, params)
        except sqlite3.OperationalError as e:
            if 'locked' in str(e).lower() and attempt < _RETRY_MAX - 1:
                log.warning('DB busy (attempt %d/%d) — retrying in %.0fms',
                            attempt + 1, _RETRY_MAX, delay * 1000)
                _time.sleep(delay)
                delay *= 2
            else:
                raise


# ═══════════════════════════════════════════════════════════════════════════════
# CONNECTION MANAGEMENT
# ═══════════════════════════════════════════════════════════════════════════════

def get_db() -> sqlite3.Connection:
    """
    Return the thread-local SQLite connection.
    Creates one if this thread hasn't connected yet.
    WAL mode: readers don't block writers and vice-versa.
    """
    conn = getattr(_local, "conn", None)
    if conn is None:
        conn = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=15)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute("PRAGMA synchronous=NORMAL")   # safe + faster than FULL
        conn.execute("PRAGMA cache_size=-8000")      # 8 MB page cache
        conn.execute("PRAGMA temp_store=MEMORY")
        _local.conn = conn
    return conn


def init_db():
    """
    Create all tables if they don't exist.
    Safe to call multiple times; idempotent via CREATE TABLE IF NOT EXISTS.
    Called once at app startup.
    """
    global _initialized
    with _init_lock:
        if _initialized:
            return
        conn = get_db()

        # Checkpoint any WAL left from a previous crash before DDL
        try:
            conn.execute("PRAGMA wal_checkpoint(PASSIVE)")
        except Exception as _wcp:
            log.warning("WAL checkpoint on init failed (non-fatal): %s", _wcp)

        conn.executescript("""
            -- ── Services catalogue ────────────────────────────────────────────
            CREATE TABLE IF NOT EXISTS services (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                name        TEXT    NOT NULL UNIQUE,
                category    TEXT    NOT NULL DEFAULT 'Other',
                price       REAL    NOT NULL DEFAULT 0.0  CHECK(price >= 0),
                cost        REAL    NOT NULL DEFAULT 0.0  CHECK(cost  >= 0),
                margin_pct  REAL    NOT NULL DEFAULT 0.0,
                role        TEXT    NOT NULL DEFAULT 'filler',
                updated_at  TEXT    NOT NULL
                                    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S','now'))
            );

            -- ── Transactions (append-only) ─────────────────────────────────────
            CREATE TABLE IF NOT EXISTS transactions (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp      TEXT    NOT NULL,
                service_name   TEXT    NOT NULL,
                qty            INTEGER NOT NULL CHECK(qty > 0),
                revenue        REAL    NOT NULL CHECK(revenue >= 0),
                cost           REAL    NOT NULL CHECK(cost    >= 0),
                profit         REAL    NOT NULL,
                payment_mode   TEXT    NOT NULL DEFAULT 'Cash',
                customer_phone TEXT,
                original_price REAL,
                final_price    REAL    CHECK(final_price IS NULL OR final_price >= 0),
                override_type  TEXT    DEFAULT 'none',
                override_value REAL
            );

            -- ── Customers ─────────────────────────────────────────────────────
            CREATE TABLE IF NOT EXISTS customers (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                phone        TEXT    NOT NULL UNIQUE
                                     CHECK(length(phone) >= 10),
                name         TEXT,
                total_spend  REAL    NOT NULL DEFAULT 0.0 CHECK(total_spend >= 0),
                visit_count  INTEGER NOT NULL DEFAULT 0   CHECK(visit_count >= 0),
                last_seen    TEXT,
                created_at   TEXT    NOT NULL
                                     DEFAULT (strftime('%Y-%m-%dT%H:%M:%S','now'))
            );

            -- ── Price audit trail ──────────────────────────────────────────────
            CREATE TABLE IF NOT EXISTS price_events (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                service_name TEXT    NOT NULL,
                old_price    REAL    NOT NULL CHECK(old_price >= 0),
                new_price    REAL    NOT NULL CHECK(new_price >= 0),
                timestamp    TEXT    NOT NULL,
                source       TEXT    NOT NULL DEFAULT 'manual'
            );

            -- ── Udhaar (Credit) Ledger ─────────────────────────────────────────
            CREATE TABLE IF NOT EXISTS udhaar_ledger (
                id                 INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_phone     TEXT    NOT NULL,
                type               TEXT    NOT NULL CHECK(type IN ('debit','credit')),
                amount             REAL    NOT NULL CHECK(amount > 0),
                note               TEXT,
                reference_txn_id   INTEGER,
                timestamp          TEXT    NOT NULL,
                created_at         TEXT    NOT NULL
                                           DEFAULT (strftime('%Y-%m-%dT%H:%M:%S','now'))
            );

            -- ── Idempotency log (24-hour TTL, prevents double-submit) ──────────
            CREATE TABLE IF NOT EXISTS idempotency_log (
                key        TEXT PRIMARY KEY,
                created_at TEXT NOT NULL
                               DEFAULT (strftime('%Y-%m-%dT%H:%M:%S','now')),
                result     TEXT
            );

            -- ── Indexes ───────────────────────────────────────────────────────
            CREATE INDEX IF NOT EXISTS idx_udhaar_phone    ON udhaar_ledger(customer_phone);
            CREATE INDEX IF NOT EXISTS idx_udhaar_ts       ON udhaar_ledger(timestamp);
            CREATE INDEX IF NOT EXISTS idx_tx_timestamp ON transactions(timestamp);
            CREATE INDEX IF NOT EXISTS idx_tx_service   ON transactions(service_name);
            CREATE INDEX IF NOT EXISTS idx_tx_phone     ON transactions(customer_phone);
            CREATE INDEX IF NOT EXISTS idx_tx_date      ON transactions(substr(timestamp,1,10));
            CREATE INDEX IF NOT EXISTS idx_svc_category ON services(category);
            CREATE INDEX IF NOT EXISTS idx_svc_role     ON services(role);
            CREATE INDEX IF NOT EXISTS idx_pe_service   ON price_events(service_name);
            CREATE INDEX IF NOT EXISTS idx_pe_timestamp ON price_events(timestamp);
            CREATE INDEX IF NOT EXISTS idx_idem_created ON idempotency_log(created_at);

            -- ── System Events (audit trail) ────────────────────────────────────
            CREATE TABLE IF NOT EXISTS system_events (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                event_type  TEXT    NOT NULL,
                entity_type TEXT    NOT NULL,
                entity_id   TEXT    NOT NULL,
                old_value   TEXT,
                new_value   TEXT,
                operator    TEXT    NOT NULL DEFAULT 'system',
                timestamp   TEXT    NOT NULL,
                created_at  TEXT    NOT NULL
                                    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S','now'))
            );
            CREATE INDEX IF NOT EXISTS idx_evt_entity ON system_events(entity_type, entity_id);
            CREATE INDEX IF NOT EXISTS idx_evt_ts     ON system_events(timestamp);
            CREATE INDEX IF NOT EXISTS idx_evt_type   ON system_events(event_type);
        """)
        conn.commit()

        # Safe migrations: add columns missing from older DBs
        _safe_migrations(conn)
        _initialized = True
        log.info("DB initialised at %s (WAL checkpoint done)", DB_PATH)


def _safe_migrations(conn):
    """Idempotent column additions for schema upgrades. Also purges stale idempotency keys."""
    for table, col, col_type in [
        ("transactions", "original_price", "REAL"),
        ("transactions", "final_price",    "REAL"),
        ("transactions", "override_type",  "TEXT DEFAULT 'none'"),
        ("transactions", "override_value", "REAL"),
        ("customers",    "name",           "TEXT"),
        # v6 credit intelligence columns
        ("customers",    "credit_limit",   "REAL DEFAULT 0.0"),
        ("customers",    "risk_level",     "TEXT DEFAULT 'low'"),
        ("customers",    "last_payment_date", "TEXT"),
        ("customers",    "repayment_score", "REAL DEFAULT 1.0"),
        ("udhaar_ledger","is_settled",     "INTEGER DEFAULT 0"),
        ("transactions", "is_collected",   "INTEGER DEFAULT 1"),
    ]:
        try:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type}")
            conn.commit()
            log.info("DB migration: added %s.%s", table, col)
        except Exception:
            pass   # column already exists
    # Purge expired idempotency keys (>24h)
    try:
        conn.execute("DELETE FROM idempotency_log WHERE created_at < datetime('now','-1 day')")
        conn.commit()
    except Exception:
        pass




# ═══════════════════════════════════════════════════════════════════════════════
# IDEMPOTENCY — double-submit prevention
# ═══════════════════════════════════════════════════════════════════════════════

def check_idempotency(key: str) -> str | None:
    """Return cached result JSON if key already processed, else None. Thread-safe."""
    if not key:
        return None
    try:
        row = get_db().execute(
            "SELECT result FROM idempotency_log WHERE key = ?", (key,)
        ).fetchone()
        return row["result"] if row else None
    except Exception as e:
        log.warning("idempotency check failed: %s", e)
        return None


def record_idempotency(key: str, result_json: str):
    """Persist result for a given idempotency key. TTL enforced by _safe_migrations."""
    if not key:
        return
    try:
        conn = get_db()
        conn.execute(
            "INSERT OR IGNORE INTO idempotency_log (key, result) VALUES (?, ?)",
            (key, result_json)
        )
        conn.commit()
    except Exception as e:
        log.warning("idempotency record failed: %s", e)


# ═══════════════════════════════════════════════════════════════════════════════
# TRANSACTION WRITES
# ═══════════════════════════════════════════════════════════════════════════════

def insert_transaction(
    timestamp:      str,
    service_name:   str,
    qty:            int,
    revenue:        float,
    cost:           float,
    profit:         float,
    payment_mode:   str,
    customer_phone: str | None = None,
    original_price: float | None = None,
    final_price:    float | None = None,
    override_type:  str | None = None,
    override_value: float | None = None,
) -> int | None:
    """
    Insert one transaction row into the DB.
    Returns the new row id, or None on failure (never raises — caller must not crash).
    override_type: 'none' | 'manual_price' | 'discount_pct'
    """
    try:
        conn = get_db()
        cur = conn.execute(
            """
            INSERT INTO transactions
                (timestamp, service_name, qty, revenue, cost, profit,
                 payment_mode, customer_phone,
                 original_price, final_price, override_type, override_value)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (timestamp, service_name, qty,
             round(revenue, 4), round(cost, 4), round(profit, 4),
             payment_mode, customer_phone or None,
             round(original_price, 4) if original_price is not None else None,
             round(final_price,    4) if final_price    is not None else None,
             override_type,
             round(override_value, 4) if override_value is not None else None),
        )
        conn.commit()
        return cur.lastrowid
    except Exception as e:
        log.error("insert_transaction failed [%s]: %s", service_name, e)
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# CUSTOMER UPSERT
# ═══════════════════════════════════════════════════════════════════════════════

def upsert_customer(phone: str, spend: float, name: str | None = None) -> bool:
    """
    Insert a new customer or accumulate spend + visit_count for an existing one.
    Stores name if provided (and not already set).
    Returns True on success, False on failure.
    """
    if not phone or not str(phone).strip():
        return False
    phone = str(phone).strip()
    name  = str(name).strip() if name else None
    now   = datetime.now().isoformat()
    try:
        conn = get_db()
        conn.execute(
            """
            INSERT INTO customers (phone, name, total_spend, visit_count, last_seen, created_at)
            VALUES (?, ?, ?, 1, ?, ?)
            ON CONFLICT(phone) DO UPDATE SET
                total_spend = total_spend + excluded.total_spend,
                visit_count = visit_count + 1,
                last_seen   = excluded.last_seen,
                name        = CASE WHEN excluded.name IS NOT NULL AND excluded.name != ''
                                   THEN excluded.name ELSE customers.name END
            """,
            (phone, name, round(spend, 4), now, now),
        )
        conn.commit()
        return True
    except Exception as e:
        log.error("upsert_customer failed [%s]: %s", phone, e)
        return False


# ═══════════════════════════════════════════════════════════════════════════════
# ATOMIC MULTI-ITEM TRANSACTION INSERT
# ═══════════════════════════════════════════════════════════════════════════════

def insert_transactions_atomic(
    items: list[dict],
    timestamp: str,
    payment_mode: str,
    customer_phone: str | None = None,
) -> list[int]:
    """
    Insert ALL items in a single SQLite transaction (BEGIN/COMMIT).
    If ANY item fails, the entire transaction is rolled back and an exception is raised.

    Each item dict must have:
        name, qty, revenue, cost_total, profit,
        base_price, final_price, override_type, override_value

    Also upserts customer within the same transaction.

    Returns list of inserted row IDs (same order as items).
    Raises RuntimeError on any failure — caller must handle and return 500.
    """
    if not items:
        raise ValueError("items list is empty")

    conn = get_db()
    row_ids = []
    total_revenue = sum(i["revenue"] for i in items)

    try:
        conn.execute("BEGIN")
        for item in items:
            cur = conn.execute(
                """
                INSERT INTO transactions
                    (timestamp, service_name, qty, revenue, cost, profit,
                     payment_mode, customer_phone,
                     original_price, final_price, override_type, override_value)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    timestamp,
                    item["name"],
                    item["qty"],
                    round(item["revenue"],    4),
                    round(item["cost_total"], 4),
                    round(item["profit"],     4),
                    payment_mode,
                    customer_phone or None,
                    round(item["base_price"],  4) if item.get("base_price")  is not None else None,
                    round(item["final_price"], 4) if item.get("final_price") is not None else None,
                    item.get("override_type"),
                    round(item["override_value"], 4) if item.get("override_value") is not None else None,
                ),
            )
            row_ids.append(cur.lastrowid)

        if customer_phone and str(customer_phone).strip():
            now = datetime.now().isoformat()
            conn.execute(
                """
                INSERT INTO customers (phone, total_spend, visit_count, last_seen, created_at)
                VALUES (?, ?, 1, ?, ?)
                ON CONFLICT(phone) DO UPDATE SET
                    total_spend = total_spend + excluded.total_spend,
                    visit_count = visit_count + 1,
                    last_seen   = excluded.last_seen
                """,
                (str(customer_phone).strip(), round(total_revenue, 4), now, now),
            )

        conn.execute("COMMIT")
        log.info("insert_transactions_atomic: committed %d items ts=%s", len(items), timestamp)
        return row_ids

    except Exception as e:
        try:
            conn.execute("ROLLBACK")
            log.error("insert_transactions_atomic: ROLLBACK after error: %s", e)
        except Exception as rb_err:
            log.error("insert_transactions_atomic: ROLLBACK itself failed: %s", rb_err)
        raise RuntimeError(f"Atomic transaction failed: {e}") from e


# ═══════════════════════════════════════════════════════════════════════════════
# PRICE EVENT INSERT
# ═══════════════════════════════════════════════════════════════════════════════

def insert_price_event(
    service_name: str,
    old_price:    float,
    new_price:    float,
    source:       str = "manual",
) -> bool:
    """
    Record a price change event for audit trail and elasticity learning.
    Returns True on success.
    """
    try:
        conn = get_db()
        conn.execute(
            """
            INSERT INTO price_events
                (service_name, old_price, new_price, timestamp, source)
            VALUES (?, ?, ?, ?, ?)
            """,
            (service_name, round(old_price, 4), round(new_price, 4),
             datetime.now().isoformat(), source),
        )
        conn.commit()
        return True
    except Exception as e:
        log.error("insert_price_event failed [%s]: %s", service_name, e)
        return False


# ═══════════════════════════════════════════════════════════════════════════════
# SERVICE UPSERT
# ═══════════════════════════════════════════════════════════════════════════════

def upsert_service(
    name:       str,
    category:   str  = "",
    price:      float = 0.0,
    cost:       float = 0.0,
    margin_pct: float = 0.0,
    role:       str  = "",
) -> bool:
    """
    Insert or update a service record in the DB mirror.
    Returns True on success.
    """
    try:
        conn = get_db()
        now  = datetime.now().isoformat()
        if margin_pct == 0.0 and price > 0:
            margin_pct = round((price - cost) / price * 100, 2)
        conn.execute(
            """
            INSERT INTO services
                (name, category, price, cost, margin_pct, role, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(name) DO UPDATE SET
                category   = excluded.category,
                price      = excluded.price,
                cost       = excluded.cost,
                margin_pct = excluded.margin_pct,
                role       = excluded.role,
                updated_at = excluded.updated_at
            """,
            (name, category, round(price, 4), round(cost, 4),
             round(margin_pct, 4), role, now),
        )
        conn.commit()
        return True
    except Exception as e:
        log.error("upsert_service failed [%s]: %s", name, e)
        return False


# ═══════════════════════════════════════════════════════════════════════════════
# READ FUNCTIONS (used when USE_DB_READ = True in app.py)
# ═══════════════════════════════════════════════════════════════════════════════

def get_services_from_db() -> dict:
    """
    Return a {name: svc_dict} map from the DB services table.
    Schema matches what _get_services() returns from Excel.
    """
    try:
        conn = get_db()
        rows = conn.execute(
            "SELECT name, category, price, cost, margin_pct, role FROM services"
        ).fetchall()
        result = {}
        for row in rows:
            name = row["name"]
            result[name] = {
                "name":       name,
                "category":   row["category"] or "Other",
                "price":      row["price"]     or 0.0,
                "cost":       row["cost"]      or 0.0,
                "margin_pct": row["margin_pct"] or 0.0,
                "role":       row["role"]      or "filler",
                "units_sold": 0,
                "priority":   5.0,
                "needs_price":row["price"] == 0.0 and name != "Other",
            }
        return result
    except Exception as e:
        log.error("get_services_from_db failed: %s", e)
        return {}


def get_customer_profile(phone: str) -> dict | None:
    """
    Return {total_spend, visit_count, avg_ticket, last_seen} for a customer.
    Returns None if not found.
    """
    if not phone or not str(phone).strip():
        return None
    try:
        conn = get_db()
        row  = conn.execute(
            """
            SELECT total_spend, visit_count, last_seen
            FROM   customers
            WHERE  phone = ?
            """,
            (str(phone).strip(),),
        ).fetchone()
        if row is None:
            return None
        spend  = row["total_spend"]  or 0.0
        visits = row["visit_count"]  or 0
        return {
            "total_spend":  round(spend, 2),
            "visit_count":  visits,
            "avg_ticket":   round(spend / visits, 2) if visits > 0 else 0.0,
            "last_seen":    row["last_seen"],
        }
    except Exception as e:
        log.error("get_customer_profile failed [%s]: %s", phone, e)
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# INTEGRITY / STATS
# ═══════════════════════════════════════════════════════════════════════════════

def get_db_stats() -> dict:
    """
    Return aggregate counts for the /db-integrity endpoint.
    All failures return 0 counts — never raises.
    """
    stats = {
        "db_connected":      False,
        "services_count":    0,
        "transactions_count":0,
        "customers_count":   0,
        "last_price_event":  None,
    }
    try:
        conn = get_db()
        stats["services_count"]     = conn.execute("SELECT COUNT(*) FROM services").fetchone()[0]
        stats["transactions_count"] = conn.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]
        stats["customers_count"]    = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
        row = conn.execute(
            "SELECT timestamp FROM price_events ORDER BY id DESC LIMIT 1"
        ).fetchone()
        stats["last_price_event"] = row[0] if row else None
        stats["db_connected"]     = True
    except Exception as e:
        log.error("get_db_stats failed: %s", e)
    return stats


# ═══════════════════════════════════════════════════════════════════════════════
# OVERRIDE ANALYTICS  (Phase 7)
# ═══════════════════════════════════════════════════════════════════════════════

def get_override_stats(service_name: str) -> dict:
    """
    Phase 7 — Override intelligence per service.
    Returns avg discount, frequency, margin loss, and discount histogram.
    Never raises — returns partial data on failure.
    """
    base = {
        "service":               service_name,
        "total_transactions":    0,
        "override_count":        0,
        "override_frequency":    0.0,
        "avg_override_price":    None,
        "avg_discount_pct":      None,
        "margin_loss_pct":       None,
        "discount_distribution": {},
    }
    try:
        conn  = get_db()
        total = conn.execute(
            "SELECT COUNT(*) FROM transactions WHERE service_name = ?",
            (service_name,),
        ).fetchone()[0]
        if total == 0:
            return base

        rows = conn.execute(
            """
            SELECT final_price, original_price, override_type, override_value
            FROM   transactions
            WHERE  service_name  = ?
              AND  override_type IS NOT NULL
              AND  override_type != 'none'
            """,
            (service_name,),
        ).fetchall()

        override_count = len(rows)
        if override_count == 0:
            return {**base, "total_transactions": total}

        final_prices = [r[0] for r in rows if r[0] is not None]
        orig_prices  = [r[1] for r in rows if r[1] is not None]
        disc_vals    = [r[3] for r in rows if r[2] == "discount_pct"
                        and r[3] is not None]

        avg_override  = round(sum(final_prices) / len(final_prices), 2) \
                        if final_prices else None
        avg_disc      = round(sum(disc_vals)    / len(disc_vals),    2) \
                        if disc_vals else None
        paired        = [(o, f) for o, f in zip(orig_prices, final_prices)
                         if o and o > 0]
        margin_loss   = round(
            sum((o - f) / o * 100 for o, f in paired) / len(paired), 2
        ) if paired else None

        dist: dict[str, int] = {}
        for d in disc_vals:
            bucket = f"{int(d // 5) * 5}-{int(d // 5) * 5 + 5}%"
            dist[bucket] = dist.get(bucket, 0) + 1

        return {
            "service":               service_name,
            "total_transactions":    total,
            "override_count":        override_count,
            "override_frequency":    round(override_count / total, 4),
            "avg_override_price":    avg_override,
            "avg_discount_pct":      avg_disc,
            "margin_loss_pct":       margin_loss,
            "discount_distribution": dist,
        }
    except Exception as e:
        log.error("get_override_stats failed [%s]: %s", service_name, e)
        return base


# ═══════════════════════════════════════════════════════════════════════════════
# BACKGROUND SYNC HELPER  (called by app.py sync thread)
# ═══════════════════════════════════════════════════════════════════════════════

def sync_services_from_dict(svc_map: dict):
    """
    Bulk-upsert all services from an in-memory svc_map into the DB.
    Called by the background sync thread every 300 s.
    Non-fatal: errors are logged, not re-raised.
    """
    if not svc_map:
        return
    try:
        conn = get_db()
        now  = datetime.now().isoformat()
        rows = []
        for name, svc in svc_map.items():
            price      = svc.get("price", 0.0)
            cost       = svc.get("cost",  0.0)
            margin_pct = (round((price - cost) / price * 100, 2)
                          if price > 0 else 0.0)
            rows.append((
                name,
                svc.get("category",   "Other"),
                round(price, 4),
                round(cost, 4),
                margin_pct,
                svc.get("role", "filler"),
                now,
            ))
        conn.executemany(
            """
            INSERT INTO services
                (name, category, price, cost, margin_pct, role, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(name) DO UPDATE SET
                category   = excluded.category,
                price      = excluded.price,
                cost       = excluded.cost,
                margin_pct = excluded.margin_pct,
                role       = excluded.role,
                updated_at = excluded.updated_at
            """,
            rows,
        )
        conn.commit()
        log.info("DB sync: upserted %d services", len(rows))
    except Exception as e:
        log.error("sync_services_from_dict failed: %s", e)


# ═══════════════════════════════════════════════════════════════════════════════
# DB-FIRST ENGINE FUNCTIONS  (v5.0 upgrade — replaces Excel scans)
# ═══════════════════════════════════════════════════════════════════════════════

def build_demand_from_db(days: int = 30) -> dict:
    """
    DB-first demand cache builder. Replaces _build_demand_cache() Excel scan.
    Returns identical schema so the rest of app.py needs zero changes.
    """
    import math, collections
    from datetime import datetime, timedelta

    now      = datetime.now()
    cutoff   = (now - timedelta(days=days)).isoformat()[:19]
    c7       = (now - timedelta(days=7)).isoformat()[:19]
    c24      = (now - timedelta(hours=24)).isoformat()[:19]
    midpoint = (now - timedelta(days=15)).isoformat()[:19]

    try:
        conn = get_db()
        rows = conn.execute(
            "SELECT service_name, qty, timestamp FROM transactions "
            "WHERE timestamp >= ? ORDER BY timestamp ASC",
            (cutoff,)
        ).fetchall()
    except Exception as e:
        log.error("build_demand_from_db query failed: %s", e)
        return {}

    totals = collections.defaultdict(lambda: {
        "qty": 0, "tx": 0, "h24": 0, "d7": 0, "d30": 0,
        "d7_days":  collections.defaultdict(int),
        "d30_days": collections.defaultdict(int),
        "recent_half": 0, "old_half": 0,
        "hourly":  collections.defaultdict(int),
        "weekday": collections.defaultdict(int),
        "last_seen": None, "decay_qty": 0.0,
    })

    for row in rows:
        svc = row["service_name"]
        qty = max(0, int(row["qty"] or 0))
        if qty == 0:
            continue
        ts_str = str(row["timestamp"])[:19]
        try:
            dt = datetime.fromisoformat(ts_str)
        except Exception:
            continue

        t = totals[svc]
        t["qty"] += qty
        t["tx"]  += 1
        t["hourly"][dt.hour]       += qty
        t["weekday"][dt.weekday()] += qty
        if t["last_seen"] is None or dt > t["last_seen"]:
            t["last_seen"] = dt

        age_days = (now - dt).total_seconds() / 86400
        t["decay_qty"] += qty * (2.0 if age_days <= 7 else 1.0)

        if ts_str >= c24:
            t["h24"] += qty
        if ts_str >= c7:
            t["d7"] += qty
            t["d7_days"][ts_str[:10]] += qty
        t["d30"] += qty
        t["d30_days"][ts_str[:10]] += qty
        if ts_str >= midpoint:
            t["recent_half"] += qty
        else:
            t["old_half"] += qty

    cache = {}
    for svc, t in totals.items():
        tx  = t["tx"]
        qty = t["qty"]
        d7  = t["d7"]
        d30 = t["d30"]
        h24 = t["h24"]

        demand_score = qty / tx if tx > 0 else 0
        smooth_denom = max(d7, max(1, math.ceil(qty / 7.0)))
        velocity     = round(h24 / smooth_denom, 4)

        recency_bonus = 0.2 if (t["last_seen"] and
            (now - t["last_seen"]).total_seconds() < 86400) else 0.0
        volume_score = min(1.0, math.log1p(tx) / math.log1p(30))
        confidence   = round(min(1.0, volume_score * 0.8 + recency_bonus), 3)

        rolling_avg_7d  = round(d7  / 7.0,  3)
        rolling_avg_30d = round(d30 / 30.0, 3) if d30 > 0 else 0.0

        if d30 > 0:
            r = t["recent_half"]
            o = t["old_half"]
            trend = ("new"     if o == 0 else
                     "rising"  if r > o * 1.2 else
                     "falling" if r < o * 0.8 else "stable")
        else:
            trend = "new" if tx > 0 else "—"

        cache[svc] = {
            "total_qty":      qty,
            "total_tx":       tx,
            "last_24h":       h24,
            "last_7d":        d7,
            "last_30d":       d30,
            "demand_score":   round(demand_score, 3),
            "velocity":       velocity,
            "confidence":     confidence,
            "rolling_avg_7d": rolling_avg_7d,
            "rolling_avg_30d":rolling_avg_30d,
            "trend":          trend,
            "decay_score":    round(t["decay_qty"] / max(tx, 1), 3),
            "hourly_dist":    dict(t["hourly"]),
            "weekday_dist":   dict(t["weekday"]),
            "last_seen":      t["last_seen"].isoformat() if t["last_seen"] else None,
        }
    log.info("build_demand_from_db: %d services from DB", len(cache))
    return cache


def build_bundle_from_db(session_window: int = 30,
                         min_session_size: int = 2,
                         min_confidence: float = 0.02) -> tuple[dict, int]:
    """
    DB-first bundle stats builder. Replaces _build_bundle_stats() Excel scan.
    Returns (stats_dict, total_tx) — identical schema.
    """
    import collections, math
    from datetime import datetime

    try:
        conn = get_db()
        rows = conn.execute(
            "SELECT timestamp, service_name FROM transactions ORDER BY timestamp ASC"
        ).fetchall()
    except Exception as e:
        log.error("build_bundle_from_db query failed: %s", e)
        return {}, 0

    if not rows:
        return {}, 0

    tx_events = []
    for row in rows:
        try:
            ts = datetime.fromisoformat(str(row["timestamp"])[:19]).timestamp()
        except Exception:
            continue
        tx_events.append((ts, row["service_name"]))

    tx_events.sort(key=lambda x: x[0])
    total_tx = len(tx_events)

    sessions   = []
    cur        = [tx_events[0][1]]
    last_ts    = tx_events[0][0]

    for ts, svc in tx_events[1:]:
        if ts - last_ts <= session_window:
            cur.append(svc)
        else:
            if len(set(cur)) >= min_session_size:
                sessions.append(cur)
            cur = [svc]
        last_ts = ts
    if len(set(cur)) >= min_session_size:
        sessions.append(cur)

    pair_counts = collections.defaultdict(int)
    for session in sessions:
        unique = list(set(session))
        for i in range(len(unique)):
            for j in range(i + 1, len(unique)):
                pair_counts[tuple(sorted([unique[i], unique[j]]))] += 1

    total_sessions = max(len(sessions), 1)
    stats = {}
    for (a, b), count in pair_counts.items():
        strength = round(count / total_sessions, 4)
        if strength < min_confidence:
            continue
        stats[(a, b)] = {
            "count":      count,
            "strength":   strength,
            "confidence": round(min(1.0, math.log1p(count) / math.log1p(20)), 3),
        }

    log.info("build_bundle_from_db: %d pairs from %d tx", len(stats), total_tx)
    return stats, total_tx


def get_daily_stats_from_db(date_str: str) -> tuple[float, float, int]:
    """
    Return (revenue, profit, count) for a given date string (YYYY-MM-DD).
    Replaces _load_daily_from_file().
    """
    try:
        conn = get_db()
        row = conn.execute(
            "SELECT COALESCE(SUM(revenue),0), COALESCE(SUM(profit),0), COUNT(*) "
            "FROM transactions WHERE timestamp LIKE ?",
            (date_str + "%",)
        ).fetchone()
        return float(row[0]), float(row[1]), int(row[2])
    except Exception as e:
        log.error("get_daily_stats_from_db failed [%s]: %s", date_str, e)
        return 0.0, 0.0, 0


def get_price_memory_from_db(service_name: str) -> dict:
    """
    Rebuild price_memory dict for one service from price_events table.
    Used at startup to restore price memory without needing price_memory.json.
    """
    try:
        conn = get_db()
        rows = conn.execute(
            "SELECT old_price, new_price, timestamp, source "
            "FROM price_events WHERE service_name = ? ORDER BY id ASC",
            (service_name,)
        ).fetchall()
        if not rows:
            return {}

        history      = []
        price_events = []
        last_price   = None
        last_update  = 0.0
        change_count = 0

        from datetime import datetime
        import time as _time

        for row in rows:
            old_p   = float(row["old_price"]  or 0)
            new_p   = float(row["new_price"]   or 0)
            ts_str  = row["timestamp"]
            try:
                ts = datetime.fromisoformat(ts_str[:19]).timestamp()
            except Exception:
                ts = 0.0

            history.append({"from": old_p, "to": new_p, "ts": ts_str})
            price_events.append({"ts": ts, "old": old_p, "new": new_p,
                                 "ratio": round(new_p / old_p, 4) if old_p > 0 else 1.0})
            last_price  = new_p
            last_update = ts
            change_count += 1

        if len(history) > 20:
            history = history[-20:]
        if len(price_events) > 10:
            price_events = price_events[-10:]

        last_row    = rows[-1]
        old_last    = float(last_row["old_price"] or 0)
        new_last    = float(last_row["new_price"]  or 0)
        direction   = ("up" if new_last > old_last else
                       "down" if new_last < old_last else "same")

        return {
            "last_price":     last_price,
            "last_update":    last_update,
            "change_count":   change_count,
            "cooldown_until": 0.0,   # expired — don't re-apply old cooldowns
            "direction":      direction,
            "history":        history,
            "price_events":   price_events,
        }
    except Exception as e:
        log.error("get_price_memory_from_db failed [%s]: %s", service_name, e)
        return {}


# ═══════════════════════════════════════════════════════════════════════════════
# UDHAAR (CREDIT) LEDGER
# ═══════════════════════════════════════════════════════════════════════════════

def add_udhaar_entry(
    phone: str,
    amount: float,
    entry_type: str,
    note: str | None = None,
    reference_txn_id: int | None = None,
) -> int | None:
    """
    Insert one udhaar ledger row.
    entry_type: 'debit'  — credit given to customer (owes us)
                'credit' — payment received from customer
    Returns new row id, or None on failure.
    """
    phone = str(phone).strip()
    if not phone or len(phone) < 10:
        log.error("add_udhaar_entry: invalid phone '%s'", phone)
        return None
    if entry_type not in ("debit", "credit"):
        log.error("add_udhaar_entry: invalid type '%s'", entry_type)
        return None
    if amount <= 0:
        log.error("add_udhaar_entry: amount must be > 0, got %s", amount)
        return None

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        conn = get_db()
        cur  = conn.execute(
            """
            INSERT INTO udhaar_ledger
                (customer_phone, type, amount, note, reference_txn_id, timestamp)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (phone, entry_type, round(amount, 4), note, reference_txn_id, timestamp),
        )
        conn.commit()
        log.info("add_udhaar_entry: %s phone=%s amount=%.2f ref=%s",
                 entry_type, phone, amount, reference_txn_id)
        return cur.lastrowid
    except Exception as e:
        log.error("add_udhaar_entry failed [%s]: %s", phone, e)
        return None


def get_customer_balance(phone: str) -> dict:
    """
    Return udhaar balance summary for a customer.
    {
        "phone": ...,
        "name":         str | None,
        "total_debit":  float,   # total credit given
        "total_credit": float,   # total repaid
        "balance":      float,   # outstanding (debit - credit)
        "entry_count":  int,
        "last_entry":   str | None,  # timestamp of most recent ledger entry
        "visit_count":  int,
    }
    """
    phone = str(phone).strip()
    try:
        conn = get_db()
        row  = conn.execute(
            """
            SELECT
                COALESCE(SUM(CASE WHEN type='debit'  THEN amount ELSE 0 END), 0) AS total_debit,
                COALESCE(SUM(CASE WHEN type='credit' THEN amount ELSE 0 END), 0) AS total_credit,
                COUNT(*) AS entry_count,
                MAX(timestamp) AS last_entry
            FROM udhaar_ledger
            WHERE customer_phone = ?
            """,
            (phone,),
        ).fetchone()
        # Also fetch customer profile if exists
        cust = conn.execute(
            "SELECT name, visit_count FROM customers WHERE phone = ?", (phone,)
        ).fetchone()
        debit  = float(row["total_debit"])
        credit = float(row["total_credit"])
        return {
            "phone":        phone,
            "name":         cust["name"] if cust else None,
            "total_debit":  round(debit,  2),
            "total_credit": round(credit, 2),
            "balance":      round(debit - credit, 2),
            "entry_count":  int(row["entry_count"]),
            "last_entry":   row["last_entry"],
            "visit_count":  int(cust["visit_count"]) if cust else 0,
        }
    except Exception as e:
        log.error("get_customer_balance failed [%s]: %s", phone, e)
        return {"phone": phone, "name": None, "total_debit": 0, "total_credit": 0,
                "balance": 0, "entry_count": 0, "last_entry": None, "visit_count": 0}


def get_udhaar_history(phone: str, limit: int = 50) -> list[dict]:
    """
    Full ledger for a customer ordered by timestamp desc.
    Returns list of dicts with: id, type, amount, note, reference_txn_id, timestamp.
    """
    phone = str(phone).strip()
    try:
        conn = get_db()
        rows = conn.execute(
            """
            SELECT id, type, amount, note, reference_txn_id, timestamp
            FROM udhaar_ledger
            WHERE customer_phone = ?
            ORDER BY timestamp DESC
            LIMIT ?
            """,
            (phone, limit),
        ).fetchall()
        return [dict(r) for r in rows]
    except Exception as e:
        log.error("get_udhaar_history failed [%s]: %s", phone, e)
        return []


def get_top_debtors(limit: int = 10) -> list[dict]:
    """
    Return top customers by outstanding balance (highest first).
    Includes aging: last_entry_ts for 30/60/90-day risk flags.
    """
    try:
        conn = get_db()
        rows = conn.execute(
            """
            SELECT
                u.customer_phone,
                c.name AS customer_name,
                SUM(CASE WHEN u.type='debit'  THEN u.amount ELSE 0 END) AS total_debit,
                SUM(CASE WHEN u.type='credit' THEN u.amount ELSE 0 END) AS total_credit,
                SUM(CASE WHEN u.type='debit'  THEN u.amount ELSE -u.amount END) AS balance,
                COUNT(*) AS entry_count,
                MAX(u.timestamp) AS last_entry_ts,
                MIN(CASE WHEN u.type='debit' THEN u.timestamp END) AS first_debit_ts
            FROM udhaar_ledger u
            LEFT JOIN customers c ON c.phone = u.customer_phone
            GROUP BY u.customer_phone
            HAVING balance > 0
            ORDER BY balance DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()

        now = datetime.now()
        result = []
        for r in rows:
            try:
                last_ts = datetime.fromisoformat(str(r["last_entry_ts"])[:19])
                days_since = (now - last_ts).days
            except Exception:
                days_since = 0

            try:
                first_ts = datetime.fromisoformat(str(r["first_debit_ts"])[:19])
                age_days = (now - first_ts).days
            except Exception:
                age_days = 0

            risk = ("high"   if days_since >= 60 else
                    "medium" if days_since >= 30 else "low")

            result.append({
                "phone":               r["customer_phone"],
                "name":                r["customer_name"],
                "total_debit":         round(float(r["total_debit"]),  2),
                "total_credit":        round(float(r["total_credit"]), 2),
                "balance":             round(float(r["balance"]),      2),
                "entry_count":         int(r["entry_count"]),
                "last_entry_ts":       r["last_entry_ts"],
                "age_days":            age_days,
                "days_since_activity": days_since,
                "risk":                risk,
            })
        return result
    except Exception as e:
        log.error("get_top_debtors failed: %s", e)
        return []


def get_udhaar_summary() -> dict:
    """Global udhaar summary: total outstanding, total debtors, today's activity."""
    try:
        conn = get_db()
        row = conn.execute(
            """
            SELECT
                COALESCE(SUM(CASE WHEN type='debit'  THEN amount ELSE 0 END), 0) AS total_given,
                COALESCE(SUM(CASE WHEN type='credit' THEN amount ELSE 0 END), 0) AS total_recovered,
                COUNT(DISTINCT customer_phone) AS total_customers
            FROM udhaar_ledger
            """
        ).fetchone()

        today = datetime.now().strftime("%Y-%m-%d")
        today_row = conn.execute(
            """
            SELECT
                COALESCE(SUM(CASE WHEN type='debit'  THEN amount ELSE 0 END), 0) AS today_given,
                COALESCE(SUM(CASE WHEN type='credit' THEN amount ELSE 0 END), 0) AS today_recovered
            FROM udhaar_ledger
            WHERE timestamp LIKE ?
            """,
            (today + "%",),
        ).fetchone()

        given     = float(row["total_given"])
        recovered = float(row["total_recovered"])
        return {
            "total_outstanding":  round(given - recovered, 2),
            "total_given":        round(given,     2),
            "total_recovered":    round(recovered, 2),
            "total_customers":    int(row["total_customers"]),
            "today_given":        round(float(today_row["today_given"]),     2),
            "today_recovered":    round(float(today_row["today_recovered"]), 2),
        }
    except Exception as e:
        log.error("get_udhaar_summary failed: %s", e)
        return {
            "total_outstanding": 0, "total_given": 0, "total_recovered": 0,
            "total_customers": 0, "today_given": 0, "today_recovered": 0,
        }
def log_system_event(
    event_type: str,
    entity_type: str,
    entity_id: str,
    old_value=None,
    new_value=None,
    operator: str = "system",
) -> bool:
    """
    Append one row to system_events audit table.
    event_type : 'credit_action' | 'price_override' | 'risk_change' | 'udhaar_blocked'
    entity_type: 'customer' | 'service'
    entity_id  : phone or service_name
    Never raises.
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        conn = get_db()
        _retry_execute(conn,
            "INSERT INTO system_events "
            "(event_type, entity_type, entity_id, old_value, new_value, operator, timestamp) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (event_type, entity_type, entity_id,
             str(old_value) if old_value is not None else None,
             str(new_value) if new_value is not None else None,
             operator, ts),
        )
        conn.commit()
        return True
    except Exception as e:
        log.error("log_system_event failed: %s", e)
        return False


def set_customer_credit_limit(phone: str, limit: float, operator: str = "manual") -> bool:
    """Set credit_limit for a customer. Logs to system_events. Returns True on success."""
    phone = str(phone).strip()
    if not phone:
        return False
    try:
        conn = get_db()
        row = conn.execute(
            "SELECT credit_limit FROM customers WHERE phone=?", (phone,)
        ).fetchone()
        old_val = float(row["credit_limit"]) if row and row["credit_limit"] is not None else 0.0
        _retry_execute(conn,
            "UPDATE customers SET credit_limit=? WHERE phone=?", (round(limit, 2), phone)
        )
        conn.commit()
        log_system_event("credit_action", "customer", phone,
                         old_value=old_val, new_value=round(limit, 2), operator=operator)
        log.info("set_customer_credit_limit: phone=%s limit=%.2f", phone, limit)
        return True
    except Exception as e:
        log.error("set_customer_credit_limit failed [%s]: %s", phone, e)
        return False


def calculate_risk_level(phone: str) -> str:
    """
    Derive risk from udhaar ledger history. No new table required.
    Risk is based on age of oldest unpaid balance.
    Returns: 'low' | 'medium' | 'high'
    """
    phone = str(phone).strip()
    try:
        conn = get_db()
        rows = conn.execute(
            "SELECT type, amount, timestamp FROM udhaar_ledger "
            "WHERE customer_phone=? ORDER BY timestamp ASC",
            (phone,)
        ).fetchall()
        if not rows:
            return "low"

        total_debit  = sum(float(r["amount"]) for r in rows if r["type"] == "debit")
        total_credit = sum(float(r["amount"]) for r in rows if r["type"] == "credit")
        outstanding  = total_debit - total_credit

        if outstanding <= 0.01:
            return "low"

        # Age of oldest unpaid debit entry
        oldest_ts = next(
            (r["timestamp"] for r in rows if r["type"] == "debit"), None
        )
        if not oldest_ts:
            return "low"
        try:
            age_days = (datetime.now() - datetime.strptime(
                str(oldest_ts)[:19], "%Y-%m-%d %H:%M:%S"
            )).days
        except Exception:
            age_days = 0

        if age_days > 30:
            return "high"
        elif age_days > 10:
            return "medium"
        return "low"
    except Exception as e:
        log.error("calculate_risk_level failed [%s]: %s", phone, e)
        return "low"


def update_customer_risk(phone: str) -> str:
    """Recompute and persist risk_level for a customer. Returns new level."""
    level = calculate_risk_level(phone)
    try:
        conn = get_db()
        old_row = conn.execute(
            "SELECT risk_level FROM customers WHERE phone=?", (str(phone).strip(),)
        ).fetchone()
        old_level = old_row["risk_level"] if old_row else "low"
        _retry_execute(conn,
            "UPDATE customers SET risk_level=? WHERE phone=?",
            (level, str(phone).strip())
        )
        conn.commit()
        if old_level != level:
            log_system_event("risk_change", "customer", phone,
                             old_value=old_level, new_value=level)
    except Exception as e:
        log.error("update_customer_risk failed [%s]: %s", phone, e)
    return level


def get_cashflow_metrics() -> dict:
    """
    Real cashflow truth: separates collected cash from credit exposure.
    Returns booked vs collected vs real_profit for today and all-time.
    Never raises.
    """
    try:
        conn  = get_db()
        today = datetime.now().strftime("%Y-%m-%d")

        # Today booked (all payment modes)
        all_today = conn.execute(
            "SELECT COALESCE(SUM(revenue),0) rev, COALESCE(SUM(profit),0) prof "
            "FROM transactions WHERE substr(timestamp,1,10)=?", (today,)
        ).fetchone()

        # Today collected (non-Udhaar only)
        cash_today = conn.execute(
            "SELECT COALESCE(SUM(revenue),0) rev, COALESCE(SUM(profit),0) prof "
            "FROM transactions WHERE substr(timestamp,1,10)=? "
            "AND payment_mode != 'Udhaar'", (today,)
        ).fetchone()

        # All-time udhaar outstanding
        ledger = conn.execute(
            "SELECT "
            "  COALESCE(SUM(CASE WHEN type='debit'  THEN amount END),0) given, "
            "  COALESCE(SUM(CASE WHEN type='credit' THEN amount END),0) recovered "
            "FROM udhaar_ledger"
        ).fetchone()

        # Active debtors count
        debtor_count = conn.execute(
            "SELECT COUNT(DISTINCT customer_phone) FROM ("
            "  SELECT customer_phone, "
            "    SUM(CASE WHEN type='debit' THEN amount ELSE -amount END) bal "
            "  FROM udhaar_ledger GROUP BY customer_phone HAVING bal > 0"
            ")"
        ).fetchone()[0]

        booked_rev  = float(all_today["rev"]  or 0)
        booked_prof = float(all_today["prof"] or 0)
        coll_rev    = float(cash_today["rev"]  or 0)
        coll_prof   = float(cash_today["prof"] or 0)
        total_given = float(ledger["given"]    or 0)
        recovered   = float(ledger["recovered"] or 0)
        outstanding = round(total_given - recovered, 2)

        # Estimated blocked profit using today's average margin
        avg_margin    = (booked_prof / booked_rev) if booked_rev > 0 else 0
        blocked_profit = round(outstanding * avg_margin, 2)

        return {
            "today":              today,
            "booked_revenue":     round(booked_rev,  2),
            "collected_revenue":  round(coll_rev,    2),
            "udhaar_revenue":     round(booked_rev - coll_rev, 2),
            "booked_profit":      round(booked_prof, 2),
            "real_profit":        round(coll_prof,   2),
            "outstanding_udhaar": outstanding,
            "total_given":        round(total_given, 2),
            "total_recovered":    round(recovered,   2),
            "blocked_profit":     blocked_profit,
            "active_debtors":     int(debtor_count or 0),
            "collection_rate":    round(coll_rev / booked_rev, 3) if booked_rev > 0 else 1.0,
        }
    except Exception as e:
        log.error("get_cashflow_metrics failed: %s", e)
        return {
            "today": "", "booked_revenue": 0, "collected_revenue": 0,
            "udhaar_revenue": 0, "booked_profit": 0, "real_profit": 0,
            "outstanding_udhaar": 0, "total_given": 0, "total_recovered": 0,
            "blocked_profit": 0, "active_debtors": 0, "collection_rate": 1.0,
        }


def get_overdue_entries(days: int = 7) -> list[dict]:
    """
    Return customers with outstanding udhaar older than `days` days.
    Sorted by balance descending.
    """
    try:
        conn = get_db()
        rows = conn.execute(
            "SELECT u.customer_phone, "
            "  SUM(CASE WHEN u.type='debit' THEN u.amount ELSE 0 END) - "
            "  SUM(CASE WHEN u.type='credit' THEN u.amount ELSE 0 END) AS balance, "
            "  MIN(u.timestamp) oldest_entry, "
            "  MAX(u.timestamp) latest_entry, "
            "  c.name, c.risk_level, c.credit_limit "
            "FROM udhaar_ledger u "
            "LEFT JOIN customers c ON c.phone = u.customer_phone "
            "GROUP BY u.customer_phone "
            "HAVING balance > 0 AND "
            "  julianday('now') - julianday(MIN(u.timestamp)) > ? "
            "ORDER BY balance DESC",
            (days,)
        ).fetchall()
        return [dict(r) for r in rows]
    except Exception as e:
        log.error("get_overdue_entries failed: %s", e)
        return []


def get_system_events(
    limit: int = 50,
    event_type: str | None = None,
    entity_id: str | None = None,
) -> list[dict]:
    """Return system audit events with optional filters."""
    try:
        conn   = get_db()
        q      = "SELECT * FROM system_events WHERE 1=1"
        params: list = []
        if event_type:
            q += " AND event_type=?"; params.append(event_type)
        if entity_id:
            q += " AND entity_id=?"; params.append(entity_id)
        q += " ORDER BY timestamp DESC LIMIT ?"
        params.append(limit)
        rows = conn.execute(q, params).fetchall()
        return [dict(r) for r in rows]
    except Exception as e:
        log.error("get_system_events failed: %s", e)
        return []
