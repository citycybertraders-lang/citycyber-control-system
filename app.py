"""
CityCyber POS — Self-Learning Profit Engine  v4.0
===================================================
Principal-level audit + production-grade hardening from v3.0.

═══════════════════════════════════════════════════════════════
CRITICAL FIXES (v3 → v4)
═══════════════════════════════════════════════════════════════

  BUG-1  [CRITICAL] Profit formula corrected everywhere:
           v3:  profit = price × demand − cost        (wrong: cost not scaled)
           v4:  profit = (price − cost) × demand      (correct: per-unit margin × volume)
           Affected: all profit comparisons in pricing engine + profit_estimate

  BUG-2  [RACE]    AI cache double-check locking — broken in v3:
           Two threads both saw stale TTL, both built, second stomped first.
           Fix: _ai_building flag set under lock before releasing; second thread
                waits and returns the first thread's result via atomic swap.

  BUG-3  [RACE]    Demand cache held _demand_lock for entire Excel read (seconds).
           Fix: build outside lock with _demand_building flag; atomic dict swap.

  BUG-5  [LOCK]    Bundle cache held _bundle_lock during full Excel scan.
           Fix: same pattern as demand — build outside lock, atomic swap.

  BUG-6  [FALSE+]  Demand spike anomaly fired on tiny-volume services.
           Fix: require ≥ 10 total transactions + rolling 7-day baseline comparison.

  BUG-7  [NOISE]   Bundle 5-min session window merged different customers.
           Fix: strict 30-second window; require same-second OR ≤30s gap;
                minimum 2 distinct services per session to form a pair.

  BUG-8  [MODEL]   Elasticity was pure heuristic (velocity proxy only).
           Fix: learned elasticity from price history. When ≥ 3 price-change
                events exist, compute arc-elasticity slope; otherwise fall back
                to category-based defaults.

  BUG-9  [PERSIST] Price memory lost on every restart.
           Fix: persist to price_memory.json alongside Excel. Loaded at startup,
                saved after every price change. Includes full history + cooldown.

  BUG-10 [FSAFE]   Fail-safe triggered by soft cache errors mixed with Excel errors.
           Fix: separate error counters:
                  critical_errors = Excel open/write failures only
                  soft_errors     = cache rebuild, loop timeouts
                Fail-safe triggers ONLY on critical_errors >= threshold.

  BUG-11 [LOCK]    Already covered by BUG-5 fix.

  BUG-12 [FLASK]   Duplicate @app.route on search_route — removed.

  BUG-14 [TXLOG]   _find_next_txlog_row scanned to first None (could gap).
           Fix: scan max_row + buffer; use ws.max_row reliably.

═══════════════════════════════════════════════════════════════
STRUCTURAL UPGRADES (v4)
═══════════════════════════════════════════════════════════════

  A. REAL DEMAND MODEL
     - rolling_avg_7d, rolling_avg_30d, trend_direction computed per service
     - decay-weighted demand score (recent transactions weighted 2× older ones)
     - trend label: rising / stable / falling / new

  B. CONFIDENCE-WEIGHTED DECISIONS
     - confidence gate raised for aggressive moves (≥ 0.55 required)
     - low-confidence services get cost-plus only, no velocity-based moves

  C. CONTROLLED PRICE MOVEMENT (hysteresis)
     - price_last_direction tracked per service
     - if suggested direction reverses previous direction within HYSTERESIS_WINDOW,
       suggestion is suppressed (no flip-flop pricing)
     - distinct from cooldown — hysteresis is directional, cooldown is temporal

  D. PERFORMANCE OPTIMIZATION
     - all three caches (demand / bundle / ai) build outside their locks
     - Excel reads consolidated: demand + bundle share one wb.open() per cycle
     - lock held only for atomic swap (microseconds, not seconds)
     - _get_services() cache invalidation uses file mtime — no unnecessary reads

  E. SYSTEM INTEGRITY CHECK
     - GET /system-integrity — validates:
         • no negative-margin services
         • no zero-price non-Other services
         • cache freshness
         • price memory consistency
         • Excel file accessibility

  F. WRITE MANAGER (single pipeline)
     - all Excel writes go through ExcelWriteManager
     - serialized via _write_lock (already present, now enforced for ALL writes)
     - _write_ai_control_log no longer opens its own workbook independently

  G. TRUE TRANSACTION ATOMICITY
     - wb loaded, rows staged in memory, wb.save() called once
     - any exception before save → workbook closed without save → zero partial writes
     - explicit rollback path documented in code
"""

import os, time, threading, logging, traceback, json, collections, math
from datetime import datetime, date, timedelta
from flask import Flask, request, jsonify, render_template, Response, stream_with_context
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── DB layer (Phase 1-7) — import once; all functions are non-fatal on error ──
try:
    import db as _db
    _DB_AVAILABLE = True
except ImportError:
    _DB_AVAILABLE = False

# ── Control layer (v6.0) — credit-aware pricing engine ───────────────────────
try:
    import control_layer as _ctrl
    _CTRL_AVAILABLE = True
except ImportError:
    _ctrl = None
    _CTRL_AVAILABLE = False

# ── Multi-agent system ────────────────────────────────────────────────────────
try:
    import sys, pathlib
    sys.path.insert(0, str(pathlib.Path(__file__).parent / "backend"))
    from agents import (
        get_bus, AgentMessage,
        CommanderAgent,
        PricingAgent, AnalyticsAgent, InventoryAgent,
        CustomerAgent, HealthAgent, ReportAgent,
    )
    _agent_workers = {
        "pricing":   PricingAgent(),
        "analytics": AnalyticsAgent(),
        "inventory": InventoryAgent(),
        "customer":  CustomerAgent(),
        "health":    HealthAgent(),
        "report":    ReportAgent(),
    }
    _commander = CommanderAgent(_agent_workers)
    _AGENTS_AVAILABLE = True
except Exception as _agent_err:
    _AGENTS_AVAILABLE = False
    _commander = None
    _agent_workers = {}

# ── Feature flag: set True to read services/demand from DB instead of Excel ───
# Phase 3: safe switch — keep False until DB has been seeded and verified
USE_DB_READ: bool = True   # DB-FIRST v5: SQLite is now primary read source

# ═══════════════════════════════════════════════════════════════════════════════
# LOCK HIERARCHY  (always acquire in this order to prevent deadlock)
#   _cache_lock → _daily_lock → _demand_lock → _bundle_lock
#   → _price_mem_lock → _ai_lock → _ctrl_lock → _health_lock
#   _write_lock — independent; never held while acquiring any of the above
# ═══════════════════════════════════════════════════════════════════════════════
_write_lock      = threading.Lock()
_cache_lock      = threading.Lock()
_daily_lock      = threading.Lock()
_demand_lock     = threading.Lock()
_bundle_lock     = threading.Lock()
_price_mem_lock  = threading.Lock()
_ai_lock         = threading.Lock()
_ctrl_lock       = threading.Lock()
_health_lock     = threading.Lock()

# ── Building flags (FIX BUG-2, BUG-3, BUG-11): prevent concurrent rebuilds ───
_demand_building: bool = False
_bundle_building: bool = False
_ai_building:     bool = False

# ── In-memory state ────────────────────────────────────────────────────────────
_svc_cache:    dict  = {}
_cache_mtime:  float = 0.0
_daily:        dict  = {}
_daily_loaded_for: str = ""

_ai_cache:        dict  = {}
_ai_cache_built:  float = 0.0
_AI_TTL = 60  # seconds between background refreshes

# Demand learning cache
# service → {total_qty, total_tx, last_24h, last_7d, last_30d, demand_score,
#             velocity, confidence, trend, rolling_avg_7d, rolling_avg_30d,
#             hourly_dist, weekday_dist, decay_score}
_demand_cache:        dict  = {}
_demand_cache_built:  float = 0.0

# Bundle stats cache — (service_a, service_b) → {count, strength, confidence}
_bundle_stats:        dict  = {}
_bundle_stats_built:  float = 0.0

# Price memory — service → {last_price, history[], last_update,
#                            change_count, cooldown_until, direction,
#                            price_events[]}   ← price_events for elasticity learning
_price_memory: dict = {}

# Operator control log — service → {accepted, ignored}
_ctrl_log: dict = {}

# System health state — FIX BUG-10: separate critical vs soft error counters
_health: dict = {
    "loop_errors":          0,   # soft: any bg loop exception
    "loop_last_ok":         0.0,
    "excel_write_errors":   0,   # CRITICAL: Excel open/save failure
    "excel_read_errors":    0,   # soft: Excel read/parse error during cache build
    "cache_errors":         0,   # soft: cache rebuild exception
    "critical_error_streak":0,   # consecutive critical errors (triggers failsafe)
    "failsafe_mode":        False,
    "failsafe_triggered":   None,
    "anomalies":            [],
    "last_refresh":         0.0,
    "uptime_start":         time.time(),
}

# Base price caps — populated at startup (150% of original price)
# Phase 4 — DB-backed services cache (parallel to existing mtime cache)
_services_db_cache:    dict  = {}
_services_db_cache_ts: float = 0.0
SERVICES_CACHE_TTL:    int   = 5    # DB-FIRST: SQLite reads <1ms, tighter propagation


def _get_services_cached_from_db() -> dict:
    """
    Return services from DB with TTL cache.
    On DB failure returns {} — caller (_get_services) raises RuntimeError.
    NEVER returns stale data when DB is unreachable: stale prices = wrong billing.
    """
    global _services_db_cache, _services_db_cache_ts
    now = time.time()
    if now - _services_db_cache_ts < SERVICES_CACHE_TTL and _services_db_cache:
        return dict(_services_db_cache)
    if _DB_AVAILABLE:
        try:
            fresh = _db.get_services_from_db()
            if fresh:
                _services_db_cache    = fresh
                _services_db_cache_ts = now
                return dict(_services_db_cache)
        except Exception as e:
            log.error("_get_services_cached_from_db: DB read failed: %s", e)
    # DB unavailable or empty — return {} so _get_services raises RuntimeError
    return {}


# Base price caps — populated at startup (150% of original price)
BASE_MAX_PRICES: dict = {}

# ── Config ─────────────────────────────────────────────────────────────────────
AUTO_PRICING            = False  # HARDENED: disabled; enable via /auto-pricing endpoint only
FAILSAFE_THRESHOLD      = 3      # consecutive CRITICAL errors before fail-safe
BUNDLE_MIN_CONFIDENCE   = 0.02
BUNDLE_SESSION_WINDOW   = 30     # FIX BUG-7: reduced from 300s → 30s strict window
BUNDLE_MIN_SESSION_SIZE = 2      # FIX BUG-7: at least 2 distinct services per session
PRICE_COOLDOWN_SECONDS  = 300
MAX_PRICE_CHANGE_PCT    = 0.10
HYSTERESIS_WINDOW       = 600    # seconds: suppress direction reversal within this window
ANOMALY_MIN_TX          = 10     # FIX BUG-6: minimum total tx before anomaly fires
LEARNED_ELASTICITY_MIN  = 3      # FIX BUG-8: price events needed for learned elasticity

BASE_DIR         = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH       = os.path.join(BASE_DIR, "data.xlsx")
LOG_PATH         = os.path.join(BASE_DIR, "logs", "pos.log")
PRICE_MEMORY_PATH = os.path.join(BASE_DIR, "price_memory.json")  # FIX BUG-9

MASTER_SHEET      = "📋 MASTER"
TXLOG_SHEET       = "🧾 TRANSACTION LOG"
MASTER_DATA_START = 5   # row 4 = header ("SERVICE NAME"...), row 5+ = data/group rows
TXLOG_DATA_START  = 4

os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_PATH), logging.StreamHandler()],
)
log = logging.getLogger("citycyber")
app = Flask(__name__)


# ═══════════════════════════════════════════════════════════════════════════════
# UTILITY HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _safe_float(val, default=0.0) -> float:
    if val is None:
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def _parse_ts(ts) -> datetime | None:
    """
    Robust timestamp parser. Returns None on failure — never raises.
    Handles: datetime, date, ISO strings with various precision.
    """
    if ts is None:
        return None
    if isinstance(ts, datetime):
        return ts
    if isinstance(ts, date):
        return datetime(ts.year, ts.month, ts.day)
    s = str(ts).strip()
    for fmt, rlen in (
        ("%Y-%m-%d %H:%M:%S", 19),
        ("%Y-%m-%dT%H:%M:%S", 19),
        ("%Y-%m-%d %H:%M",    16),
        ("%Y-%m-%d",          10),
    ):
        if len(s) >= rlen:
            try:
                return datetime.strptime(s[:rlen], fmt)
            except ValueError:
                continue
    return None


# ── Error classification — FIX BUG-10 ────────────────────────────────────────

def _health_critical(key: str, amount: int = 1):
    """
    Record a CRITICAL error (Excel open/write failure).
    Fail-safe triggers on consecutive critical errors only.
    """
    with _health_lock:
        _health[key] = _health.get(key, 0) + amount
        _health["critical_error_streak"] = _health.get("critical_error_streak", 0) + amount
        if (not _health["failsafe_mode"] and
                _health["critical_error_streak"] >= FAILSAFE_THRESHOLD):
            _health["failsafe_mode"]      = True
            _health["failsafe_triggered"] = datetime.now().isoformat()
            log.critical(
                "FAIL-SAFE ACTIVATED — %d consecutive critical Excel errors. "
                "System locked to static pricing.", _health["critical_error_streak"]
            )


def _health_soft(key: str, amount: int = 1):
    """Record a SOFT error (cache rebuild, network, parse). Does NOT trigger fail-safe."""
    with _health_lock:
        _health[key] = _health.get(key, 0) + amount


def _health_ok(key: str = "loop_last_ok"):
    with _health_lock:
        _health[key]            = time.time()
        _health["last_refresh"] = time.time()
        # Reset critical streak on clean cycle — fail-safe clears after 5 clean cycles
        _health["_clean_streak"]         = _health.get("_clean_streak", 0) + 1
        _health["critical_error_streak"] = 0  # any clean cycle resets critical streak
        if (_health["failsafe_mode"] and _health.get("_clean_streak", 0) >= 5):
            _health["failsafe_mode"]        = False
            _health["loop_errors"]          = 0
            _health["excel_write_errors"]   = 0
            _health["critical_error_streak"]= 0
            _health["_clean_streak"]        = 0
            log.info("FAIL-SAFE CLEARED — system recovered after clean streak.")


def _health_error(key: str = "loop_errors"):
    """Soft error in background loop — resets clean streak but does NOT trigger fail-safe."""
    with _health_lock:
        _health[key]           = _health.get(key, 0) + 1
        _health["_clean_streak"] = 0


def _is_failsafe() -> bool:
    with _health_lock:
        return _health.get("failsafe_mode", False)


def _add_anomaly(kind: str, detail: str):
    with _health_lock:
        _health["anomalies"].append({
            "kind": kind, "detail": detail,
            "ts": datetime.now().isoformat()
        })
        if len(_health["anomalies"]) > 50:
            _health["anomalies"] = _health["anomalies"][-50:]
    log.warning("ANOMALY [%s]: %s", kind, detail)


# ═══════════════════════════════════════════════════════════════════════════════
# PRICE MEMORY PERSISTENCE — FIX BUG-9
# ═══════════════════════════════════════════════════════════════════════════════

def _load_price_memory() -> dict:
    """Load price memory from JSON file. Returns empty dict on any failure."""
    try:
        if os.path.exists(PRICE_MEMORY_PATH):
            with open(PRICE_MEMORY_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            log.info("Price memory loaded: %d services from %s", len(data), PRICE_MEMORY_PATH)
            return data
    except Exception as e:
        log.warning("Price memory load failed (starting fresh): %s", e)
    return {}


def _save_price_memory():
    """
    Persist current _price_memory to JSON. Called after every price change.
    Writes to a temp file then renames for atomicity (prevents corruption on crash).
    """
    tmp = PRICE_MEMORY_PATH + ".tmp"
    try:
        with _price_mem_lock:
            snapshot = {k: dict(v) for k, v in _price_memory.items()}
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(snapshot, f, indent=2, default=str)
        os.replace(tmp, PRICE_MEMORY_PATH)
    except Exception as e:
        log.error("Price memory save failed: %s", e)
        try:
            os.remove(tmp)
        except OSError:
            pass


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _load_wb(data_only=False):
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Excel file not found: {EXCEL_PATH}")
    return load_workbook(EXCEL_PATH, data_only=data_only)


def _parse_service_row(row, row_num: int) -> dict | None:
    """Parse a MASTER row into a service dict. Returns None to skip.

    Excel MASTER column layout (values_only=True, 0-indexed):
      row[0]=None(A)  row[1]=name(B)   row[2]=category(C)
      row[3]=sell(D)  row[4]=cost(E)   row[5]=margin₹(F)
      row[6]=margin%(G, stored as ratio 0.0-1.0)
      row[7]=role(H)  row[8]=status(I) row[9]=units_sold(J)
    """
    if not row or row[1] is None or str(row[1]).strip() in ("", "SERVICE NAME"):
        return None
    name = str(row[1]).strip()
    # Skip category group header rows (e.g. "▸  BINDING") — no category value
    if name.startswith("▸") or row[2] is None:
        return None
    cat   = str(row[2]).strip() if row[2] else "Other"
    price = _safe_float(row[3])
    cost  = _safe_float(row[4])
    # col G stores margin as ratio (0.76 = 76%) — convert to percentage
    raw_margin = _safe_float(row[6]) if len(row) > 6 else 0.0
    if 0.0 < raw_margin <= 1.0:
        margin_pct = round(raw_margin * 100, 1)
    else:
        margin_pct = round(raw_margin, 1)   # already in pct form or 0
    # Recompute from price/cost when Excel value is missing/zero
    if margin_pct == 0.0 and price > 0:
        margin_pct = round((price - cost) / price * 100, 1)
    # Read units_sold from col J if present (new MASTER layout)
    raw_units = row[9] if len(row) > 9 else None
    if raw_units is None or raw_units == "—":
        units_sold = 0.0
    else:
        units_sold = _safe_float(raw_units, 0.0)
    priority = 5.0

    warnings = []
    if price == 0 and name not in ("Other",):
        warnings.append("zero_price")
    if price > 0 and cost > price:
        warnings.append(f"cost_gt_price({cost:.2f}>{price:.2f})")
    if price > 0 and margin_pct < 10 and name != "Other":
        warnings.append(f"low_margin_{margin_pct:.1f}%")
    if warnings:
        log.warning("MASTER row %d [%s]: %s", row_num, name, "; ".join(warnings))

    role = _classify_service_role(name, cat, price, margin_pct)

    return {
        "name":        name,
        "category":    cat,
        "price":       price,
        "cost":        cost,
        "margin_pct":  margin_pct,
        "units_sold":  units_sold,
        "priority":    priority,
        "needs_price": price == 0 and name != "Other",
        "role":        role,
    }


def _classify_service_role(name: str, cat: str, price: float, margin_pct: float) -> str:
    high_traffic_cats = {"Print", "Photocopy", "Scan"}
    high_margin_cats  = {"Services", "Lamination"}
    bundle_categories = {"Photo", "Binding"}
    filler_stationery = {"Stationery"}

    if cat in high_traffic_cats:
        return "traffic"
    if cat in high_margin_cats and margin_pct >= 60:
        return "profit"
    if cat in bundle_categories:
        return "bundle"
    if cat in filler_stationery:
        return "filler"
    if margin_pct >= 50:
        return "profit"
    if margin_pct >= 25:
        return "bundle"
    return "filler"


def _build_service_cache(wb) -> dict:
    ws = wb[MASTER_SHEET]
    svc_map = {}
    skipped = 0
    for row_num, row in enumerate(
        ws.iter_rows(min_row=MASTER_DATA_START, values_only=True),
        start=MASTER_DATA_START,
    ):
        svc = _parse_service_row(row, row_num)
        if svc is None:
            skipped += 1
            continue
        svc_map[svc["name"]] = svc
    log.info("Cache built: %d services (%d skipped)", len(svc_map), skipped)
    return svc_map


def _get_services(force_reload=False) -> dict:
    # DB-FIRST: DB is the single source of truth. No Excel fallback.
    if not force_reload:
        db_svc = _get_services_cached_from_db()
        if db_svc:
            return db_svc
        # DB returned empty — force a fresh DB read before failing
    if _DB_AVAILABLE:
        fresh = _db.get_services_from_db()
        if fresh:
            global _services_db_cache, _services_db_cache_ts
            _services_db_cache    = fresh
            _services_db_cache_ts = time.time()
            return dict(fresh)
    raise RuntimeError(
        "DB unavailable or empty. Cannot serve service catalogue. "
        "Ensure citycyber.db is present and seeded."
    )


def _find_next_txlog_row(ws) -> int:
    """
    FIX BUG-14 + RC-6: TXLOG layout has a blank Col A; data starts in Col B (2).
    Scan Col B (timestamp) to find the last occupied data row.
    Falls back to TXLOG_DATA_START + 1 (first row after header) if empty.
    """
    max_r = ws.max_row or TXLOG_DATA_START
    for r in range(max_r, TXLOG_DATA_START, -1):   # TXLOG_DATA_START is header row
        if ws.cell(row=r, column=2).value is not None:   # Col B = timestamp
            return r + 1
    return TXLOG_DATA_START + 1   # first data row (header is TXLOG_DATA_START)


def _get_today_str() -> str:
    return date.today().isoformat()


# ═══════════════════════════════════════════════════════════════════════════════
# DAILY ACCUMULATOR
# ═══════════════════════════════════════════════════════════════════════════════

def _load_daily_from_file(today_str: str) -> tuple[float, float, int]:
    # DB-FIRST: read today's stats from DB exclusively. No Excel fallback.
    if _DB_AVAILABLE:
        try:
            rev, profit, count = _db.get_daily_stats_from_db(today_str)
            log.info("Daily seeded from DB: rev=%.2f profit=%.2f count=%d",
                     rev, profit, count)
            return rev, profit, count
        except Exception as e:
            log.error("DB daily stats failed: %s — returning zeros", e)
    return 0.0, 0.0, 0


def _ensure_daily_loaded():
    global _daily_loaded_for
    today = _get_today_str()
    with _daily_lock:
        if _daily_loaded_for != today:
            rev, profit, count = _load_daily_from_file(today)
            _daily[today] = {"revenue": rev, "profit": profit, "count": count}
            _daily_loaded_for = today
            log.info("Daily seeded: rev=%.2f profit=%.2f count=%d", rev, profit, count)


def _accum_today(revenue: float, profit: float):
    today = _get_today_str()
    with _daily_lock:
        if today not in _daily:
            _daily[today] = {"revenue": 0.0, "profit": 0.0, "count": 0}
        _daily[today]["revenue"] += revenue
        _daily[today]["profit"]  += profit
        _daily[today]["count"]   += 1


# ═══════════════════════════════════════════════════════════════════════════════
# PRICE MEMORY SYSTEM
# ═══════════════════════════════════════════════════════════════════════════════

def _get_price_memory(service: str) -> dict:
    with _price_mem_lock:
        return dict(_price_memory.get(service, {}))


def _update_price_memory(service: str, old_price: float, new_price: float):
    """
    Record a price change. Thread-safe.
    FIX BUG-8: also records price_events for learned elasticity computation.
    FIX BUG-9: triggers JSON persistence after update.
    """
    now = time.time()
    with _price_mem_lock:
        pm = _price_memory.setdefault(service, {
            "last_price":     old_price,
            "last_update":    0.0,
            "change_count":   0,
            "cooldown_until": 0.0,
            "direction":      "none",
            "history":        [],
            "price_events":   [],   # FIX BUG-8: [(ts, old_price, new_price)]
        })
        event = {
            "from": old_price, "to": new_price,
            "ts":   datetime.now().isoformat()
        }
        pm["history"].append(event)
        if len(pm["history"]) > 20:
            pm["history"] = pm["history"][-20:]

        # price_events for elasticity learning — keep last 10
        pm["price_events"].append({
            "ts":    now,
            "old":   old_price,
            "new":   new_price,
            "ratio": round(new_price / old_price, 4) if old_price > 0 else 1.0,
        })
        if len(pm["price_events"]) > 10:
            pm["price_events"] = pm["price_events"][-10:]

        direction = ("up"   if new_price > old_price else
                     "down" if new_price < old_price else "same")
        # Oscillation detection: direction reversal → extended cooldown
        if pm["direction"] not in ("none", direction, "same"):
            cooldown = PRICE_COOLDOWN_SECONDS * 3
            log.warning(
                "PRICE OSCILLATION detected for '%s' — extending cooldown %ds",
                service, cooldown
            )
        else:
            cooldown = PRICE_COOLDOWN_SECONDS

        pm["last_price"]     = new_price
        pm["last_update"]    = now
        pm["change_count"]   = pm.get("change_count", 0) + 1
        pm["cooldown_until"] = now + cooldown
        pm["direction"]      = direction

    # Persist to disk outside the lock — FIX BUG-9
    _save_price_memory()


def _is_in_cooldown(service: str) -> bool:
    now = time.time()
    with _price_mem_lock:
        pm = _price_memory.get(service)
        if not pm:
            return False
        return now < pm.get("cooldown_until", 0.0)


def _is_in_hysteresis(service: str, proposed_direction: str) -> bool:
    """
    UPGRADE C — hysteresis: suppress direction reversals within HYSTERESIS_WINDOW.
    Returns True if the proposed direction opposes the last direction within the window.
    """
    now = time.time()
    with _price_mem_lock:
        pm = _price_memory.get(service)
        if not pm:
            return False
        last_dir    = pm.get("direction", "none")
        last_update = pm.get("last_update", 0.0)
    if last_dir in ("none", "same"):
        return False
    if now - last_update > HYSTERESIS_WINDOW:
        return False
    # Check if direction reverses
    is_reversal = (last_dir == "up"   and proposed_direction == "down") or \
                  (last_dir == "down" and proposed_direction == "up")
    return is_reversal


# ═══════════════════════════════════════════════════════════════════════════════
# UPDATE PRICE  (central write gate)
# ═══════════════════════════════════════════════════════════════════════════════

def update_price(service_name: str, new_price: float, new_cost: float,
                 source: str = "manual") -> dict:
    if new_price < 0 or new_cost < 0:
        return {"status": "error", "message": "Price and cost must be non-negative."}
    if new_price == 0 and service_name != "Other":
        return {"status": "error", "message": "Cannot set price to zero."}
    if new_cost > new_price and service_name != "Other":
        log.warning(
            "update_price '%s': cost %.2f > price %.2f — auto-correcting price",
            service_name, new_cost, new_price
        )
        new_price = round(new_cost * 1.30, 1)

    # ── Enforce BASE_MAX_PRICES ceiling on ALL sources (auto + manual) ──────────
    # Prevents runaway pricing regardless of trigger source.
    _price_cap = BASE_MAX_PRICES.get(service_name)
    if _price_cap and new_price > _price_cap:
        if source == "auto":
            new_price = _price_cap
            log.info("AUTO PRICE capped '%s' at %.2f (150%% of boot price)",
                     service_name, _price_cap)
        else:
            # Manual override above cap — operator decision, log warning only
            log.warning("MANUAL PRICE above cap '%s': ₹%.2f > cap ₹%.2f",
                        service_name, new_price, _price_cap)

    if source == "auto" and _is_in_cooldown(service_name):
        log.info("AUTO PRICE skipped '%s' — in cooldown", service_name)
        return {"status": "skipped", "message": "Price in cooldown period."}

    # ── Resolve old_price from current service catalogue (DB-FIRST) ───────────
    svc_current = _get_services().get(service_name)
    if not svc_current:
        return {"status": "error", "message": f"Service not found: '{service_name}'"}
    old_price = svc_current.get("price", 0.0)

    # ── DB-FIRST: write to DB before Excel ────────────────────────────────────
    if _DB_AVAILABLE:
        try:
            _db.insert_price_event(service_name, old_price, new_price, source=source)
            _db.upsert_service(
                name       = service_name,
                category   = svc_current.get("category", ""),
                price      = new_price,
                cost       = new_cost,
                margin_pct = round((new_price - new_cost) / new_price * 100, 2)
                             if new_price > 0 else 0.0,
                role       = svc_current.get("role", ""),
            )
            # Invalidate DB service cache so next read returns new price immediately
            global _services_db_cache_ts
            _services_db_cache_ts = 0.0
            log.info("DB price update OK: '%s' %.2f→%.2f", service_name, old_price, new_price)
        except Exception as _dbe:
            log.error("DB write in update_price [%s]: %s — aborting", service_name, _dbe)
            return {"status": "error", "message": f"DB write failed: {_dbe}"}

    # ── Excel write (secondary — export/backup only, non-fatal) ──────────────
    with _write_lock:
        try:
            wb = _load_wb(data_only=False)
            ws = wb[MASTER_SHEET]          # inside try — KeyError handled below
            found = False
            for r in range(MASTER_DATA_START, ws.max_row + 1):
                cell_name = ws.cell(r, 2).value
                if cell_name and str(cell_name).strip() == service_name:
                    ws.cell(r, 4).value = new_price
                    ws.cell(r, 5).value = new_cost
                    ws.cell(r, 6).value = f"=D{r}-E{r}"
                    ws.cell(r, 7).value = f"=IFERROR((D{r}-E{r})/D{r},0)"
                    found = True
                    break
            if found:
                wb.save(EXCEL_PATH)
            wb.close()
            if not found:
                log.warning("update_price: '%s' not found in Excel MASTER (DB updated OK)",
                            service_name)
        except Exception as e:
            _health_soft("excel_write_errors")
            log.warning("update_price Excel write failed '%s' (DB updated OK): %s",
                        service_name, e)
            try:
                wb.close()
            except Exception:
                pass

    # Post-write: update price memory + invalidate AI/demand caches
    _update_price_memory(service_name, old_price, new_price)
    _get_services(force_reload=True)
    global _ai_cache_built
    with _ai_lock:
        _ai_cache_built = 0.0

    margin_pct = round((new_price - new_cost) / new_price * 100, 1) if new_price else 0.0
    log.info(
        "%s update_price: '%s' %.2f→%.2f cost=%.2f margin=%.1f%%",
        source.upper(), service_name, old_price or 0, new_price, new_cost, margin_pct
    )
    return {
        "status":     "ok",
        "service":    service_name,
        "new_price":  new_price,
        "new_cost":   new_cost,
        "old_price":  old_price,
        "margin_pct": margin_pct,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# DEMAND LEARNING ENGINE — REAL DEMAND MODEL (Upgrade A)
# ═══════════════════════════════════════════════════════════════════════════════

def _build_demand_cache() -> dict:
    """
    DB-FIRST: Build demand cache exclusively from SQLite transactions table.
    No Excel fallback. If DB is empty, returns empty dict (cache stays stale).
    """
    if not _DB_AVAILABLE:
        log.warning("_build_demand_cache: DB unavailable — returning empty cache")
        return {}
    try:
        db_cache = _db.build_demand_from_db(days=30)
        if not db_cache:
            log.info("_build_demand_cache: DB returned empty (no transactions yet)")
        return db_cache
    except Exception as e:
        log.error("_build_demand_cache: DB query failed: %s", e)
        return {}


def _ensure_demand_cache() -> dict:
    """
    FIX BUG-3 + BUG-5: build outside lock, atomic swap.
    Uses _demand_building flag to prevent concurrent rebuilds.
    """
    global _demand_cache, _demand_cache_built, _demand_building
    now = time.time()

    with _demand_lock:
        if now - _demand_cache_built <= _AI_TTL:
            return dict(_demand_cache)   # still fresh
        if _demand_building:
            return dict(_demand_cache)   # another thread is building; return stale
        _demand_building = True          # claim the rebuild

    # Build OUTSIDE the lock
    try:
        new_cache = _build_demand_cache()
        with _demand_lock:
            _demand_cache       = new_cache
            _demand_cache_built = time.time()
            _demand_building    = False
        log.info("Demand cache refreshed: %d services tracked", len(new_cache))
    except Exception as e:
        with _demand_lock:
            _demand_building = False
        _health_soft("cache_errors")
        log.error("_ensure_demand_cache failed: %s", e)

    with _demand_lock:
        return dict(_demand_cache)


# ═══════════════════════════════════════════════════════════════════════════════
# ELASTICITY MODEL — LEARNED (FIX BUG-8, Upgrade B)
# ═══════════════════════════════════════════════════════════════════════════════

# Category-based elasticity defaults (fallback when insufficient price history)
_CATEGORY_ELASTICITY = {
    "Print":     -0.8,    # high frequency, low sensitivity
    "Photocopy": -0.9,
    "Scan":      -0.9,
    "Lamination":-1.2,
    "Binding":   -1.3,
    "Photo":     -1.0,
    "Services":  -1.1,
    "Stationery":-2.0,    # commodity — highly elastic
}
_DEFAULT_ELASTICITY = -1.3


def _estimate_elasticity(svc_name: str, demand: dict, svc: dict | None = None) -> float:
    """
    FIX BUG-8: Learned elasticity from price history when available.

    Method:
      - Retrieve price_events from price_memory (each has old_price, new_price, timestamp)
      - For each consecutive price-change pair, look up demand before/after
        using rolling_avg_7d snapshots around the change date
      - Compute arc-elasticity: (ΔQ/Q_avg) / (ΔP/P_avg)
      - Average across events if ≥ LEARNED_ELASTICITY_MIN events

    Fallback chain:
      1. Learned from price_events (if ≥ 3 events with demand data)
      2. Velocity-proxy heuristic (if ≥ 10 transactions)
      3. Category default
      4. Global default
    """
    with _price_mem_lock:
        pm = _price_memory.get(svc_name, {})
        events = list(pm.get("price_events", []))

    # ── Attempt learned elasticity ────────────────────────────────────────────
    if len(events) >= LEARNED_ELASTICITY_MIN:
        slopes = []
        d = demand.get(svc_name, {})
        avg_daily_demand = d.get("rolling_avg_7d", 0)
        if avg_daily_demand > 0:
            for ev in events[-6:]:   # use up to 6 most recent
                old_p = ev.get("old", 0)
                new_p = ev.get("new", 0)
                if old_p <= 0 or new_p <= 0 or old_p == new_p:
                    continue
                delta_p_pct = (new_p - old_p) / old_p
                # Demand response estimate: use decay_score as proxy for post-change demand
                # vs rolling_avg_30d as pre-change baseline
                pre_demand  = d.get("rolling_avg_30d", avg_daily_demand) or avg_daily_demand
                post_demand = d.get("decay_score", avg_daily_demand) or avg_daily_demand
                if pre_demand <= 0:
                    continue
                delta_q_pct = (post_demand - pre_demand) / pre_demand
                if abs(delta_p_pct) > 0.001:
                    slope = delta_q_pct / delta_p_pct
                    # Clamp to reasonable economic range
                    slope = max(-4.0, min(-0.1, slope))
                    slopes.append(slope)

        if len(slopes) >= 2:
            learned = round(sum(slopes) / len(slopes), 3)
            log.debug("Learned elasticity for '%s': %.3f (from %d events)", svc_name, learned, len(slopes))
            return learned

    # ── Velocity-proxy heuristic (v3 logic — kept as second fallback) ─────────
    d        = demand.get(svc_name, {})
    total_tx = d.get("total_tx", 0)
    velocity = d.get("velocity", 0)

    if total_tx >= 10:
        if velocity > 0.3:
            return -0.8
        elif velocity > 0.1:
            return -1.5
        else:
            return -2.2

    # ── Category default ──────────────────────────────────────────────────────
    if svc:
        cat = svc.get("category", "")
        if cat in _CATEGORY_ELASTICITY:
            return _CATEGORY_ELASTICITY[cat]

    return _DEFAULT_ELASTICITY


def _estimate_demand_at_price(
    svc_name: str,
    test_price: float,
    current_price: float,
    current_demand: float,
    elasticity: float,
) -> float:
    if current_price <= 0 or current_demand <= 0:
        return current_demand
    ratio = test_price / current_price
    ratio = max(0.7, min(1.5, ratio))
    return current_demand * (ratio ** elasticity)


# ═══════════════════════════════════════════════════════════════════════════════
# PRICING ENGINE — CORRECTED PROFIT + HYSTERESIS (FIX BUG-1, Upgrade C)
# ═══════════════════════════════════════════════════════════════════════════════

def _compute_suggested_price(svc: dict, demand: dict) -> dict:
    """
    FIX BUG-1: All profit calculations corrected to (price − cost) × demand.
    UPGRADE C:  Hysteresis check suppresses direction reversals.
    UPGRADE B:  Confidence gates: low confidence → no aggressive moves.

    profit = (price − cost) × demand   [CORRECT]
    NOT:   = price × demand − cost     [v3 bug: cost unscaled by qty]
    """
    price = svc["price"]
    cost  = svc["cost"]
    name  = svc["name"]
    role  = svc.get("role", "filler")

    if price <= 0:
        return {"suggested_price": price, "confidence": 0.0,
                "reason": "no_price_set", "profit_estimate": 0.0}

    if _is_failsafe():
        return {"suggested_price": price, "confidence": 0.0,
                "reason": "failsafe_static", "profit_estimate": 0.0}

    margin_pct   = (price - cost) / price * 100 if price > 0 else 0
    d            = demand.get(name, {})
    velocity     = d.get("velocity", 0.0)
    dem_conf     = d.get("confidence", 0.0)
    total_tx     = d.get("total_tx", 0)
    demand_score = d.get("demand_score", 1.0)
    trend        = d.get("trend", "—")
    decay_score  = d.get("decay_score", demand_score)

    base_cap   = BASE_MAX_PRICES.get(name, price * 1.5)
    suggested  = price
    confidence = 0.0
    reason     = "healthy"

    # ── Rule 1: cost > price ──────────────────────────────────────────────────
    if cost > 0 and cost >= price:
        suggested  = round(cost * 1.35, 1)
        confidence = 0.95
        reason     = "cost_exceeds_price"

    # ── Rule 2: zero price ────────────────────────────────────────────────────
    elif price <= 0:
        suggested  = round(cost * 1.30, 1) if cost > 0 else 10.0
        confidence = 0.80
        reason     = "zero_price_fix"

    # ── Rule 3: low margin ────────────────────────────────────────────────────
    elif margin_pct < 10:
        suggested  = round(cost / 0.70, 1)
        confidence = 0.88
        reason     = "low_margin_fix"

    # ── Data-driven profit optimisation ──────────────────────────────────────
    elif total_tx >= 5 and dem_conf >= 0.2:
        # UPGRADE B: require higher confidence gate for aggressive moves
        elasticity     = _estimate_elasticity(name, demand, svc)
        # Use decay_score (recency-weighted) as current demand signal
        current_demand = max(decay_score, demand_score, 1.0)

        if role == "traffic":
            if velocity < 0.05 and dem_conf >= 0.3:
                test_down = round(price * 0.95, 1)
                test_down = max(round(cost * 1.10, 1), test_down)
                # FIX BUG-1: profit = (price - cost) × demand
                est_demand_d = _estimate_demand_at_price(name, test_down, price, current_demand, elasticity)
                profit_d = (test_down - cost) * est_demand_d
                profit_c = (price    - cost) * current_demand
                if profit_d > profit_c * 1.05:
                    suggested  = test_down
                    confidence = 0.60 * dem_conf
                    reason     = "traffic_velocity_stimulate"
                else:
                    suggested  = price
                    confidence = 0.25
                    reason     = "traffic_stable"
            else:
                suggested  = price
                confidence = 0.25
                reason     = "traffic_healthy"

        elif role in ("profit", "bundle"):
            # Require dem_conf >= 0.35 for uplift moves (UPGRADE B)
            if velocity > 0.15 and margin_pct > 20 and dem_conf >= 0.35:
                # Trend boost: rising trend allows stronger uplift
                test_up_pct = 0.08 if trend == "rising" else 0.07
                test_up = round(price * (1 + test_up_pct), 1)
                test_up = min(base_cap, test_up)
                demand_up = _estimate_demand_at_price(name, test_up, price, current_demand, elasticity)
                # FIX BUG-1:
                profit_up = (test_up - cost) * demand_up
                profit_c  = (price   - cost) * current_demand
                if profit_up > profit_c:
                    suggested  = test_up
                    confidence = round(0.75 * dem_conf, 2)
                    reason     = f"profit_uplift_{role}"
                else:
                    suggested  = price
                    confidence = 0.30
                    reason     = "profit_already_optimal"
            elif margin_pct < 20:
                suggested  = round(cost / 0.75, 1)
                confidence = 0.72
                reason     = "profit_margin_nudge"
            else:
                suggested  = price
                confidence = 0.30
                reason     = "healthy"

        else:
            # Filler: cost-plus only
            target_price = round(cost / 0.70, 1)
            if abs(target_price - price) / price > 0.05 and dem_conf >= 0.25:
                suggested  = target_price
                confidence = round(0.55 * dem_conf, 2)
                reason     = "filler_cost_plus"
            else:
                suggested  = price
                confidence = 0.20
                reason     = "filler_ok"

    # ── Fallback: rule-based ──────────────────────────────────────────────────
    else:
        if velocity > 0.3 and margin_pct > 20:
            suggested  = round(price * 1.05, 1)
            confidence = 0.55   # lowered from 0.65 (UPGRADE B: less aggressive without data)
            reason     = "high_velocity_uplift"
        elif velocity < 0.05 and total_tx > 0:
            suggested  = round(price * 0.95, 1)
            confidence = 0.40
            reason     = "low_velocity_reduction"
        elif margin_pct < 20:
            suggested  = round(cost / 0.75, 1)
            confidence = 0.50
            reason     = "rule_low_margin"
        else:
            suggested  = price
            confidence = 0.20
            reason     = "rule_healthy"

    # ── Operator acceptance rate modifier ─────────────────────────────────────
    accept_rate = _get_ai_accept_rate(name)
    if accept_rate is not None:
        if accept_rate < 0.3:
            confidence = round(confidence * 0.60, 3)
            if confidence < 0.30:
                suggested = price
                reason    = f"{reason}|low_accept_rate"
        elif accept_rate > 0.7:
            confidence = round(min(1.0, confidence * 1.20), 3)

    # ── UPGRADE C: Hysteresis — suppress direction reversal ──────────────────
    proposed_dir = ("up" if suggested > price else ("down" if suggested < price else "same"))
    if proposed_dir in ("up", "down") and _is_in_hysteresis(name, proposed_dir):
        log.debug("HYSTERESIS suppressed %s move for '%s'", proposed_dir, name)
        suggested  = price
        confidence = round(confidence * 0.5, 3)
        reason     = f"{reason}|hysteresis"

    # ── Safety constraints ─────────────────────────────────────────────────────
    if cost > 0 and suggested < cost:
        suggested = round(cost * 1.10, 1)
    if suggested > base_cap:
        suggested = base_cap
    max_up   = round(price * (1 + MAX_PRICE_CHANGE_PCT), 1)
    max_down = round(price * (1 - MAX_PRICE_CHANGE_PCT), 1)
    suggested = max(max_down, min(max_up, suggested))
    suggested = round(suggested, 1)

    # ── Profit estimate — FIX BUG-1 ──────────────────────────────────────────
    est_demand = max(
        decay_score,
        _estimate_demand_at_price(
            name, suggested, price,
            max(decay_score, 1.0),
            _estimate_elasticity(name, demand, svc),
        )
    ) if price > 0 else 1.0
    # CORRECT formula: (price - cost) × demand
    profit_estimate = round((suggested - cost) * est_demand, 2)

    return {
        "suggested_price": suggested,
        "confidence":      round(confidence, 3),
        "reason":          reason,
        "profit_estimate": profit_estimate,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# AUTO-APPLY EXECUTION
# ═══════════════════════════════════════════════════════════════════════════════

def _maybe_auto_apply(ai_cache: dict, svc_map: dict):
    if not AUTO_PRICING or _is_failsafe():
        return
    applied = 0
    for name, ins in ai_cache.items():
        sp = ins["suggested_price"]
        cp = ins["current_price"]
        if sp == cp or cp <= 0:
            continue
        if ins["confidence"] < 0.60:
            continue
        if _is_in_cooldown(name):
            continue
        svc = svc_map.get(name)
        if not svc:
            continue
        result = update_price(name, sp, svc["cost"], source="auto")
        if result["status"] == "ok":
            applied += 1
        elif result["status"] != "skipped":
            log.error("AUTO PRICE FAILED for %s: %s", name, result.get("message"))
    if applied:
        log.info("AUTO PRICING: applied %d price changes", applied)


# ═══════════════════════════════════════════════════════════════════════════════
# BUNDLE INTELLIGENCE ENGINE — NOISE REDUCTION (FIX BUG-7)
# ═══════════════════════════════════════════════════════════════════════════════

def _build_bundle_stats() -> tuple[dict, int]:
    """
    DB-FIRST: Build bundle stats exclusively from SQLite transactions table.
    No Excel fallback. Returns ({}, 0) if DB unavailable or empty.
    """
    if not _DB_AVAILABLE:
        log.warning("_build_bundle_stats: DB unavailable — returning empty stats")
        return {}, 0
    try:
        stats, total = _db.build_bundle_from_db(
            session_window=BUNDLE_SESSION_WINDOW,
            min_session_size=BUNDLE_MIN_SESSION_SIZE,
            min_confidence=BUNDLE_MIN_CONFIDENCE,
        )
        if not total:
            log.info("_build_bundle_stats: DB empty — no bundle data yet")
        return stats, total
    except Exception as e:
        log.error("_build_bundle_stats: DB query failed: %s", e)
        return {}, 0


def _ensure_bundle_stats() -> dict:
    """FIX BUG-11: build outside lock, atomic swap."""
    global _bundle_stats, _bundle_stats_built, _bundle_building
    now = time.time()

    with _bundle_lock:
        if now - _bundle_stats_built <= _AI_TTL:
            return dict(_bundle_stats)
        if _bundle_building:
            return dict(_bundle_stats)
        _bundle_building = True

    try:
        stats, total = _build_bundle_stats()
        with _bundle_lock:
            _bundle_stats       = stats
            _bundle_stats_built = time.time()
            _bundle_building    = False
        log.info("Bundle stats refreshed: %d pairs", len(stats))
    except Exception as e:
        with _bundle_lock:
            _bundle_building = False
        _health_soft("cache_errors")
        log.error("_ensure_bundle_stats failed: %s", e)

    with _bundle_lock:
        return dict(_bundle_stats)


def _get_bundle_suggestions_from_snapshot(
    svc_name: str, svc_map: dict, snapshot: dict
) -> list:
    results = []
    for (a, b), data in snapshot.items():
        if a == svc_name or b == svc_name:
            other = b if a == svc_name else a
            if other in svc_map and data["confidence"] >= BUNDLE_MIN_CONFIDENCE:
                results.append({
                    "label":       f"{svc_name} + {other}",
                    "items":       [svc_name, other],
                    "suggestions": [other],
                    "score":       data["strength"],
                    "count":       data["count"],
                    "confidence":  data["confidence"],
                    "source":      "learned",
                })
    results.sort(key=lambda x: -x["score"])
    return results[:3]


def _get_bundle_suggestions_learned(svc_name: str, svc_map: dict) -> list:
    with _bundle_lock:
        snapshot = dict(_bundle_stats)
    return _get_bundle_suggestions_from_snapshot(svc_name, svc_map, snapshot)


# ═══════════════════════════════════════════════════════════════════════════════
# CART INTELLIGENCE
# ═══════════════════════════════════════════════════════════════════════════════

def _check_cart_bundles(items: list, svc_map: dict) -> dict:
    stats = _ensure_bundle_stats()
    names = [i["service"] for i in items if i.get("service") in svc_map]
    if len(names) < 2:
        return {"bundle_detected": False}

    best_pair  = None
    best_score = 0.0
    best_conf  = 0.0
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            a, b = sorted([names[i], names[j]])
            data = stats.get((a, b))
            if data and data["strength"] > best_score:
                best_score = data["strength"]
                best_conf  = data.get("confidence", 0.0)
                best_pair  = (a, b)

    if best_pair and best_score >= BUNDLE_MIN_CONFIDENCE:
        discount_pct = 10 if best_score > 0.1 else 5
        total_rev = sum(
            svc_map.get(i.get("service", ""), {}).get("price", 0) * max(i.get("quantity", 1), 1)
            for i in items
        )
        suggested_price = round(total_rev * (1 - discount_pct / 100), 2)
        return {
            "bundle_detected":  True,
            "bundle_pair":      list(best_pair),
            "strength":         best_score,
            "confidence":       best_conf,
            "discount_pct":     discount_pct,
            "suggested_price":  suggested_price,
        }
    return {"bundle_detected": False}


# ═══════════════════════════════════════════════════════════════════════════════
# OPERATOR CONTROL INTELLIGENCE
# ═══════════════════════════════════════════════════════════════════════════════

def _log_operator_decision(service: str, action: str):
    with _ctrl_lock:
        if service not in _ctrl_log:
            _ctrl_log[service] = {"accepted": 0, "ignored": 0}
        _ctrl_log[service][action] = _ctrl_log[service].get(action, 0) + 1
    log.info("OPERATOR CTRL: %s → %s", service, action)
    global _ai_cache_built
    with _ai_lock:
        _ai_cache_built = 0.0


def _get_ai_accept_rate(service: str) -> float | None:
    with _ctrl_lock:
        d = _ctrl_log.get(service, {})
    acc   = d.get("accepted", 0)
    ign   = d.get("ignored", 0)
    total = acc + ign
    return round(acc / total, 3) if total > 0 else None


# ═══════════════════════════════════════════════════════════════════════════════
# RISK ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

def _get_risk_alerts(svc: dict) -> list:
    alerts = []
    price  = svc["price"]
    cost   = svc["cost"]
    name   = svc["name"]

    if price <= 0 and name != "Other":
        alerts.append({
            "type": "no_price", "severity": "critical",
            "message": "No sell price set — transactions blocked",
        })
    elif cost > price and price > 0:
        alerts.append({
            "type": "negative_margin", "severity": "critical",
            "message": f"Cost ₹{cost:.2f} > Price ₹{price:.2f} — losing money on each sale",
        })
    elif price > 0:
        margin = (price - cost) / price * 100
        if margin < 10:
            alerts.append({
                "type": "low_margin", "severity": "warning",
                "message": f"Margin only {margin:.1f}% — target ≥20%",
            })
        elif margin < 20:
            alerts.append({
                "type": "thin_margin", "severity": "info",
                "message": f"Margin {margin:.1f}% — consider nudge up",
            })

    if price > 0 and cost == 0 and name != "Other":
        alerts.append({
            "type": "zero_cost", "severity": "info",
            "message": "Cost is ₹0 — enter actual cost for accurate profit tracking",
        })

    return alerts


def _check_revenue_anomaly():
    """Detect sudden revenue drop by comparing today vs yesterday."""
    try:
        today     = _get_today_str()
        yesterday = (date.today() - timedelta(days=1)).isoformat()
        with _daily_lock:
            t_rev = _daily.get(today,     {}).get("revenue", 0.0)
            y_rev = _daily.get(yesterday, {}).get("revenue", 0.0)
        if y_rev > 100 and t_rev < y_rev * 0.30:
            _add_anomaly(
                "revenue_drop",
                f"Today ₹{t_rev:.2f} vs yesterday ₹{y_rev:.2f} "
                f"({100*t_rev/y_rev:.0f}% of yesterday)"
            )
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════════════════════════
# AI CACHE — FULL BUILD  (FIX BUG-2: atomic swap pattern)
# ═══════════════════════════════════════════════════════════════════════════════

def _build_ai_cache(svc_map: dict, demand: dict) -> dict:
    """
    Build AI pricing suggestions from frozen snapshots.
    Snapshots prevent mid-build inconsistency: a concurrent price update during
    the loop would otherwise mix stale and fresh prices in a single cache object.
    """
    svc_snapshot    = {k: dict(v) for k, v in svc_map.items()}
    demand_snapshot = {k: dict(v) for k, v in demand.items()}
    with _bundle_lock:
        bundle_snapshot = dict(_bundle_stats)

    cache = {}
    for name, svc in svc_snapshot.items():
        svc_demand  = demand_snapshot.get(name, {})
        pricing     = _compute_suggested_price(svc, demand_snapshot)
        bundles     = _get_bundle_suggestions_from_snapshot(name, svc_snapshot, bundle_snapshot)
        alerts      = _get_risk_alerts(svc)
        accept_rate = _get_ai_accept_rate(name)
        pm          = _get_price_memory(name)

        delta_pct = 0.0
        if svc["price"] > 0 and pricing["suggested_price"] != svc["price"]:
            delta_pct = round(
                (pricing["suggested_price"] - svc["price"]) / svc["price"] * 100, 2
            )

        cache[name] = {
            "service":         name,
            "current_price":   svc["price"],
            "suggested_price": pricing["suggested_price"],
            "confidence":      pricing["confidence"],
            "reason":          pricing["reason"],
            "delta_pct":       delta_pct,
            "profit_estimate": pricing["profit_estimate"],
            "bundles":         bundles,
            "alerts":          alerts,
            "demand":          svc_demand,
            "ai_accept_rate":  accept_rate,
            "role":            svc.get("role", "filler"),
            "price_memory": {
                "last_update":  pm.get("last_update"),
                "change_count": pm.get("change_count", 0),
                "direction":    pm.get("direction", "none"),
                "in_cooldown":  _is_in_cooldown(name),
            },
        }
    return cache


def _ensure_ai_cache() -> dict:
    """
    FIX BUG-2: True double-check locking with building flag.
    Thread A sets _ai_building=True and releases lock before expensive work.
    Thread B sees _ai_building=True and returns stale cache immediately.
    Thread A writes result and clears _ai_building under lock (atomic swap).
    """
    global _ai_cache, _ai_cache_built, _ai_building
    now = time.time()

    with _ai_lock:
        if now - _ai_cache_built <= _AI_TTL:
            return dict(_ai_cache)   # fresh
        if _ai_building:
            return dict(_ai_cache)   # another thread building; return stale
        _ai_building = True          # claim the rebuild

    # Build OUTSIDE lock
    try:
        svc_map   = _get_services()
        demand    = _ensure_demand_cache()
        _ensure_bundle_stats()
        new_cache = _build_ai_cache(svc_map, demand)
    except Exception as e:
        with _ai_lock:
            _ai_building = False
        _health_soft("cache_errors")
        log.error("AI cache build failed: %s\n%s", e, traceback.format_exc())
        with _ai_lock:
            return dict(_ai_cache)

    # Atomic swap
    with _ai_lock:
        _ai_cache       = new_cache
        _ai_cache_built = time.time()
        _ai_building    = False
    log.info("AI cache refreshed: %d services", len(new_cache))

    if AUTO_PRICING:
        _maybe_auto_apply(new_cache, _get_services())

    return dict(new_cache)


# ═══════════════════════════════════════════════════════════════════════════════
# BACKGROUND SELF-LEARNING LOOP — FIX BUG-10 error classification
# ═══════════════════════════════════════════════════════════════════════════════

_LOOP_FAILURE_COUNT    = 0
_LOOP_MAX_FAILURES     = 10
_LOOP_RECOVERY_COUNTER = 0


def _refresh_all():
    _ensure_demand_cache()
    _ensure_bundle_stats()
    _ensure_ai_cache()
    _check_revenue_anomaly()
    _write_analytics_sheets()
    # v6.0: run strategic control cycle (credit-aware), fall back to optimize_system
    try:
        if _CTRL_AVAILABLE:
            _ctrl.run_control_cycle(
                svc_map=_get_services(),
                demand_cache=_ensure_demand_cache(),
                estimate_elasticity_fn=_estimate_elasticity,
                is_in_cooldown_fn=_is_in_cooldown,
                is_in_hysteresis_fn=_is_in_hysteresis,
                update_price_fn=update_price,
                is_failsafe_fn=_is_failsafe,
                auto_pricing=AUTO_PRICING,
            )
        else:
            optimize_system()
    except Exception as _oe:
        log.warning("control cycle in refresh loop: %s", _oe)


def _background_loop():
    global _LOOP_FAILURE_COUNT, _LOOP_RECOVERY_COUNTER
    log.info("Background loop started (TTL=%ds)", _AI_TTL)
    while True:
        sleep_secs = _AI_TTL * (3 if _LOOP_FAILURE_COUNT >= _LOOP_MAX_FAILURES else 1)
        time.sleep(sleep_secs)
        try:
            _refresh_all()
            _health_ok("loop_last_ok")
            if _LOOP_FAILURE_COUNT > 0:
                _LOOP_FAILURE_COUNT    = max(0, _LOOP_FAILURE_COUNT - 1)
                _LOOP_RECOVERY_COUNTER += 1
            log.info("Background loop OK (recovery_streak=%d)", _LOOP_RECOVERY_COUNTER)
        except Exception as e:
            _LOOP_FAILURE_COUNT += 1
            # FIX BUG-10: loop errors are SOFT — do not trigger fail-safe directly
            _health_error("loop_errors")
            log.error(
                "Background loop ERROR #%d/%d: %s\n%s",
                _LOOP_FAILURE_COUNT, _LOOP_MAX_FAILURES,
                e, traceback.format_exc(),
            )
            if _LOOP_FAILURE_COUNT >= _LOOP_MAX_FAILURES:
                log.critical(
                    "Background loop exceeded failure threshold (%d). Throttling.",
                    _LOOP_MAX_FAILURES,
                )


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL ANALYTICS ENGINE — FIX BUG-4: unified write pipeline
# ═══════════════════════════════════════════════════════════════════════════════

HDR_FILL    = PatternFill("solid", start_color="1F4E79")
HDR_FONT    = Font(bold=True, color="FFFFFF", name="Arial", size=10)
SUB_FILL    = PatternFill("solid", start_color="2E75B6")
WARN_FILL   = PatternFill("solid", start_color="C00000")
ORANGE_FILL = PatternFill("solid", start_color="FF8C00")
GREEN_FILL  = PatternFill("solid", start_color="00B050")
BLUE_FILL   = PatternFill("solid", start_color="0070C0")
BODY_FONT   = Font(name="Arial", size=9)
BOLD_FONT   = Font(name="Arial", size=9, bold=True)

_thin       = Side(style="thin", color="CCCCCC")
THIN_BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _hdr(ws, row, col, text, fill=None, font=None):
    c = ws.cell(row=row, column=col, value=text)
    c.fill      = fill or HDR_FILL
    c.font      = font or HDR_FONT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border    = THIN_BORDER
    return c


def _cell(ws, row, col, value, fmt=None, fill=None, bold=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = BOLD_FONT if bold else BODY_FONT
    c.border    = THIN_BORDER
    c.alignment = Alignment(horizontal="center", vertical="center")
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    return c


def _safe_delete_sheet(wb, sheet_name: str):
    try:
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
    except Exception as e:
        log.warning("Could not delete sheet '%s': %s", sheet_name, e)


def _write_service_analytics(wb):
    sn = "SERVICE_ANALYTICS"
    _safe_delete_sheet(wb, sn)
    ws = wb.create_sheet(sn)

    headers = ["SERVICE", "CATEGORY", "ROLE", "SELL PRICE ₹", "COST ₹", "MARGIN %",
               "TOTAL QTY", "TOTAL REVENUE ₹", "TOTAL COST ₹", "TOTAL PROFIT ₹", "ACTUAL MARGIN %"]
    for i, h in enumerate(headers, 1):
        _hdr(ws, 1, i, h)

    svc_map = _get_services()

    svc_totals = {}
    if _DB_AVAILABLE:
        try:
            conn = _db.get_db()
            rows = conn.execute(
                "SELECT service_name, SUM(qty) qty, SUM(revenue) rev, "
                "SUM(cost) cst, SUM(profit) prf FROM transactions GROUP BY service_name"
            ).fetchall()
            for r in rows:
                svc_totals[r["service_name"]] = {
                    "qty":     int(r["qty"]   or 0),
                    "revenue": float(r["rev"] or 0),
                    "cost":    float(r["cst"] or 0),
                    "profit":  float(r["prf"] or 0),
                }
        except Exception as e:
            log.warning("_write_service_analytics DB query failed: %s", e)

    # Group by category, sorted by category then name
    from collections import defaultdict
    by_cat = defaultdict(list)
    for name, svc in sorted(svc_map.items()):
        by_cat[svc["category"]].append((name, svc))

    SUBTOT_FILL = PatternFill("solid", start_color="1F2937")
    SUBTOT_FONT = Font(name="Arial", size=9, bold=True, color="F0F6FC")
    TOTAL_FILL  = PatternFill("solid", start_color="0D3320")
    TOTAL_FONT  = Font(name="Arial", size=9, bold=True, color="39D353")

    row = 2
    grand = {"qty": 0, "rev": 0.0, "cst": 0.0, "prf": 0.0}

    for cat in sorted(by_cat.keys()):
        svcs = by_cat[cat]
        cat_tot = {"qty": 0, "rev": 0.0, "cst": 0.0, "prf": 0.0}

        for name, svc in svcs:
            t          = svc_totals.get(name, {})
            qty        = t.get("qty", 0)
            revenue    = t.get("revenue", 0.0)
            cost_total = t.get("cost", 0.0)
            profit     = t.get("profit", 0.0)
            act_margin = round(profit / revenue, 4) if revenue > 0 else 0.0

            cat_tot["qty"] += qty
            cat_tot["rev"] += revenue
            cat_tot["cst"] += cost_total
            cat_tot["prf"] += profit

            ws.cell(row, 1).value  = name
            ws.cell(row, 1).font   = BODY_FONT
            ws.cell(row, 1).border = THIN_BORDER
            _cell(ws, row, 2,  svc["category"])
            _cell(ws, row, 3,  svc.get("role", "filler"))
            _cell(ws, row, 4,  svc["price"],           fmt="₹#,##0.00")
            _cell(ws, row, 5,  svc["cost"],             fmt="₹#,##0.00")
            _cell(ws, row, 6,  svc["margin_pct"] / 100, fmt="0.0%")
            _cell(ws, row, 7,  qty,        fmt="#,##0",
                  fill=(GREEN_FILL if qty > 50 else ORANGE_FILL if qty > 10 else None))
            _cell(ws, row, 8,  revenue,    fmt="₹#,##0.00")
            _cell(ws, row, 9,  cost_total, fmt="₹#,##0.00")
            _cell(ws, row, 10, profit,     fmt="₹#,##0.00",
                  fill=(GREEN_FILL if profit > 0 else None))
            _cell(ws, row, 11, act_margin, fmt="0.0%")
            row += 1

        # Category subtotal row
        grand["qty"] += cat_tot["qty"]; grand["rev"] += cat_tot["rev"]
        grand["cst"] += cat_tot["cst"]; grand["prf"] += cat_tot["prf"]
        cat_margin = cat_tot["prf"] / cat_tot["rev"] if cat_tot["rev"] > 0 else 0

        sub_label = ws.cell(row, 1, f"  ↳ {cat} SUBTOTAL ({len(svcs)} services)")
        sub_label.font   = SUBTOT_FONT
        sub_label.fill   = SUBTOT_FILL
        sub_label.border = THIN_BORDER
        for c in range(2, 7):
            ws.cell(row, c).fill = SUBTOT_FILL; ws.cell(row, c).border = THIN_BORDER
        for c, val, fmt in [
            (7,  cat_tot["qty"], "#,##0"),
            (8,  cat_tot["rev"], "₹#,##0.00"),
            (9,  cat_tot["cst"], "₹#,##0.00"),
            (10, cat_tot["prf"], "₹#,##0.00"),
            (11, cat_margin,     "0.0%"),
        ]:
            c_ = ws.cell(row, c, val)
            c_.font = SUBTOT_FONT; c_.fill = SUBTOT_FILL
            c_.border = THIN_BORDER; c_.number_format = fmt
            c_.alignment = Alignment(horizontal="center", vertical="center")
        row += 1

    # Grand totals row
    grand_margin = grand["prf"] / grand["rev"] if grand["rev"] > 0 else 0
    tot_label = ws.cell(row, 1, "  GRAND TOTAL")
    tot_label.font = TOTAL_FONT; tot_label.fill = TOTAL_FILL; tot_label.border = THIN_BORDER
    for c in range(2, 7):
        ws.cell(row, c).fill = TOTAL_FILL; ws.cell(row, c).border = THIN_BORDER
    for c, val, fmt in [
        (7,  grand["qty"], "#,##0"),
        (8,  grand["rev"], "₹#,##0.00"),
        (9,  grand["cst"], "₹#,##0.00"),
        (10, grand["prf"], "₹#,##0.00"),
        (11, grand_margin, "0.0%"),
    ]:
        c_ = ws.cell(row, c, val)
        c_.font = TOTAL_FONT; c_.fill = TOTAL_FILL
        c_.border = THIN_BORDER; c_.number_format = fmt
        c_.alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 10
    for col in ["D", "E", "F", "G", "H", "I", "J", "K"]:
        ws.column_dimensions[col].width = 16
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


def _write_demand_analytics(wb):
    sn = "DEMAND_ANALYTICS"
    _safe_delete_sheet(wb, sn)
    ws = wb.create_sheet(sn)

    headers = ["SERVICE", "ROLE", "TOTAL QTY", "TOTAL TX", "LAST 24H",
               "LAST 7D", "LAST 30D", "VELOCITY", "DEMAND SCORE", "DECAY SCORE",
               "CONFIDENCE", "TREND", "STATUS", "PEAK HOUR", "TOP WEEKDAY",
               "ROLLING AVG 7D", "ROLLING AVG 30D"]
    for i, h in enumerate(headers, 1):
        _hdr(ws, 1, i, h)

    demand        = _ensure_demand_cache()
    svc_map       = _get_services()
    weekday_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

    row = 2
    for name in sorted(svc_map.keys()):
        d = demand.get(name, {})
        vel   = d.get("velocity", 0)
        conf  = d.get("confidence", 0)
        trend = d.get("trend", "—")
        h_dist = d.get("hourly_dist", {})
        w_dist = d.get("weekday_dist", {})

        status = ("🔥 Hot"  if vel > 0.3 else
                  ("🟡 Warm" if vel > 0.1 else
                   ("❄️ Cold" if d.get("total_tx", 0) > 0 else "—")))
        fill   = (GREEN_FILL  if vel > 0.3 else
                  (ORANGE_FILL if vel > 0.1 else None))

        peak_hour   = max(h_dist, key=h_dist.get, default="—")
        if peak_hour != "—":
            peak_hour = f"{peak_hour:02d}:00"
        top_weekday = max(w_dist, key=w_dist.get, default="—")
        if top_weekday != "—":
            top_weekday = weekday_names[int(top_weekday)]

        trend_fill = (GREEN_FILL  if trend == "rising"  else
                      (WARN_FILL  if trend == "falling" else None))

        ws.cell(row, 1).value  = name
        ws.cell(row, 1).font   = BODY_FONT
        ws.cell(row, 1).border = THIN_BORDER
        _cell(ws, row, 2,  svc_map[name].get("role", "filler"))
        _cell(ws, row, 3,  d.get("total_qty", 0))
        _cell(ws, row, 4,  d.get("total_tx", 0))
        _cell(ws, row, 5,  d.get("last_24h", 0))
        _cell(ws, row, 6,  d.get("last_7d", 0))
        _cell(ws, row, 7,  d.get("last_30d", 0))
        _cell(ws, row, 8,  vel,  fmt="0.0000")
        _cell(ws, row, 9,  d.get("demand_score", 0), fmt="0.000")
        _cell(ws, row, 10, d.get("decay_score", 0),  fmt="0.000")
        _cell(ws, row, 11, conf, fmt="0.0%")
        _cell(ws, row, 12, trend, fill=trend_fill)
        _cell(ws, row, 13, status, fill=fill)
        _cell(ws, row, 14, str(peak_hour))
        _cell(ws, row, 15, str(top_weekday))
        _cell(ws, row, 16, d.get("rolling_avg_7d", 0),  fmt="0.00")
        _cell(ws, row, 17, d.get("rolling_avg_30d", 0), fmt="0.00")
        row += 1

    ws.column_dimensions["A"].width = 32
    for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J",
                "K", "L", "M", "N", "O", "P", "Q"]:
        ws.column_dimensions[col].width = 14
    ws.freeze_panes = "A2"


def _write_bundle_analytics(wb):
    sn = "BUNDLE_ANALYTICS"
    _safe_delete_sheet(wb, sn)
    ws = wb.create_sheet(sn)

    headers = ["SERVICE A", "SERVICE B", "CO-OCCURRENCE", "STRENGTH", "CONFIDENCE", "RANK"]
    for i, h in enumerate(headers, 1):
        _hdr(ws, 1, i, h)

    stats  = _ensure_bundle_stats()
    ranked = sorted(stats.items(), key=lambda x: -x[1]["count"])

    for rank, ((a, b), data) in enumerate(ranked, 1):
        row  = rank + 1
        fill = (GREEN_FILL  if data["strength"] > 0.05 else
                (ORANGE_FILL if data["strength"] > 0.02 else None))
        ws.cell(row, 1).value  = a
        ws.cell(row, 1).font   = BODY_FONT
        ws.cell(row, 1).border = THIN_BORDER
        ws.cell(row, 2).value  = b
        ws.cell(row, 2).font   = BODY_FONT
        ws.cell(row, 2).border = THIN_BORDER
        _cell(ws, row, 3, data["count"])
        _cell(ws, row, 4, data["strength"],             fmt="0.0000", fill=fill)
        _cell(ws, row, 5, data.get("confidence", 0),   fmt="0.0%")
        _cell(ws, row, 6, rank)
        if rank >= 200:
            break

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 32
    for col in ["C", "D", "E", "F"]:
        ws.column_dimensions[col].width = 18
    ws.freeze_panes = "A2"


def _write_risk_dashboard(wb):
    sn = "RISK_DASHBOARD"
    _safe_delete_sheet(wb, sn)
    ws = wb.create_sheet(sn)

    headers = ["SERVICE", "CATEGORY", "ROLE", "PRICE ₹", "COST ₹", "MARGIN %",
               "RISK LEVEL", "ISSUE"]
    for i, h in enumerate(headers, 1):
        _hdr(ws, 1, i, h)

    svc_map = _get_services()
    risks   = []
    for name, svc in svc_map.items():
        alerts = _get_risk_alerts(svc)
        if alerts:
            risks.append((name, svc, alerts))

    risks.sort(key=lambda x: {"critical": 0, "warning": 1, "info": 2}
               .get(x[2][0]["severity"], 3))

    row = 2
    for name, svc, alerts in risks:
        top   = alerts[0]
        sev   = top["severity"]
        fill  = (WARN_FILL   if sev == "critical" else
                 (ORANGE_FILL if sev == "warning"  else None))
        label = ("🔴 CRITICAL" if sev == "critical" else
                 ("🟠 WARNING" if sev == "warning"  else "🔵 INFO"))

        ws.cell(row, 1).value  = name
        ws.cell(row, 1).font   = BODY_FONT
        ws.cell(row, 1).border = THIN_BORDER
        _cell(ws, row, 2, svc["category"])
        _cell(ws, row, 3, svc.get("role", "filler"))
        _cell(ws, row, 4, svc["price"],  fmt="₹#,##0.00")
        _cell(ws, row, 5, svc["cost"],   fmt="₹#,##0.00")
        _cell(ws, row, 6, svc["margin_pct"] / 100, fmt="0.0%", fill=fill)
        _cell(ws, row, 7, label, fill=fill)
        _cell(ws, row, 8, top["message"])
        row += 1

    ws.column_dimensions["A"].width = 32
    for col in ["B", "C", "D", "E", "F", "G", "H"]:
        ws.column_dimensions[col].width = 18
    ws.freeze_panes = "A2"


def _write_ai_pricing_sheet(wb):
    sn = "AI_PRICING"
    _safe_delete_sheet(wb, sn)
    ws = wb.create_sheet(sn)

    headers = ["SERVICE", "ROLE", "CURRENT ₹", "SUGGESTED ₹", "DELTA %",
               "CONFIDENCE", "REASON", "PROFIT EST.", "ACTION", "ACCEPT RATE",
               "COOLDOWN", "TREND", "ELASTICITY SRC"]
    for i, h in enumerate(headers, 1):
        _hdr(ws, 1, i, h)

    ai_map = _ensure_ai_cache()
    demand = _ensure_demand_cache()
    items  = sorted(ai_map.values(), key=lambda x: -x["confidence"])

    row = 2
    for ins in items:
        if ins["confidence"] < 0.20:
            continue
        delta      = ins["delta_pct"]
        fill       = (GREEN_FILL  if delta > 0 else
                      (ORANGE_FILL if delta < 0 else None))
        action     = "↑ Raise" if delta > 0 else ("↓ Lower" if delta < 0 else "✓ OK")
        accept     = ins.get("ai_accept_rate")
        accept_str = f"{accept:.0%}" if accept is not None else "—"
        pm         = ins.get("price_memory", {})
        cooldown   = "⏸ Yes" if pm.get("in_cooldown") else "—"
        trend      = demand.get(ins["service"], {}).get("trend", "—")

        # Show elasticity source
        with _price_mem_lock:
            n_events = len(_price_memory.get(ins["service"], {}).get("price_events", []))
        elast_src = f"learned({n_events})" if n_events >= LEARNED_ELASTICITY_MIN else "heuristic"

        ws.cell(row, 1).value  = ins["service"]
        ws.cell(row, 1).font   = BODY_FONT
        ws.cell(row, 1).border = THIN_BORDER
        _cell(ws, row, 2,  ins.get("role", "filler"))
        _cell(ws, row, 3,  ins["current_price"],   fmt="₹#,##0.00")
        _cell(ws, row, 4,  ins["suggested_price"], fmt="₹#,##0.00", fill=fill)
        _cell(ws, row, 5,  delta / 100,             fmt="+0.0%;-0.0%;0.0%")
        _cell(ws, row, 6,  ins["confidence"],       fmt="0%")
        _cell(ws, row, 7,  ins["reason"])
        _cell(ws, row, 8,  ins.get("profit_estimate", 0), fmt="₹#,##0.00")
        _cell(ws, row, 9,  action, fill=fill)
        _cell(ws, row, 10, accept_str)
        _cell(ws, row, 11, cooldown)
        _cell(ws, row, 12, trend)
        _cell(ws, row, 13, elast_src)
        row += 1

    ws.column_dimensions["A"].width = 32
    for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]:
        ws.column_dimensions[col].width = 15
    ws.freeze_panes = "A2"


def _write_ai_control_log(wb):
    """
    FIX BUG-4: always receives wb from caller — never opens its own workbook.
    All writes go through the single _write_analytics_sheets pipeline.
    """
    sn = "AI_CONTROL_LOG"
    _safe_delete_sheet(wb, sn)
    ws = wb.create_sheet(sn)

    headers = ["SERVICE", "ACCEPTED", "IGNORED", "TOTAL", "ACCEPT RATE", "TRUST LEVEL"]
    for i, h in enumerate(headers, 1):
        _hdr(ws, 1, i, h)

    with _ctrl_lock:
        log_copy = dict(_ctrl_log)

    row = 2
    for service, d in sorted(log_copy.items()):
        acc   = d.get("accepted", 0)
        ign   = d.get("ignored", 0)
        total = acc + ign
        rate  = round(acc / total, 3) if total > 0 else 0
        trust = ("High" if rate > 0.7 else ("Low" if rate < 0.3 else "Medium"))
        fill  = (GREEN_FILL  if rate > 0.7 else
                 (ORANGE_FILL if rate < 0.3 else None))
        ws.cell(row, 1).value  = service
        ws.cell(row, 1).font   = BODY_FONT
        ws.cell(row, 1).border = THIN_BORDER
        _cell(ws, row, 2, acc)
        _cell(ws, row, 3, ign)
        _cell(ws, row, 4, total)
        _cell(ws, row, 5, rate,  fmt="0%", fill=fill)
        _cell(ws, row, 6, trust, fill=fill)
        row += 1

    ws.column_dimensions["A"].width = 32
    for col in ["B", "C", "D", "E", "F"]:
        ws.column_dimensions[col].width = 16
    ws.freeze_panes = "A2"


def _write_price_memory_sheet(wb):
    sn = "PRICE_MEMORY"
    _safe_delete_sheet(wb, sn)
    ws = wb.create_sheet(sn)

    headers = ["SERVICE", "LAST PRICE ₹", "LAST UPDATE", "CHANGE COUNT",
               "DIRECTION", "COOLDOWN ACTIVE", "PRICE EVENTS", "LAST 5 CHANGES"]
    for i, h in enumerate(headers, 1):
        _hdr(ws, 1, i, h)

    now = time.time()
    with _price_mem_lock:
        pm_copy = {k: dict(v) for k, v in _price_memory.items()}

    row = 2
    for service, pm in sorted(pm_copy.items()):
        lu     = pm.get("last_update", 0)
        lu_str = datetime.fromtimestamp(lu).strftime("%d %b %H:%M") if lu else "—"
        in_cd  = now < pm.get("cooldown_until", 0)
        hist   = pm.get("history", [])[-5:]
        hist_str = " → ".join(f"₹{h['from']}→₹{h['to']}" for h in hist)
        n_events = len(pm.get("price_events", []))

        ws.cell(row, 1).value  = service
        ws.cell(row, 1).font   = BODY_FONT
        ws.cell(row, 1).border = THIN_BORDER
        _cell(ws, row, 2, pm.get("last_price", 0), fmt="₹#,##0.00")
        _cell(ws, row, 3, lu_str)
        _cell(ws, row, 4, pm.get("change_count", 0))
        _cell(ws, row, 5, pm.get("direction", "none"))
        _cell(ws, row, 6, "⏸ YES" if in_cd else "—",
              fill=(ORANGE_FILL if in_cd else None))
        _cell(ws, row, 7, n_events)
        _cell(ws, row, 8, hist_str if hist_str else "—")
        row += 1

    ws.column_dimensions["A"].width = 32
    for col in ["B", "C", "D", "E", "F", "G", "H"]:
        ws.column_dimensions[col].width = 20
    ws.freeze_panes = "A2"


def _write_price_history_from_db(wb):
    """
    DB-FIRST v5: Populate PRICE_HISTORY sheet from DB price_events table.
    Always has data — not dependent on in-memory price_memory state.
    """
    sn = "PRICE_HISTORY"
    _safe_delete_sheet(wb, sn)
    ws = wb.create_sheet(sn)

    headers = ["SERVICE", "TIMESTAMP", "OLD PRICE", "NEW PRICE", "SOURCE", "CHANGE ₹", "CHANGE %"]
    for i, h in enumerate(headers, 1):
        _hdr(ws, 1, i, h)

    if not _DB_AVAILABLE:
        ws.cell(row=2, column=1).value = "DB not available"
        return

    try:
        conn = _db.get_db()
        rows = conn.execute(
            "SELECT service_name, old_price, new_price, timestamp, source "
            "FROM price_events ORDER BY id DESC LIMIT 500"
        ).fetchall()
    except Exception as e:
        log.warning("_write_price_history_from_db query failed: %s", e)
        return

    for r_idx, r in enumerate(rows, 2):
        old_p  = float(r["old_price"] or 0)
        new_p  = float(r["new_price"] or 0)
        change = round(new_p - old_p, 4)
        pct    = round((new_p - old_p) / old_p, 4) if old_p > 0 else 0
        fill   = (GREEN_FILL if change > 0 else (ORANGE_FILL if change < 0 else None))

        ws.cell(r_idx, 1).value  = r["service_name"]
        ws.cell(r_idx, 1).font   = BODY_FONT
        ws.cell(r_idx, 1).border = THIN_BORDER
        _cell(ws, r_idx, 2, str(r["timestamp"])[:19])
        _cell(ws, r_idx, 3, old_p,   fmt="₹#,##0.00")
        _cell(ws, r_idx, 4, new_p,   fmt="₹#,##0.00", fill=fill)
        _cell(ws, r_idx, 5, r["source"] or "manual")
        _cell(ws, r_idx, 6, change,  fmt="+₹#,##0.00;-₹#,##0.00", fill=fill)
        _cell(ws, r_idx, 7, pct,     fmt="+0.0%;-0.0%", fill=fill)

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 20
    for col in ["C", "D", "E", "F", "G"]:
        ws.column_dimensions[col].width = 16
    ws.freeze_panes = "A2"
    log.info("PRICE_HISTORY written from DB: %d rows", len(rows))


def _write_dashboard(wb):
    """
    Top-level analytical dashboard. Real data from DB. No decorative filler.
    Sections:
      A. Business KPIs (revenue, profit, margin, transactions) — today / 7d / 30d
      B. Top 10 services by revenue (all time)
      C. Top 10 services by profit (all time)
      D. Payment method split
      E. Daily revenue last 14 days
      F. Category performance (revenue, profit, tx count, margin)
      G. Risk summary (zero-price, negative-margin, low-margin counts)
    """
    sn = "📊 DASHBOARD"
    if sn in wb.sheetnames:
        del wb[sn]
    ws = wb.create_sheet(sn, 0)   # first sheet

    # ── colour palette (reuse module constants) ──
    TITLE_FILL  = PatternFill("solid", start_color="0D1117")
    SECT_FILL   = PatternFill("solid", start_color="1F2937")
    VAL_FONT    = Font(name="Arial", size=11, bold=True, color="F0F6FC")
    TITLE_FONT  = Font(name="Arial", size=14, bold=True, color="00D4FF")
    SECT_FONT   = Font(name="Arial", size=9,  bold=True, color="8B949E")
    KPI_FONT    = Font(name="Arial", size=18, bold=True, color="00D4FF")
    KPI_G_FONT  = Font(name="Arial", size=18, bold=True, color="39D353")
    KPI_A_FONT  = Font(name="Arial", size=18, bold=True, color="F59E0B")
    SUB_FONT    = Font(name="Arial", size=9,  color="8B949E")

    def title_cell(r, c, val, font=None, fill=None):
        cell = ws.cell(r, c, val)
        cell.font      = font or TITLE_FONT
        cell.fill      = fill or TITLE_FILL
        cell.alignment = Alignment(horizontal="left", vertical="center")
        return cell

    def sect(r, c, val):
        cell = ws.cell(r, c, val)
        cell.font      = SECT_FONT
        cell.fill      = SECT_FILL
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        return cell

    def kpi(r, c, val, fmt=None, font=None, fill=None):
        cell = ws.cell(r, c, val)
        cell.font      = font or KPI_FONT
        cell.fill      = fill or TITLE_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = THIN_BORDER
        if fmt: cell.number_format = fmt
        return cell

    def tbl_hdr(r, c, val):
        cell = ws.cell(r, c, val)
        cell.font      = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        cell.fill      = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = THIN_BORDER
        return cell

    def tbl_val(r, c, val, fmt=None, fill=None, bold=False):
        cell = ws.cell(r, c, val)
        cell.font      = Font(name="Arial", size=9, bold=bold, color="F0F6FC")
        cell.fill      = fill or PatternFill("solid", start_color="161B22")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = THIN_BORDER
        if fmt: cell.number_format = fmt
        return cell

    def tbl_lbl(r, c, val):
        cell = ws.cell(r, c, val)
        cell.font      = Font(name="Arial", size=9, color="F0F6FC")
        cell.fill      = PatternFill("solid", start_color="161B22")
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell.border    = THIN_BORDER
        return cell

    # ── Pull data from DB ─────────────────────────────────────────────────────
    today     = date.today().isoformat()
    day7_str  = (date.today() - timedelta(days=7)).isoformat()
    day30_str = (date.today() - timedelta(days=30)).isoformat()

    def db_period(start: str):
        try:
            conn = _db.get_db()
            r = conn.execute(
                "SELECT COALESCE(SUM(revenue),0), COALESCE(SUM(profit),0), COUNT(*) "
                "FROM transactions WHERE timestamp >= ?", (start,)
            ).fetchone()
            rev, prf, cnt = float(r[0]), float(r[1]), int(r[2])
            margin = round(prf/rev*100, 1) if rev > 0 else 0.0
            avg    = round(rev/cnt, 2)    if cnt > 0 else 0.0
            return rev, prf, cnt, margin, avg
        except Exception:
            return 0.0, 0.0, 0, 0.0, 0.0

    t_rev, t_prf, t_cnt, t_margin, t_avg = db_period(today)
    w_rev, w_prf, w_cnt, w_margin, w_avg = db_period(day7_str)
    m_rev, m_prf, m_cnt, m_margin, m_avg = db_period(day30_str)

    # Top services
    def top_services(order_col: str, limit=10):
        try:
            conn = _db.get_db()
            rows = conn.execute(
                f"SELECT service_name, SUM(qty) qty, SUM(revenue) rev, SUM(profit) prf "
                f"FROM transactions GROUP BY service_name ORDER BY {order_col} DESC LIMIT ?",
                (limit,)
            ).fetchall()
            return [(r["service_name"], int(r["qty"] or 0),
                     float(r["rev"] or 0), float(r["prf"] or 0)) for r in rows]
        except Exception:
            return []

    top_rev  = top_services("SUM(revenue)")
    top_prf  = top_services("SUM(profit)")

    # Payment split
    def payment_split():
        try:
            conn = _db.get_db()
            rows = conn.execute(
                "SELECT payment_mode, COUNT(*) cnt, COALESCE(SUM(revenue),0) rev "
                "FROM transactions GROUP BY payment_mode ORDER BY rev DESC"
            ).fetchall()
            return [(r["payment_mode"], int(r["cnt"]), float(r["rev"])) for r in rows]
        except Exception:
            return []

    pay_data = payment_split()

    # Daily last 14 days
    def daily_14():
        try:
            conn = _db.get_db()
            rows = conn.execute(
                "SELECT substr(timestamp,1,10) day, "
                "COALESCE(SUM(revenue),0) rev, COALESCE(SUM(profit),0) prf, COUNT(*) cnt "
                "FROM transactions "
                "WHERE timestamp >= date('now','-14 days') "
                "GROUP BY day ORDER BY day DESC"
            ).fetchall()
            return [(r["day"], float(r["rev"]), float(r["prf"]), int(r["cnt"])) for r in rows]
        except Exception:
            return []

    daily = daily_14()

    # Category performance
    def cat_perf():
        try:
            conn = _db.get_db()
            rows = conn.execute(
                "SELECT s.category, "
                "COALESCE(SUM(t.revenue),0) rev, COALESCE(SUM(t.profit),0) prf, "
                "COUNT(t.id) cnt "
                "FROM transactions t "
                "JOIN services s ON t.service_name = s.name "
                "GROUP BY s.category ORDER BY rev DESC"
            ).fetchall()
            return [(r["category"], float(r["rev"]), float(r["prf"]), int(r["cnt"])) for r in rows]
        except Exception:
            return []

    cat_data = cat_perf()

    # Risk counts
    svc_map    = _get_services()
    zero_p     = sum(1 for s in svc_map.values() if s["price"]==0 and s["name"]!="Other")
    neg_margin = sum(1 for s in svc_map.values() if s["price"]>0 and s["cost"]>s["price"])
    low_margin = sum(1 for s in svc_map.values() if 0<s["price"] and 0<(s["price"]-s["cost"])/s["price"]<0.10 and s["name"]!="Other")

    # ── LAYOUT ───────────────────────────────────────────────────────────────
    ROW = 1

    # Title
    ws.row_dimensions[ROW].height = 32
    title_cell(ROW, 1, "⚡  CITYCYBER — OPERATIONS DASHBOARD", TITLE_FONT, TITLE_FILL)
    ws.merge_cells(start_row=ROW, start_column=1, end_row=ROW, end_column=14)
    ROW += 1

    sub = ws.cell(ROW, 1, f"Generated: {datetime.now().strftime('%d %b %Y  %H:%M')}   |   DB-first · real-time analytics")
    sub.font      = SUB_FONT
    sub.fill      = TITLE_FILL
    sub.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(start_row=ROW, start_column=1, end_row=ROW, end_column=14)
    ws.row_dimensions[ROW].height = 16
    ROW += 2

    # ── Section A: KPIs ───────────────────────────────────────────────────────
    sect(ROW, 1, "A.  BUSINESS PERFORMANCE")
    ws.merge_cells(start_row=ROW, start_column=1, end_row=ROW, end_column=14)
    ROW += 1

    # KPI column headers
    for c, lbl in enumerate(["", "TODAY", "LAST 7 DAYS", "LAST 30 DAYS"], 1):
        h = ws.cell(ROW, c, lbl)
        h.font = Font(name="Arial", size=9, bold=True, color="8B949E")
        h.fill = SECT_FILL
        h.alignment = Alignment(horizontal="center" if c>1 else "left", vertical="center", indent=1 if c==1 else 0)
        h.border = THIN_BORDER
    ROW += 1

    kpi_rows = [
        ("Revenue ₹",     f"₹#,##0.00", KPI_FONT,   t_rev,    w_rev,    m_rev),
        ("Profit ₹",      f"₹#,##0.00", KPI_G_FONT,  t_prf,    w_prf,    m_prf),
        ("Transactions",  "#,##0",      KPI_A_FONT,  t_cnt,    w_cnt,    m_cnt),
        ("Avg Ticket ₹",  "₹#,##0.00", KPI_FONT,   t_avg,    w_avg,    m_avg),
        ("Margin %",      "0.0%",       KPI_G_FONT,  t_margin/100, w_margin/100, m_margin/100),
    ]
    for label, fmt, font, v1, v2, v3 in kpi_rows:
        ws.row_dimensions[ROW].height = 28
        lbl_c = ws.cell(ROW, 1, label)
        lbl_c.font = Font(name="Arial", size=10, bold=True, color="F0F6FC")
        lbl_c.fill = TITLE_FILL
        lbl_c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        lbl_c.border = THIN_BORDER
        for c, val in enumerate([v1, v2, v3], 2):
            cell = ws.cell(ROW, c, val)
            cell.font = font
            cell.fill = TITLE_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER
            cell.number_format = fmt
        ROW += 1

    ROW += 1

    # ── Sections B & C: Top services — side by side ───────────────────────────
    sect(ROW, 1, "B.  TOP 10 SERVICES BY REVENUE (ALL TIME)")
    ws.merge_cells(start_row=ROW, start_column=1, end_row=ROW, end_column=6)
    sect(ROW, 8, "C.  TOP 10 SERVICES BY PROFIT (ALL TIME)")
    ws.merge_cells(start_row=ROW, start_column=8, end_row=ROW, end_column=13)
    ROW += 1

    for c, h in enumerate(["SERVICE", "QTY", "REVENUE ₹", "PROFIT ₹", "MARGIN %", "RANK"], 1):
        tbl_hdr(ROW, c, h)
    for c, h in enumerate(["SERVICE", "QTY", "REVENUE ₹", "PROFIT ₹", "MARGIN %", "RANK"], 8):
        tbl_hdr(ROW, c, h)
    ROW += 1

    for rank in range(10):
        alt = PatternFill("solid", start_color="0D1117") if rank % 2 == 0 else PatternFill("solid", start_color="161B22")
        # Revenue table
        if rank < len(top_rev):
            name, qty, rev, prf = top_rev[rank]
            marg = prf/rev if rev > 0 else 0
            fill = GREEN_FILL if rank == 0 else alt
            tbl_lbl(ROW, 1, name[:30])
            ws.cell(ROW,1).fill = fill
            tbl_val(ROW, 2, qty,  "#,##0",    fill)
            tbl_val(ROW, 3, rev,  "₹#,##0.00", fill, bold=(rank==0))
            tbl_val(ROW, 4, prf,  "₹#,##0.00", fill)
            tbl_val(ROW, 5, marg, "0.0%",      fill)
            tbl_val(ROW, 6, rank+1, "#",        fill, bold=(rank==0))
        # Profit table
        if rank < len(top_prf):
            name2, qty2, rev2, prf2 = top_prf[rank]
            marg2 = prf2/rev2 if rev2 > 0 else 0
            fill2 = GREEN_FILL if rank == 0 else alt
            tbl_lbl(ROW, 8, name2[:30])
            ws.cell(ROW,8).fill = fill2
            tbl_val(ROW, 9,  qty2,  "#,##0",     fill2)
            tbl_val(ROW, 10, rev2,  "₹#,##0.00", fill2)
            tbl_val(ROW, 11, prf2,  "₹#,##0.00", fill2, bold=(rank==0))
            tbl_val(ROW, 12, marg2, "0.0%",       fill2)
            tbl_val(ROW, 13, rank+1, "#",          fill2, bold=(rank==0))
        ROW += 1

    ROW += 1

    # ── Section D: Payment split ──────────────────────────────────────────────
    sect(ROW, 1, "D.  PAYMENT METHOD SPLIT")
    ws.merge_cells(start_row=ROW, start_column=1, end_row=ROW, end_column=5)
    ROW += 1
    for c, h in enumerate(["PAYMENT MODE", "TRANSACTIONS", "REVENUE ₹", "SHARE %", ""], 1):
        tbl_hdr(ROW, c, h)
    ROW += 1
    total_pay_rev = sum(r[2] for r in pay_data) or 1
    for mode, cnt, rev in pay_data:
        share = rev / total_pay_rev
        fill  = (PatternFill("solid", start_color="0D3320") if mode=="Cash" else
                 PatternFill("solid", start_color="001A2E") if mode=="UPI" else
                 PatternFill("solid", start_color="1A0D38"))
        tbl_lbl(ROW, 1, mode)
        ws.cell(ROW,1).fill = fill
        tbl_val(ROW, 2, cnt,   "#,##0",    fill)
        tbl_val(ROW, 3, rev,   "₹#,##0.00", fill)
        tbl_val(ROW, 4, share, "0.0%",      fill, bold=True)
        ROW += 1
    ROW += 1

    # ── Section E: Daily trend last 14 days ──────────────────────────────────
    sect(ROW, 1, "E.  DAILY REVENUE — LAST 14 DAYS")
    ws.merge_cells(start_row=ROW, start_column=1, end_row=ROW, end_column=6)
    ROW += 1
    for c, h in enumerate(["DATE", "REVENUE ₹", "PROFIT ₹", "TRANSACTIONS", "AVG TICKET ₹", "MARGIN %"], 1):
        tbl_hdr(ROW, c, h)
    ROW += 1
    max_daily_rev = max((r[1] for r in daily), default=1) or 1
    for day_str, rev, prf, cnt in daily:
        bar_pct = rev / max_daily_rev
        fill = (GREEN_FILL  if bar_pct > 0.7 else
                ORANGE_FILL if bar_pct > 0.3 else
                PatternFill("solid", start_color="161B22"))
        avg_t  = rev/cnt if cnt > 0 else 0
        margin = prf/rev  if rev > 0 else 0
        tbl_lbl(ROW, 1, day_str)
        ws.cell(ROW,1).fill = fill
        tbl_val(ROW, 2, rev,    "₹#,##0.00", fill, bold=True)
        tbl_val(ROW, 3, prf,    "₹#,##0.00", fill)
        tbl_val(ROW, 4, cnt,    "#,##0",      fill)
        tbl_val(ROW, 5, avg_t,  "₹#,##0.00", fill)
        tbl_val(ROW, 6, margin, "0.0%",       fill)
        ROW += 1
    ROW += 1

    # ── Section F: Category performance ──────────────────────────────────────
    sect(ROW, 1, "F.  CATEGORY PERFORMANCE (ALL TIME)")
    ws.merge_cells(start_row=ROW, start_column=1, end_row=ROW, end_column=6)
    ROW += 1
    for c, h in enumerate(["CATEGORY", "REVENUE ₹", "PROFIT ₹", "TRANSACTIONS", "MARGIN %", "AVG TICKET ₹"], 1):
        tbl_hdr(ROW, c, h)
    ROW += 1
    for i, (cat, rev, prf, cnt) in enumerate(cat_data):
        margin = prf/rev if rev > 0 else 0
        avg_t  = rev/cnt if cnt > 0 else 0
        fill   = (PatternFill("solid", start_color="0D1117") if i%2==0
                  else PatternFill("solid", start_color="161B22"))
        tbl_lbl(ROW, 1, cat or "Unknown")
        ws.cell(ROW,1).fill = fill
        tbl_val(ROW, 2, rev,    "₹#,##0.00", fill)
        tbl_val(ROW, 3, prf,    "₹#,##0.00", fill)
        tbl_val(ROW, 4, cnt,    "#,##0",      fill)
        tbl_val(ROW, 5, margin, "0.0%",
                GREEN_FILL if margin > 0.4 else ORANGE_FILL if margin < 0.15 else fill)
        tbl_val(ROW, 6, avg_t,  "₹#,##0.00", fill)
        ROW += 1
    ROW += 1

    # ── Section G: Risk summary ───────────────────────────────────────────────
    sect(ROW, 1, "G.  RISK SUMMARY")
    ws.merge_cells(start_row=ROW, start_column=1, end_row=ROW, end_column=4)
    ROW += 1
    risks = [
        ("Services with ZERO price",     zero_p,     WARN_FILL   if zero_p     > 0 else None),
        ("Services with NEGATIVE margin", neg_margin, WARN_FILL   if neg_margin > 0 else None),
        ("Services with margin < 10%",   low_margin, ORANGE_FILL if low_margin > 0 else None),
        ("Total services in catalogue",  len(svc_map), None),
        ("DB available",                 "YES" if _DB_AVAILABLE else "NO",
                                          None if _DB_AVAILABLE else WARN_FILL),
    ]
    for label, val, fill in risks:
        tbl_lbl(ROW, 1, label)
        ws.cell(ROW, 1).fill = fill or PatternFill("solid", start_color="161B22")
        v_cell = tbl_val(ROW, 2, val, fill=fill or PatternFill("solid", start_color="161B22"), bold=True)
        if isinstance(val, int) and val > 0 and fill == WARN_FILL:
            v_cell.font = Font(name="Arial", size=9, bold=True, color="FF6B6B")
        ROW += 1

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 34
    for col in ["B","C","D","E","F"]:
        ws.column_dimensions[col].width = 16
    ws.column_dimensions["G"].width = 4   # spacer
    ws.column_dimensions["H"].width = 34
    for col in ["I","J","K","L","M"]:
        ws.column_dimensions[col].width = 16
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"
    log.info("DASHBOARD written: %d KPI rows, %d top-rev, %d days trend, %d categories",
             len(kpi_rows), len(top_rev), len(daily), len(cat_data))


def _write_cashflow_sheet(wb):
    """
    Write CASHFLOW_INTELLIGENCE sheet: collected vs outstanding, top debtors,
    daily breakdown, risk distribution.
    """
    sn = "CASHFLOW_INTELLIGENCE"
    _safe_delete_sheet(wb, sn)
    ws = wb.create_sheet(sn)

    CLR_BG      = "0D1117"
    CLR_SURF    = "161B22"
    CLR_SURF2   = "21262D"
    CLR_TEXT    = "E6EDF3"
    CLR_MUTED   = "8B949E"
    CLR_GREEN   = "39D353"
    CLR_RED     = "F85149"
    CLR_AMBER   = "E3B341"
    CLR_BLUE    = "58A6FF"
    CLR_UDHAAR  = "E6A817"
    FMT_RS      = "₹#,##0.00"
    FMT_PCT     = "0.0%"

    def _fill(h):
        from openpyxl.styles import PatternFill
        return PatternFill("solid", fgColor=h)

    def _font(color=CLR_TEXT, bold=False, size=10):
        return Font(name="Segoe UI", size=size, bold=bold, color=color)

    def _align(h="left", v="center"):
        return Alignment(horizontal=h, vertical=v, wrap_text=False)

    def _bdr():
        t = Side(style="thin", color="30363D")
        return Border(left=t, right=t, top=t, bottom=t)

    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width  = 2
    ws.column_dimensions["B"].width  = 22
    ws.column_dimensions["C"].width  = 18
    ws.column_dimensions["D"].width  = 18
    ws.column_dimensions["E"].width  = 18
    ws.column_dimensions["F"].width  = 18
    ws.column_dimensions["G"].width  = 18
    ws.column_dimensions["H"].width  = 12
    ws.column_dimensions["I"].width  = 12

    if not _DB_AVAILABLE:
        ws.cell(3, 2).value = "Database unavailable"
        return

    try:
        cf      = _db.get_cashflow_metrics()
        debtors = _db.get_top_debtors(limit=20)
        overdue = _db.get_overdue_entries(days=7)
        summary = _db.get_udhaar_summary()
    except Exception as _ce:
        log.warning("_write_cashflow_sheet: DB read failed: %s", _ce)
        return

    # ── Title ──
    R = 2
    ws.row_dimensions[R].height = 32
    ws.merge_cells(f"B{R}:I{R}")
    c = ws.cell(R, 2)
    c.value = "⚡  CITYCYBER — CASHFLOW INTELLIGENCE"
    c.font  = _font(CLR_UDHAAR, bold=True, size=15)
    c.fill  = _fill(CLR_SURF)
    c.alignment = _align()

    R += 1
    ws.row_dimensions[R].height = 14
    ws.merge_cells(f"B{R}:I{R}")
    c = ws.cell(R, 2)
    c.value = f"Snapshot: {cf.get('today','')}  ·  Collected vs Booked vs Outstanding"
    c.font  = _font(CLR_MUTED, size=9)
    c.fill  = _fill(CLR_SURF)
    c.alignment = _align()

    # ── KPI row ──
    R += 2
    ws.row_dimensions[R].height = 14
    kpis = [
        ("BOOKED REVENUE",    cf.get("booked_revenue",   0), CLR_TEXT,   FMT_RS),
        ("COLLECTED CASH",    cf.get("collected_revenue", 0), CLR_GREEN,  FMT_RS),
        ("UDHAAR REVENUE",    cf.get("udhaar_revenue",   0), CLR_AMBER,  FMT_RS),
        ("OUTSTANDING",       cf.get("outstanding_udhaar",0), CLR_RED,   FMT_RS),
        ("REAL PROFIT",       cf.get("real_profit",      0), CLR_BLUE,   FMT_RS),
        ("BLOCKED PROFIT",    cf.get("blocked_profit",   0), CLR_AMBER,  FMT_RS),
        ("COLLECTION RATE",   cf.get("collection_rate",  1), CLR_GREEN if cf.get("collection_rate",1) >= 0.9 else CLR_AMBER, FMT_PCT),
        ("ACTIVE DEBTORS",    cf.get("active_debtors",   0), CLR_RED if cf.get("active_debtors",0) > 0 else CLR_GREEN, "#,##0"),
    ]
    for ci, (lbl, val, clr, fmt) in enumerate(kpis, 2):
        if ci > 9: break
        ws.row_dimensions[R].height = 13
        lc = ws.cell(R, ci)
        lc.value = lbl
        lc.font  = _font(CLR_MUTED, size=8)
        lc.fill  = _fill(CLR_SURF2)
        lc.alignment = _align("center")
        lc.border = _bdr()

        ws.row_dimensions[R+1].height = 22
        vc = ws.cell(R+1, ci)
        vc.value  = val
        vc.font   = _font(clr, bold=True, size=12)
        vc.fill   = _fill(CLR_SURF)
        vc.number_format = fmt
        vc.alignment = _align("center")
        vc.border = _bdr()

    R += 3

    # ── Top Debtors table ──
    R += 1
    ws.row_dimensions[R].height = 20
    ws.merge_cells(f"B{R}:I{R}")
    c = ws.cell(R, 2)
    c.value = "TOP DEBTORS"
    c.font  = _font(CLR_UDHAAR, bold=True, size=11)
    c.fill  = _fill(CLR_SURF)
    c.alignment = _align()

    R += 1
    hdrs = ["PHONE","NAME","OUTSTANDING","TOTAL GIVEN","RECOVERED","RISK","AGE (DAYS)","LAST ENTRY"]
    ws.row_dimensions[R].height = 15
    for ci, h in enumerate(hdrs, 2):
        c = ws.cell(R, ci)
        c.value = h
        c.font  = _font(CLR_MUTED, bold=True, size=9)
        c.fill  = _fill(CLR_SURF2)
        c.border = _bdr()
        c.alignment = _align("center")

    for d in debtors:
        R += 1
        ws.row_dimensions[R].height = 14
        risk_clr = CLR_RED if d.get("risk") == "high" else (CLR_AMBER if d.get("risk") == "medium" else CLR_GREEN)
        row_data = [
            (d.get("phone",""),          CLR_TEXT),
            (d.get("name") or "—",       CLR_MUTED),
            (round(d.get("balance",0),2), CLR_RED),
            (round(d.get("total_debit",0),2), CLR_TEXT),
            (round(d.get("total_credit",0),2), CLR_GREEN),
            (str(d.get("risk","low")).upper(), risk_clr),
            (d.get("age_days",0),         CLR_MUTED),
            (str(d.get("last_entry_ts",""))[:16], CLR_MUTED),
        ]
        for ci, (val, clr) in enumerate(row_data, 2):
            c = ws.cell(R, ci)
            c.value = val
            c.font  = _font(clr, size=9)
            c.fill  = _fill("1A2233" if R % 2 == 0 else CLR_SURF2)
            c.border = _bdr()
            c.alignment = _align("center" if ci > 3 else "left")
            if isinstance(val, float) and ci in (4,5,6):
                c.number_format = FMT_RS

    # ── Overdue summary ──
    if overdue:
        R += 2
        ws.row_dimensions[R].height = 20
        ws.merge_cells(f"B{R}:I{R}")
        c = ws.cell(R, 2)
        c.value = f"OVERDUE (> 7 DAYS) — {len(overdue)} customer(s)"
        c.font  = _font(CLR_RED, bold=True, size=10)
        c.fill  = _fill("3D1F1F")
        c.alignment = _align()
        R += 1
        for d in overdue[:10]:
            ws.row_dimensions[R].height = 13
            c2 = ws.cell(R, 2); c2.value = d.get("customer_phone",""); c2.font = _font(CLR_RED, size=9); c2.fill = _fill(CLR_SURF2); c2.border = _bdr()
            c3 = ws.cell(R, 3); c3.value = d.get("name") or "—";       c3.font = _font(CLR_MUTED, size=9); c3.fill = _fill(CLR_SURF2); c3.border = _bdr()
            c4 = ws.cell(R, 4); c4.value = round(float(d.get("balance",0)),2); c4.font = _font(CLR_RED, bold=True, size=9); c4.fill = _fill(CLR_SURF2); c4.border = _bdr(); c4.number_format = FMT_RS
            c5 = ws.cell(R, 5); c5.value = str(d.get("oldest_entry",""))[:10]; c5.font = _font(CLR_MUTED, size=9); c5.fill = _fill(CLR_SURF2); c5.border = _bdr()
            c6 = ws.cell(R, 6); c6.value = str(d.get("risk_level","low")).upper(); c6.font = _font(CLR_RED if d.get("risk_level")=="high" else CLR_AMBER, size=9); c6.fill = _fill(CLR_SURF2); c6.border = _bdr()
            R += 1

    ws.freeze_panes = "B6"
    log.info("_write_cashflow_sheet: wrote %d debtors, %d overdue", len(debtors), len(overdue))


def _write_system_events_sheet(wb):
    """Write SYSTEM_EVENTS audit trail sheet from DB."""
    sn = "SYSTEM_EVENTS"
    _safe_delete_sheet(wb, sn)
    ws = wb.create_sheet(sn)

    CLR_SURF2  = "21262D"; CLR_TEXT = "E6EDF3"; CLR_MUTED = "8B949E"
    CLR_GREEN  = "39D353"; CLR_RED  = "F85149"; CLR_AMBER = "E3B341"; CLR_BLUE = "58A6FF"

    def _fill(h):
        from openpyxl.styles import PatternFill
        return PatternFill("solid", fgColor=h)
    def _font(color=CLR_TEXT, bold=False, size=9):
        return Font(name="Segoe UI", size=size, bold=bold, color=color)
    def _bdr():
        t = Side(style="thin", color="30363D")
        return Border(left=t, right=t, top=t, bottom=t)

    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHI", [2,16,14,20,20,20,16,12,2]):
        ws.column_dimensions[col].width = w

    R = 2
    ws.merge_cells(f"B{R}:H{R}")
    c = ws.cell(R, 2); c.value = "SYSTEM AUDIT TRAIL"; c.font = _font(CLR_BLUE, bold=True, size=13)
    c.fill = _fill("161B22"); c.alignment = Alignment(horizontal="left", vertical="center")

    R += 2
    hdrs = ["TIMESTAMP","EVENT TYPE","ENTITY","ENTITY ID","OLD VALUE","NEW VALUE","OPERATOR"]
    for ci, h in enumerate(hdrs, 2):
        c = ws.cell(R, ci); c.value = h
        c.font = _font(CLR_MUTED, bold=True, size=9)
        c.fill = _fill(CLR_SURF2); c.border = _bdr()
        c.alignment = Alignment(horizontal="center", vertical="center")

    if not _DB_AVAILABLE:
        return

    events = _db.get_system_events(limit=500)
    clr_map = {"credit_action": CLR_AMBER, "risk_change": CLR_RED,
               "price_override": CLR_BLUE, "udhaar_blocked": CLR_RED}
    for ev in events:
        R += 1
        ws.row_dimensions[R].height = 13
        ec = clr_map.get(ev.get("event_type",""), CLR_TEXT)
        row = [
            (str(ev.get("timestamp",""))[:19],  CLR_MUTED),
            (ev.get("event_type",""),            ec),
            (ev.get("entity_type",""),           CLR_MUTED),
            (ev.get("entity_id",""),             CLR_TEXT),
            (ev.get("old_value") or "—",         CLR_MUTED),
            (ev.get("new_value") or "—",         CLR_GREEN),
            (ev.get("operator","system"),        CLR_MUTED),
        ]
        for ci, (val, clr) in enumerate(row, 2):
            c = ws.cell(R, ci); c.value = val
            c.font = _font(clr, size=9)
            c.fill = _fill("1A2233" if R % 2 == 0 else CLR_SURF2)
            c.border = _bdr()
            c.alignment = Alignment(horizontal="left", vertical="center")

    ws.freeze_panes = "B5"
    log.info("_write_system_events_sheet: wrote %d events", len(events))


def _write_analytics_sheets():
    """
    Single unified write pipeline. All sheets in one wb.open() → save() cycle.
    Sheet order: DASHBOARD first, then analytics sheets.
    """
    try:
        with _write_lock:
            wb = _load_wb(data_only=False)
            _write_dashboard(wb)               # ← new analytical dashboard
            _write_service_analytics(wb)
            _write_demand_analytics(wb)
            _write_bundle_analytics(wb)
            _write_risk_dashboard(wb)
            _write_ai_pricing_sheet(wb)
            _write_ai_control_log(wb)
            _write_price_memory_sheet(wb)
            _write_price_history_from_db(wb)
            _write_cashflow_sheet(wb)          # ← v6.0 cashflow intelligence
            _write_system_events_sheet(wb)     # ← v6.0 audit trail
            wb.save(EXCEL_PATH)
            wb.close()
        log.info("Analytics sheets written (11 sheets, single pipeline)")
    except PermissionError as e:
        # File is open in desktop Excel — analytics are secondary, not critical
        _health_soft("excel_write_errors")
        log.warning("_write_analytics_sheets: file locked (Excel open?): %s", e)
        try: wb.close()
        except Exception: pass
    except Exception as e:
        # All other analytics failures — soft (DB is authoritative, Excel is export only)
        _health_soft("excel_write_errors")
        log.error("_write_analytics_sheets FAILED: %s\n%s", e, traceback.format_exc())
        try: wb.close()
        except Exception: pass


# ═══════════════════════════════════════════════════════════════════════════════
# SYSTEM INTEGRITY CHECK — NEW ENDPOINT (Requirement 9)
# ═══════════════════════════════════════════════════════════════════════════════

def _run_integrity_check() -> dict:
    """
    Internal validation sweep. Returns structured report.
    Checks:
      - No negative-margin services (cost > price)
      - No zero-price non-Other services
      - Cache freshness (age vs TTL)
      - Price memory consistency (no inf/nan in cooldown timestamps)
      - Excel file accessible and readable
    """
    issues     = []
    warnings   = []
    info       = []
    now        = time.time()

    # ── Service data checks ───────────────────────────────────────────────────
    try:
        svc_map = _get_services()
        neg_margin = [n for n, s in svc_map.items()
                      if s["price"] > 0 and s["cost"] > s["price"]]
        zero_price = [n for n, s in svc_map.items()
                      if s["price"] == 0 and n != "Other"]
        zero_cost  = [n for n, s in svc_map.items()
                      if s["price"] > 0 and s["cost"] == 0 and n != "Other"]

        if neg_margin:
            issues.append({
                "check":   "negative_margin",
                "severity":"critical",
                "count":   len(neg_margin),
                "services":neg_margin[:10],
            })
        if zero_price:
            issues.append({
                "check":   "zero_price",
                "severity":"critical",
                "count":   len(zero_price),
                "services":zero_price[:10],
            })
        if zero_cost:
            warnings.append({
                "check":   "zero_cost",
                "severity":"warning",
                "count":   len(zero_cost),
                "services":zero_cost[:10],
            })
        info.append({"check": "service_load", "status": "ok", "count": len(svc_map)})
    except Exception as e:
        issues.append({"check": "service_load", "severity": "critical", "error": str(e)})

    # ── Cache freshness ───────────────────────────────────────────────────────
    with _demand_lock:
        demand_age = round(now - _demand_cache_built, 1)
    with _bundle_lock:
        bundle_age = round(now - _bundle_stats_built, 1)
    with _ai_lock:
        ai_age = round(now - _ai_cache_built, 1)

    cache_ttl = _AI_TTL * 3   # stale if 3× TTL old
    for name, age in [("demand", demand_age), ("bundle", bundle_age), ("ai", ai_age)]:
        if age > cache_ttl:
            warnings.append({
                "check":    f"cache_stale_{name}",
                "severity": "warning",
                "age_s":    age,
                "ttl_s":    cache_ttl,
            })
        else:
            info.append({"check": f"cache_{name}", "status": "ok", "age_s": age})

    # ── Price memory consistency ──────────────────────────────────────────────
    corrupted_pm = []
    with _price_mem_lock:
        for svc, pm in _price_memory.items():
            cd = pm.get("cooldown_until", 0)
            lp = pm.get("last_price", 0)
            if not isinstance(cd, (int, float)) or math.isnan(cd) or math.isinf(cd):
                corrupted_pm.append(svc)
            elif not isinstance(lp, (int, float)) or math.isnan(lp) or lp < 0:
                corrupted_pm.append(svc)
    if corrupted_pm:
        issues.append({
            "check":    "price_memory_corruption",
            "severity": "critical",
            "services": corrupted_pm,
        })
    else:
        info.append({"check": "price_memory", "status": "ok"})

    # ── Excel accessibility ───────────────────────────────────────────────────
    try:
        if not os.path.exists(EXCEL_PATH):
            issues.append({"check": "excel_file", "severity": "critical",
                            "error": "File not found"})
        else:
            fsize = os.path.getsize(EXCEL_PATH)
            if fsize < 1000:
                warnings.append({"check": "excel_file_size", "severity": "warning",
                                  "size_bytes": fsize})
            else:
                info.append({"check": "excel_file", "status": "ok",
                             "size_bytes": fsize})
    except Exception as e:
        issues.append({"check": "excel_file", "severity": "critical", "error": str(e)})

    # ── Fail-safe state ───────────────────────────────────────────────────────
    with _health_lock:
        fs_mode = _health.get("failsafe_mode", False)
        crit_streak = _health.get("critical_error_streak", 0)

    if fs_mode:
        issues.append({
            "check":    "failsafe_active",
            "severity": "critical",
            "message":  "System in fail-safe mode — pricing locked to static",
        })

    overall = ("critical" if issues else ("warning" if warnings else "ok"))

    return {
        "status":          overall,
        "failsafe_active": fs_mode,
        "critical_streak": crit_streak,
        "issues":          issues,
        "warnings":        warnings,
        "info":            info,
        "checked_at":      datetime.now().isoformat(),
    }


# ═══════════════════════════════════════════════════════════════════════════════
# FUZZY SEARCH
# ═══════════════════════════════════════════════════════════════════════════════

def _fuzzy_score(text: str, query: str) -> int:
    t = text.lower()
    q = query.lower().strip()
    if not q:
        return 1
    if t == q:
        return 1000
    if t.startswith(q):
        return 900
    if q in t:
        return 800
    qw = q.split()
    if len(qw) > 1 and all(w in t for w in qw):
        return 700
    qi = score = 0
    prev = -1
    for i, ch in enumerate(t):
        if qi < len(q) and ch == q[qi]:
            score += 10 if prev == i - 1 else 1
            prev = i; qi += 1
    if qi == len(q):
        return max(1, score)
    if qw and qw[0] in t:
        return 300
    return 0


# ═══════════════════════════════════════════════════════════════════════════════
# STARTUP
# ═══════════════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 5 — CUSTOMER INTELLIGENCE
# ═══════════════════════════════════════════════════════════════════════════════

def get_customer_profile(phone: str) -> dict | None:
    """
    Return {total_spend, visit_count, avg_ticket, last_seen} for a customer.
    Delegates to DB layer. Returns None if phone not found or DB unavailable.
    """
    if not _DB_AVAILABLE:
        return None
    return _db.get_customer_profile(phone)


# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 6 — BACKGROUND DB SYNC  (DB → Excel prices kept in sync every 300 s)
# ═══════════════════════════════════════════════════════════════════════════════

_DB_SYNC_INTERVAL = 300   # seconds between DB→service catalogue sync


def _db_sync_loop():
    """
    DB health monitor (replaces the old DB→DB sync no-op loop).
    Previously this read from DB via _get_services() then wrote back via
    sync_services_from_dict() — a read-modify-write on the same source: no-op.
    Now: periodic health check that verifies service count + transaction count
    and alerts if data appears to have been wiped (integrity guard).
    """
    log.info("DB health monitor started (interval=%ds)", _DB_SYNC_INTERVAL)
    while True:
        time.sleep(_DB_SYNC_INTERVAL)
        if not _DB_AVAILABLE:
            continue
        try:
            conn      = _db.get_db()
            svc_count = conn.execute("SELECT COUNT(*) FROM services").fetchone()[0]
            tx_count  = conn.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]
            log.info("DB health: services=%d transactions=%d", svc_count, tx_count)
            if svc_count == 0:
                log.error("DB INTEGRITY ALERT: services table empty — possible data loss!")
                _health_soft("cache_errors")
        except Exception as e:
            log.warning("DB health monitor error (non-fatal): %s", e)


def _startup_validation():
    log.info("=== CityCyber POS v4 — Production-Grade Self-Learning Profit Engine ===")

    # FIX BUG-9: load price memory from disk before any pricing decisions
    global _price_memory
    with _price_mem_lock:
        loaded = _load_price_memory()
        _price_memory.update(loaded)
    log.info("Price memory loaded: %d services", len(_price_memory))

    try:
        svc_map = _get_services()
        for name, svc in svc_map.items():
            BASE_MAX_PRICES[name] = round(svc["price"] * 1.5, 1)

        zero_price = [n for n, s in svc_map.items() if s["price"] == 0 and n != "Other"]
        neg_margin = [n for n, s in svc_map.items()
                      if s["price"] > 0 and s["cost"] > s["price"]]
        by_role    = collections.Counter(s["role"] for s in svc_map.values())

        log.info("Loaded %d services | roles: %s", len(svc_map), dict(by_role))
        if zero_price:
            log.warning("Services needing price: %s", zero_price)
        if neg_margin:
            log.error("CRITICAL — cost > price: %s", neg_margin)
    except Exception as e:
        log.error("Startup validation failed: %s\n%s", e, traceback.format_exc())

    _ensure_daily_loaded()

    try:
        _ensure_demand_cache()
        _ensure_bundle_stats()
        _ensure_ai_cache()
        log.info("AI + Demand + Bundle caches warmed up")
    except Exception as e:
        log.warning("Cache pre-warm partial failure (non-fatal): %s", e)

    try:
        _write_analytics_sheets()
    except Exception as e:
        log.warning("Initial analytics write skipped (non-fatal): %s", e)

    t = threading.Thread(target=_background_loop, daemon=True)
    t.start()
    log.info("Background self-learning loop started (interval=%ds)", _AI_TTL)

    # Phase 1 + 6: initialise DB and start background sync thread
    if _DB_AVAILABLE:
        try:
            _db.init_db()
            log.info("DB layer initialised: %s", _db.DB_PATH)
            # Seed DB with current service catalogue on first run
            try:
                _db.sync_services_from_dict(_get_services())
            except Exception as _se:
                log.warning("Initial DB service sync failed (non-fatal): %s", _se)
        except Exception as e:
            log.error("DB init failed (system continues without DB): %s", e)

        db_sync_thread = threading.Thread(target=_db_sync_loop, daemon=True)
        db_sync_thread.start()
        log.info("DB health monitor thread started (interval=300s)")

        # DB-FIRST v5: restore price memory from DB price_events table
        # (supplements price_memory.json — DB is authoritative)
        try:
            conn = _db.get_db()
            svc_names = [r[0] for r in conn.execute(
                "SELECT DISTINCT service_name FROM price_events"
            ).fetchall()]
            restored = 0
            with _price_mem_lock:
                for svc in svc_names:
                    if svc not in _price_memory:
                        pm = _db.get_price_memory_from_db(svc)
                        if pm:
                            _price_memory[svc] = pm
                            restored += 1
            if restored:
                log.info("Price memory restored from DB: %d services", restored)
        except Exception as _pme:
            log.warning("Price memory DB restore failed (non-fatal): %s", _pme)


# ═══════════════════════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/get-services", methods=["GET"])
def get_services_route():
    try:
        svc_map    = _get_services()
        services   = list(svc_map.values())
        categories = sorted(set(s["category"] for s in services if s["category"]))
        priced     = [s for s in services if s["price"] > 0]
        unpriced   = [s for s in services if s["price"] == 0 and s["name"] != "Other"]
        by_role    = collections.Counter(s["role"] for s in services)
        return jsonify({
            "status":     "ok",
            "services":   services,
            "categories": categories,
            "stats": {
                "total":         len(services),
                "priced":        len(priced),
                "needs_pricing": len(unpriced),
                "by_role":       dict(by_role),
            },
        })
    except FileNotFoundError as e:
        return jsonify({"status": "error", "message": str(e)}), 500
    except Exception as e:
        log.error("get-services: %s", e)
        return jsonify({"status": "error", "message": "Internal server error"}), 500


@app.route("/search", methods=["GET"])
def search_route():    # FIX BUG-12: removed duplicate @app.route decorator
    try:
        q           = request.args.get("q", "").strip()
        cat_filter  = request.args.get("cat", "All").strip()
        limit       = min(int(request.args.get("limit", 50)), 200)
        priced_only = request.args.get("priced_only", "0") == "1"

        svc_map    = _get_services()
        candidates = list(svc_map.values())
        if cat_filter and cat_filter != "All":
            candidates = [s for s in candidates if s["category"] == cat_filter]
        if priced_only:
            candidates = [s for s in candidates if s["price"] > 0]
        if not q:
            candidates.sort(key=lambda s: (s["category"], s["name"]))
            return jsonify({"status": "ok", "results": candidates[:limit],
                            "query": q, "total": len(candidates)})
        scored = sorted(
            [(sc, s) for s in candidates for sc in [_fuzzy_score(s["name"], q)] if sc > 0],
            key=lambda x: -x[0],
        )
        results = [s for _, s in scored[:limit]]
        return jsonify({"status": "ok", "results": results,
                        "query": q, "total": len(scored)})
    except Exception as e:
        log.error("search: %s", e)
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/ai-insights", methods=["GET"])
def ai_insights():
    try:
        ai_map  = _ensure_ai_cache()
        svc_arg = request.args.get("service", "").strip()

        if svc_arg:
            insight = ai_map.get(svc_arg)
            if not insight:
                for k, v in ai_map.items():
                    if svc_arg.lower() in k.lower():
                        insight = v
                        break
            if not insight:
                return jsonify({"status": "ok", "pricing": None,
                                "bundles": [], "alerts": []}), 200
            return jsonify({
                "status": "ok",
                "pricing": {
                    "service":         insight["service"],
                    "current_price":   insight["current_price"],
                    "suggested_price": insight["suggested_price"],
                    "confidence":      insight["confidence"],
                    "delta_pct":       insight["delta_pct"],
                    "reason":          insight["reason"],
                    "profit_estimate": insight.get("profit_estimate", 0),
                    "role":            insight.get("role"),
                    "price_memory":    insight.get("price_memory", {}),
                },
                "bundles":        insight["bundles"],
                "alerts":         insight["alerts"],
                "demand":         insight["demand"],
                "ai_accept_rate": insight.get("ai_accept_rate"),
            })

        all_alerts, top_bundles, pricing_list = [], [], []
        seen_bundle_labels: set = set()

        for name, ins in ai_map.items():
            for alert in ins["alerts"]:
                all_alerts.append({**alert, "service": name})
            for b in ins["bundles"]:
                if b["label"] not in seen_bundle_labels:
                    top_bundles.append(b)
                    seen_bundle_labels.add(b["label"])
            if ins["confidence"] >= 0.40 and ins["delta_pct"] != 0:
                pricing_list.append({
                    "service":         name,
                    "current_price":   ins["current_price"],
                    "suggested_price": ins["suggested_price"],
                    "confidence":      ins["confidence"],
                    "delta_pct":       ins["delta_pct"],
                    "reason":          ins["reason"],
                    "profit_estimate": ins.get("profit_estimate", 0),
                    "role":            ins.get("role"),
                })

        pricing_list.sort(key=lambda x: -x["confidence"])
        top_bundles.sort(key=lambda x: -x["score"])
        all_alerts.sort(key=lambda x: {"critical": 0, "warning": 1, "info": 2}
                        .get(x.get("severity", "info"), 3))

        return jsonify({
            "status":  "ok",
            "pricing": pricing_list[:10],
            "bundles": top_bundles[:10],
            "alerts":  all_alerts[:20],
        })
    except Exception as e:
        log.error("ai-insights: %s\n%s", e, traceback.format_exc())
        return jsonify({"status": "error", "message": "AI engine unavailable"}), 500


@app.route("/demand", methods=["GET"])
def demand_route():
    try:
        demand  = _ensure_demand_cache()
        svc_arg = request.args.get("service", "").strip()
        if svc_arg:
            return jsonify({"status": "ok", "service": svc_arg,
                            "demand": demand.get(svc_arg, {})})
        return jsonify({"status": "ok", "demand": demand,
                        "total_services_tracked": len(demand)})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/bundles", methods=["GET"])
def bundles_route():
    try:
        stats   = _ensure_bundle_stats()
        svc_arg = request.args.get("service", "").strip()
        if svc_arg:
            svc_map = _get_services()
            sugg    = _get_bundle_suggestions_learned(svc_arg, svc_map)
            return jsonify({"status": "ok", "service": svc_arg, "bundles": sugg})
        ranked = sorted(
            [{"service_a": a, "service_b": b, **d} for (a, b), d in stats.items()],
            key=lambda x: -x["count"]
        )
        return jsonify({"status": "ok", "bundles": ranked[:50],
                        "total_pairs": len(stats)})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/cart-check", methods=["POST"])
def cart_check():
    try:
        data  = request.get_json(silent=True) or {}
        items = data.get("items", [])
        if not isinstance(items, list):
            return jsonify({"status": "error", "message": "items must be a list"}), 400
        svc_map = _get_services()
        result  = _check_cart_bundles(items, svc_map)
        return jsonify({"status": "ok", **result})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/operator-decision", methods=["POST"])
def operator_decision():
    try:
        data    = request.get_json(silent=True) or {}
        service = str(data.get("service", "")).strip()
        action  = str(data.get("action", "")).strip()
        if not service or action not in ("accepted", "ignored"):
            return jsonify({"status": "error",
                            "message": "service and action (accepted|ignored) required"}), 400
        _log_operator_decision(service, action)
        rate = _get_ai_accept_rate(service)
        return jsonify({"status": "ok", "service": service,
                        "action": action, "accept_rate": rate})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/add-transaction", methods=["POST"])
def add_transaction():
    """
    DB-FIRST atomic transaction endpoint.
    Idempotency key (X-Idempotency-Key header) prevents double-billing on retry.
    Validates ALL items, writes atomically to DB, then async-writes to Excel.
    """
    # ── Idempotency: reject duplicate submissions ─────────────────────────────
    idem_key = (request.headers.get("X-Idempotency-Key") or "").strip() or None
    if idem_key and _DB_AVAILABLE:
        cached = _db.check_idempotency(idem_key)
        if cached:
            log.info("Idempotent replay served: key=%s", idem_key)
            import json as _json
            try:
                return jsonify(_json.loads(cached)), 200
            except Exception:
                pass  # corrupted cache entry — fall through and reprocess

    data         = request.get_json(silent=True) or {}
    payment_mode = str(data.get("payment", "Cash")).strip()
    notes        = str(data.get("notes", "")).strip()

    if payment_mode not in ("Cash", "UPI", "Card", "Udhaar"):
        return jsonify({"status": "error",
                        "message": "Payment must be Cash, UPI, Card, or Udhaar."}), 400

    # Udhaar requires a customer phone to track the debt
    if payment_mode == "Udhaar":
        _udhaar_phone = str(data.get("customer_phone", "")).strip()
        if not _udhaar_phone or len(_udhaar_phone) < 10:
            return jsonify({"status": "error",
                            "message": "Udhaar requires a valid customer phone number (min 10 digits)."}), 400
        # Credit risk gate: block high-risk customers from new udhaar
        if _DB_AVAILABLE:
            try:
                _risk_level = _db.update_customer_risk(_udhaar_phone)
                if _risk_level == "high":
                    return jsonify({
                        "status":    "error",
                        "message":   "Udhaar blocked: customer risk level is HIGH. "
                                     "Outstanding balance must be cleared before new credit.",
                        "risk_level": _risk_level,
                        "phone":      _udhaar_phone,
                    }), 400
            except Exception as _re:
                log.warning("Risk check failed (non-critical, allowing): %s", _re)

    raw_items = data.get("items")
    if raw_items is None:
        service_name = str(data.get("service", "")).strip()
        if not service_name:
            return jsonify({"status": "error",
                            "message": "Service name is required."}), 400
        try:
            qty = int(data.get("quantity", 0))
        except (TypeError, ValueError):
            return jsonify({"status": "error",
                            "message": "Quantity must be an integer."}), 400
        raw_items = [{"service": service_name, "quantity": qty}]

    if not isinstance(raw_items, list) or len(raw_items) == 0:
        return jsonify({"status": "error",
                        "message": "items must be a non-empty list."}), 400

    try:
        svc_map = _get_services()
    except FileNotFoundError as e:
        return jsonify({"status": "error", "message": str(e)}), 500

    # ── Phase 1: Validate + merge ALL items before any write ──────────────────
    merged: dict[str, int] = {}
    for idx, item in enumerate(raw_items):
        name = str(item.get("service", "")).strip()
        if not name:
            return jsonify({"status": "error",
                            "message": f"Item #{idx+1}: service name required."}), 400
        try:
            qty = int(item.get("quantity", 0))
        except (TypeError, ValueError):
            return jsonify({"status": "error",
                            "message": f"Item #{idx+1} '{name}': quantity must be integer."}), 400
        if qty <= 0:
            return jsonify({"status": "error",
                            "message": f"Item #{idx+1} '{name}': quantity must be > 0."}), 400
        if qty > 10_000:
            return jsonify({"status": "error",
                            "message": f"'{name}': quantity {qty} exceeds 10,000."}), 400
        merged[name] = merged.get(name, 0) + qty

    resolved = []
    for name, qty in merged.items():
        if name not in svc_map:
            return jsonify({"status": "error",
                            "message": f"Service not found: '{name}'"}), 400
        svc        = svc_map[name]
        base_price = svc["price"]
        cost       = svc["cost"]
        if base_price <= 0 and name != "Other":
            return jsonify({
                "status":     "error",
                "message":    f"'{name}' has no sell price. Set it in MASTER first.",
                "needs_price": True,
            }), 400

        # ── Phase 2: resolve override for this item (per-item, from items[] array) ──
        # Item-level override_price takes priority; falls back to top-level for
        # single-item legacy calls.
        override_price = item.get("override_price", data.get("override_price"))
        discount_pct   = item.get("discount_pct",   data.get("discount_pct"))

        override_type  = "none"
        override_value = None

        if override_price is not None:
            try:
                op = float(override_price)
            except (TypeError, ValueError):
                op = 0.0
            if op > 0:
                final_price    = op
                override_type  = "manual_price"
                override_value = final_price
        if override_type == "none" and discount_pct is not None:
            try:
                dp = float(discount_pct)
            except (TypeError, ValueError):
                dp = 0.0
            if dp > 0:
                final_price    = round(base_price * (1 - dp / 100), 2)
                override_type  = "discount_pct"
                override_value = dp
        if override_type == "none":
            final_price = base_price

        # ── Phase 6: safety floor — never sell below cost ──────────────────
        if final_price < cost:
            final_price = round(cost * 1.05, 2)
            log.warning(
                "Override safety floor applied for '%s': "
                "final_price raised to %.2f (cost=%.2f)",
                name, final_price, cost
            )

        if override_type != "none":
            log.warning(
                "Override used: %s base=%.2f final=%.2f type=%s value=%s",
                name, base_price, final_price, override_type, override_value
            )

        revenue    = round(qty * final_price, 2)
        profit     = round(qty * (final_price - cost), 2)
        cost_total = round(qty * cost, 2)

        resolved.append({
            "name":          name,
            "qty":           qty,
            "base_price":    base_price,
            "final_price":   final_price,
            "price":         final_price,   # kept for response compat
            "cost":          cost,
            "revenue":       revenue,
            "profit":        profit,
            "cost_total":    cost_total,
            "override_type": override_type,
            "override_value":override_value,
        })

    # ── DB-FIRST: authoritative atomic write — all items or none ────────────
    timestamp      = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    customer_phone = str(data.get("customer_phone", "")).strip() or None

    # Extract customer name from notes field (frontend sends "Name: Ramesh · note text")
    customer_name = None
    _notes_clean  = notes
    if notes and notes.startswith("Name:"):
        _parts = notes.split("·", 1)
        customer_name  = _parts[0].replace("Name:", "").strip() or None
        _notes_clean   = _parts[1].strip() if len(_parts) > 1 else ""

    if not _DB_AVAILABLE:
        return jsonify({"status": "error",
                        "message": "Database unavailable. Transaction not saved."}), 503

    try:
        _db.insert_transactions_atomic(
            items=resolved,
            timestamp=timestamp,
            payment_mode=payment_mode,
            customer_phone=customer_phone,
        )
        log.info("DB atomic write OK: %d item(s) ts=%s", len(resolved), timestamp)
    except Exception as _dbe:
        log.error("DB atomic write FAILED — transaction aborted: %s", _dbe)
        return jsonify({"status": "error",
                        "message": "Database write failed. Transaction not saved."}), 500

    # ── Upsert customer record with name ─────────────────────────────────────
    if customer_phone and _DB_AVAILABLE:
        try:
            _db.upsert_customer(
                customer_phone,
                sum(i["revenue"] for i in resolved),
                name=customer_name,
            )
        except Exception as _ce:
            log.warning("upsert_customer failed (non-critical): %s", _ce)

    # ── UDHAAR: write debit entries to ledger after successful transaction ────
    if payment_mode == "Udhaar" and customer_phone:
        try:
            # Build a compact sale note
            if len(resolved) == 1:
                _u_note_base = f"Sale: {resolved[0]['name']} ×{resolved[0]['qty']}"
            else:
                _u_note_base = f"Sale: {len(resolved)} items"
            if customer_name:
                _u_note_base = f"{customer_name} — {_u_note_base}"
            if _notes_clean:
                _u_note_base = f"{_u_note_base} ({_notes_clean})"

            for _u_item in resolved:
                _item_note = f"Sale: {_u_item['name']} ×{_u_item['qty']}"
                if customer_name:
                    _item_note = f"{customer_name} — {_item_note}"
                _db.add_udhaar_entry(
                    phone=customer_phone,
                    amount=_u_item["revenue"],
                    entry_type="debit",
                    note=_item_note,
                )
            log.info("Udhaar ledger entries written: phone=%s items=%d", customer_phone, len(resolved))
        except Exception as _ue:
            # Non-fatal — transaction is already saved, just log the ledger failure
            log.warning("Udhaar ledger write failed (non-critical, transaction saved): %s", _ue)

    total_revenue = sum(i["revenue"] for i in resolved)
    total_profit  = sum(i["profit"]  for i in resolved)

    # ── SECONDARY: Excel export (best-effort async — never blocks or fails tx) ──
    # All args passed explicitly — no closure capture of mutable outer scope.
    def _excel_export_fn(_items, _ts, _pay, _phone, _notes):
        # ── TXLOG cell styles — matches existing sheet design exactly ────────
        _FONT_BASE  = Font(name="Segoe UI", size=9, bold=False)
        _FONT_BOLD  = Font(name="Segoe UI", size=9, bold=True)
        _CLR_WHITE  = "F0F6FC"
        _CLR_DIM    = "8B949E"
        _CLR_GREEN  = "39D353"
        _FMT_RUPEE  = '₹#,##0.00'
        _FMT_INT    = '#,##0'
        _ALIGN_C    = Alignment(horizontal="center", vertical="center")
        _ALIGN_L    = Alignment(horizontal="left",   vertical="center")
        _BORDER_T   = Border(
            left=Side(style="thin", color="30363D"),
            right=Side(style="thin", color="30363D"),
            top=Side(style="thin", color="30363D"),
            bottom=Side(style="thin", color="30363D"),
        )
        # Per-column style spec: (font, color_hex, number_format, alignment)
        _COL_STYLES = {
            2:  (_FONT_BASE, _CLR_DIM,   None,       _ALIGN_L),   # Timestamp
            3:  (_FONT_BASE, _CLR_WHITE,  None,       _ALIGN_L),   # Service
            4:  (_FONT_BASE, _CLR_WHITE,  _FMT_INT,   _ALIGN_C),   # Qty
            5:  (_FONT_BOLD, _CLR_WHITE,  _FMT_RUPEE, _ALIGN_C),   # Revenue
            6:  (_FONT_BOLD, _CLR_GREEN,  _FMT_RUPEE, _ALIGN_C),   # Profit
            7:  (_FONT_BASE, _CLR_DIM,    _FMT_RUPEE, _ALIGN_C),   # Cost
            8:  (_FONT_BOLD, _CLR_WHITE,  None,       _ALIGN_C),   # Payment
            9:  (_FONT_BASE, _CLR_DIM,    None,       _ALIGN_C),   # Customer
            10: (_FONT_BASE, _CLR_DIM,    _FMT_RUPEE, _ALIGN_C),   # Base Price
            11: (_FONT_BASE, _CLR_DIM,    _FMT_RUPEE, _ALIGN_C),   # Final Price
            12: (_FONT_BASE, _CLR_DIM,    None,       _ALIGN_C),   # Override Type
            13: (_FONT_BASE, _CLR_DIM,    None,       _ALIGN_C),   # Override Value
        }

        def _style_cell(cell, col):
            """Apply consistent TXLOG styling to a cell."""
            style = _COL_STYLES.get(col)
            if not style:
                return
            font_base, color_hex, num_fmt, align = style
            cell.font      = Font(name=font_base.name, size=font_base.size,
                                  bold=font_base.bold, color=color_hex)
            cell.alignment = align
            cell.border    = _BORDER_T
            if num_fmt:
                cell.number_format = num_fmt

        try:
            with _write_lock:
                wb2 = _load_wb(data_only=False)
                ws  = wb2[TXLOG_SHEET]
                r   = _find_next_txlog_row(ws)
                for item in _items:
                    ws.cell(r, 2).value  = _ts
                    ws.cell(r, 3).value  = item["name"]
                    ws.cell(r, 4).value  = item["qty"]
                    ws.cell(r, 5).value  = item["revenue"]
                    ws.cell(r, 6).value  = item["profit"]
                    ws.cell(r, 7).value  = item["cost_total"]
                    ws.cell(r, 8).value  = _pay
                    ws.cell(r, 9).value  = _phone or (_notes if _notes else "—")
                    ws.cell(r, 10).value = item["base_price"]
                    ws.cell(r, 11).value = item["final_price"]
                    ws.cell(r, 12).value = item["override_type"] or "—"
                    ws.cell(r, 13).value = item["override_value"] if item["override_value"] is not None else "—"
                    # Apply consistent styling to every cell in this row
                    for col in range(2, 14):
                        _style_cell(ws.cell(r, col), col)
                    r += 1
                wb2.save(EXCEL_PATH)
                wb2.close()
        except PermissionError as _pe:
            _health_soft("excel_write_errors")
            log.warning("Excel export locked (data safe in DB): %s", _pe)
        except Exception as _xe:
            _health_soft("excel_write_errors")
            log.warning("Excel export failed (data safe in DB): %s", _xe)
            try: wb2.close()
            except Exception: pass

    threading.Thread(
        target=_excel_export_fn,
        args=(list(resolved), timestamp, payment_mode, customer_phone, notes),
        daemon=True,
    ).start()

    # ── Post-write bookkeeping ────────────────────────────────────────────────
    _accum_today(round(total_revenue, 2), round(total_profit, 2))

    global _demand_cache_built, _ai_cache_built, _bundle_stats_built
    with _demand_lock:
        _demand_cache_built = 0.0
    with _bundle_lock:
        _bundle_stats_built = 0.0
    with _ai_lock:
        _ai_cache_built = 0.0

    for item in resolved:
        log.info("tx OK: %s x%d %s rev=%.2f profit=%.2f",
                 item["name"], item["qty"], payment_mode,
                 item["revenue"], item["profit"])

    # ── Build and record response for idempotency ────────────────────────────
    import json as _ijson
    if len(resolved) == 1 and data.get("items") is None:
        item = resolved[0]
        resp_body = {
            "status":  "ok",
            "message": "Transaction saved.",
            "details": {
                "timestamp": timestamp, "service": item["name"],
                "quantity":  item["qty"],   "price":   item["price"],
                "revenue":   item["revenue"], "profit": item["profit"],
                "payment":   payment_mode,
            },
        }
    else:
        resp_body = {
            "status":  "ok",
            "message": f"{len(resolved)} item(s) saved.",
            "summary": {
                "total_items":   len(resolved),
                "total_revenue": round(total_revenue, 2),
                "total_profit":  round(total_profit, 2),
                "timestamp":     timestamp,
                "payment":       payment_mode,
            },
            "details": [
                {"service": i["name"], "quantity": i["qty"],
                 "price": i["price"], "revenue": i["revenue"], "profit": i["profit"]}
                for i in resolved
            ],
        }

    if idem_key and _DB_AVAILABLE:
        try:
            _db.record_idempotency(idem_key, _ijson.dumps(resp_body))
        except Exception as _ie:
            log.warning("Idempotency record failed (non-critical): %s", _ie)

    return jsonify(resp_body)


@app.route("/summary", methods=["GET"])
def summary():
    try:
        _ensure_daily_loaded()
        today = _get_today_str()
        with _daily_lock:
            d = _daily.get(today, {"revenue": 0.0, "profit": 0.0, "count": 0})
        return jsonify({
            "status":        "ok",
            "today_revenue": round(d["revenue"], 2),
            "today_profit":  round(d["profit"],  2),
            "today_count":   d["count"],
            "date":          date.today().strftime("%d %b %Y"),
        })
    except Exception as e:
        log.error("summary: %s", e)
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/reload-services", methods=["POST"])
def reload_services():
    try:
        svc_map = _get_services(force_reload=True)
        for name, svc in svc_map.items():
            BASE_MAX_PRICES[name] = round(svc["price"] * 1.5, 1)
        global _ai_cache_built
        with _ai_lock:
            _ai_cache_built = 0.0
        return jsonify({"status": "ok", "count": len(svc_map), "message": "Cache reloaded."})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/update-price", methods=["POST"])
def update_price_route():
    data = request.get_json(silent=True) or {}
    name = str(data.get("service", "")).strip()
    if not name:
        return jsonify({"status": "error", "message": "service is required"}), 400
    try:
        price = float(data["price"])
        cost  = float(data.get("cost", 0))
    except (KeyError, TypeError, ValueError):
        return jsonify({"status": "error", "message": "price must be a number"}), 400
    if price < 0 or cost < 0:
        return jsonify({"status": "error",
                        "message": "price and cost must be >= 0"}), 400
    result = update_price(name, price, cost, source="manual")
    return jsonify(result), (200 if result["status"] == "ok" else 400)


@app.route("/write-analytics", methods=["POST"])
def write_analytics_route():
    try:
        _write_analytics_sheets()
        return jsonify({"status": "ok", "message": "Analytics sheets updated."})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/auto-pricing", methods=["POST"])
def toggle_auto_pricing():
    global AUTO_PRICING
    data = request.get_json(silent=True) or {}
    AUTO_PRICING = bool(data.get("enabled", False))
    log.info("AUTO_PRICING set to %s", AUTO_PRICING)
    return jsonify({"status": "ok", "AUTO_PRICING": AUTO_PRICING})


@app.route("/system-health", methods=["GET"])
def system_health():
    try:
        now = time.time()
        with _health_lock:
            h = dict(_health)

        with _ai_lock:
            ai_cache_age  = round(now - _ai_cache_built, 1)
            ai_cache_size = len(_ai_cache)
        with _demand_lock:
            dem_cache_age = round(now - _demand_cache_built, 1)
        with _bundle_lock:
            bun_cache_age = round(now - _bundle_stats_built, 1)

        uptime_secs = round(now - h.get("uptime_start", now), 0)
        last_ok     = h.get("loop_last_ok", 0)
        last_ok_str = (datetime.fromtimestamp(last_ok).isoformat()
                       if last_ok > 0 else "never")

        status = "degraded" if h.get("failsafe_mode") else (
            "warning"  if h.get("loop_errors", 0) > 0 else "ok"
        )

        return jsonify({
            "status":           status,
            "failsafe_mode":    h.get("failsafe_mode", False),
            "failsafe_at":      h.get("failsafe_triggered"),
            "critical_streak":  h.get("critical_error_streak", 0),
            "uptime_seconds":   uptime_secs,
            "loop": {
                "errors":       h.get("loop_errors", 0),
                "last_ok":      last_ok_str,
                "last_refresh": datetime.fromtimestamp(h.get("last_refresh", 0)).isoformat()
                                 if h.get("last_refresh", 0) > 0 else "never",
            },
            "cache": {
                "ai":     {"age_s": ai_cache_age,  "size": ai_cache_size},
                "demand": {"age_s": dem_cache_age},
                "bundle": {"age_s": bun_cache_age},
            },
            "errors": {
                "loop":          h.get("loop_errors", 0),
                "excel_write":   h.get("excel_write_errors", 0),
                "excel_read":    h.get("excel_read_errors", 0),
                "cache":         h.get("cache_errors", 0),
            },
            "anomalies":   h.get("anomalies", [])[-10:],
            "auto_pricing":AUTO_PRICING,
        })
    except Exception as e:
        log.error("system-health: %s", e)
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/system-integrity", methods=["GET"])
def system_integrity():
    """NEW — Requirement 9: full internal consistency check."""
    try:
        report = _run_integrity_check()
        code   = 200 if report["status"] != "critical" else 500
        return jsonify(report), code
    except Exception as e:
        log.error("system-integrity: %s", e)
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/price-memory", methods=["GET"])
def price_memory_route():
    try:
        svc_arg = request.args.get("service", "").strip()
        now     = time.time()
        with _price_mem_lock:
            if svc_arg:
                pm = _price_memory.get(svc_arg)
                if not pm:
                    return jsonify({"status": "ok", "service": svc_arg, "memory": None})
                result = dict(pm)
                result["in_cooldown"] = now < pm.get("cooldown_until", 0.0)
                return jsonify({"status": "ok", "service": svc_arg, "memory": result})
            all_pm = {}
            for name, pm in _price_memory.items():
                r = dict(pm)
                r["in_cooldown"] = now < pm.get("cooldown_until", 0.0)
                all_pm[name] = r
        return jsonify({"status": "ok", "price_memory": all_pm})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


def _segment_customer(profile: dict) -> str:
    """
    DB-FIRST v5 — Classify customer into a segment based on spend + visit_count.
    Segments: vip / regular / occasional / new
    """
    visits = profile.get("visit_count", 0)
    spend  = profile.get("total_spend", 0.0)
    avg    = profile.get("avg_ticket", 0.0)
    if visits >= 20 or spend >= 5000:
        return "vip"
    if visits >= 8 or spend >= 1500:
        return "regular"
    if visits >= 2:
        return "occasional"
    return "new"


# ── Phase 5: Customer profile endpoint ────────────────────────────────────────

@app.route("/customer-profile", methods=["GET"])
def customer_profile_route():
    """DB-FIRST v5 — Return spend intelligence + segment for a customer phone."""
    try:
        phone = request.args.get("phone", "").strip()
        if not phone:
            return jsonify({"status": "error",
                            "message": "phone query param required"}), 400
        if not _DB_AVAILABLE:
            return jsonify({"status": "error",
                            "message": "DB layer not available"}), 503
        profile = _db.get_customer_profile(phone)
        if profile is None:
            return jsonify({"status": "ok", "phone": phone, "profile": None,
                            "message": "Customer not found"})
        segment = _segment_customer(profile)
        try:
            conn = _db.get_db()
            recent = conn.execute(
                "SELECT service_name, SUM(qty) as qty, SUM(revenue) as rev "
                "FROM transactions WHERE customer_phone=? "
                "GROUP BY service_name ORDER BY rev DESC LIMIT 5",
                (phone,)
            ).fetchall()
            profile["top_services"] = [
                {"service": r["service_name"], "qty": int(r["qty"] or 0),
                 "revenue": round(float(r["rev"] or 0), 2)} for r in recent
            ]
        except Exception:
            profile["top_services"] = []
        profile["segment"] = segment
        return jsonify({"status": "ok", "phone": phone, "profile": profile})
    except Exception as e:
        log.error("customer-profile: %s", e)
        return jsonify({"status": "error", "message": str(e)}), 500


# ═══════════════════════════════════════════════════════════════════════════════
# DB-FIRST v5 — CONTROL LAYER  (optimize_system)
# Shapes demand proactively instead of just reacting.
# ═══════════════════════════════════════════════════════════════════════════════

def optimize_system() -> dict:
    """
    Control engine: runs after every background refresh cycle.
    Reads demand + pricing state from DB and emits adjustment signals.

    Returns a dict with three signal groups:
      price_adjustments   — services where auto-pricing should be applied now
      priority_adjustments— services whose display priority should change
      bundle_promotions   — pairs to promote as bundles in UI
    """
    if _is_failsafe() or not _DB_AVAILABLE:
        return {"price_adjustments": [], "priority_adjustments": [], "bundle_promotions": []}

    try:
        svc_map = _get_services()
        demand  = _ensure_demand_cache()
        bundles = _ensure_bundle_stats()
        ai      = _ensure_ai_cache()
    except Exception as e:
        log.error("optimize_system: cache fetch failed: %s", e)
        return {"price_adjustments": [], "priority_adjustments": [], "bundle_promotions": []}

    price_adjustments    = []
    priority_adjustments = []
    bundle_promotions    = []

    # ── 1. Price adjustments: high-confidence AI suggestions not in cooldown ──
    for svc_name, ins in ai.items():
        sp  = ins.get("suggested_price", 0)
        cp  = ins.get("current_price", 0)
        conf= ins.get("confidence", 0)
        if sp == cp or cp <= 0 or conf < 0.55:
            continue
        if _is_in_cooldown(svc_name):
            continue
        direction = "up" if sp > cp else "down"
        price_adjustments.append({
            "service":    svc_name,
            "current":    cp,
            "suggested":  sp,
            "confidence": conf,
            "direction":  direction,
            "reason":     ins.get("reason", ""),
        })

    # ── 2. Priority adjustments: boost rising-trend services; suppress falling ─
    for svc_name, d in demand.items():
        svc      = svc_map.get(svc_name)
        if not svc:
            continue
        trend    = d.get("trend", "—")
        conf     = d.get("confidence", 0)
        cur_pri  = svc.get("priority", 5.0)
        new_pri  = cur_pri

        if trend == "rising" and conf >= 0.4 and cur_pri > 3:
            new_pri = max(1.0, cur_pri - 1.0)   # promote
        elif trend == "falling" and conf >= 0.4 and cur_pri < 8:
            new_pri = min(9.0, cur_pri + 1.0)   # demote

        if new_pri != cur_pri:
            priority_adjustments.append({
                "service":      svc_name,
                "old_priority": cur_pri,
                "new_priority": new_pri,
                "trend":        trend,
            })

    # ── 3. Bundle promotions: high-confidence pairs not yet promoted ──────────
    for (a, b), stats in bundles.items():
        if stats.get("confidence", 0) >= 0.4 and stats.get("strength", 0) >= 0.05:
            bundle_promotions.append({
                "service_a":  a,
                "service_b":  b,
                "strength":   stats["strength"],
                "confidence": stats["confidence"],
                "count":      stats["count"],
            })

    bundle_promotions.sort(key=lambda x: x["confidence"], reverse=True)

    log.info(
        "optimize_system: %d price adj, %d priority adj, %d bundle promos",
        len(price_adjustments), len(priority_adjustments), len(bundle_promotions)
    )
    return {
        "price_adjustments":    price_adjustments,
        "priority_adjustments": priority_adjustments,
        "bundle_promotions":    bundle_promotions[:10],   # top 10 only
    }


@app.route("/optimize", methods=["GET"])
def optimize_route():
    """
    DB-FIRST v5 — Control layer endpoint.
    Returns system optimization signals. Safe to call at any time — read-only.
    Set apply=1 query param to auto-apply high-confidence price adjustments.
    """
    try:
        signals = optimize_system()
        apply   = request.args.get("apply", "0") == "1"

        applied = []
        if apply and not _is_failsafe():
            for adj in signals.get("price_adjustments", []):
                if adj["confidence"] >= 0.65:
                    svc = _get_services().get(adj["service"], {})
                    result = update_price(
                        adj["service"], adj["suggested"], svc.get("cost", 0),
                        source="optimize"
                    )
                    if result.get("status") == "ok":
                        applied.append(adj["service"])

        return jsonify({
            "status":  "ok",
            "signals": signals,
            "applied": applied,
            "failsafe":_is_failsafe(),
        })
    except Exception as e:
        log.error("optimize route: %s", e)
        return jsonify({"status": "error", "message": str(e)}), 500


# ── Phase 7: DB integrity endpoint ────────────────────────────────────────────

@app.route("/db-integrity", methods=["GET"])
def db_integrity():
    """
    Phase 7 — Validate DB connectivity and return row counts.
    Returns:
        db_connected:       bool
        services_count:     int
        transactions_count: int
        customers_count:    int
        last_price_event:   ISO timestamp | null
    """
    try:
        if not _DB_AVAILABLE:
            return jsonify({
                "db_connected":      False,
                "services_count":    0,
                "transactions_count":0,
                "customers_count":   0,
                "last_price_event":  None,
                "message":           "db.py not found — DB layer disabled",
            }), 503
        stats = _db.get_db_stats()
        code  = 200 if stats["db_connected"] else 503
        return jsonify(stats), code
    except Exception as e:
        log.error("db-integrity: %s", e)
        return jsonify({"status": "error", "message": str(e)}), 500


# ── Phase 7: Override intelligence endpoint ────────────────────────────────────

@app.route("/override-insights/<path:service>", methods=["GET"])
def override_insights(service: str):
    """
    Phase 7 — Return override analytics for a service.
    Strategic use: identify services where customers consistently negotiate,
    base price is too high, or bulk pricing opportunity exists.

    Returns:
        override_frequency:    fraction of transactions with override
        avg_override_price:    mean charged price when overridden
        avg_discount_pct:      mean discount % (discount_pct type only)
        margin_loss_pct:       avg margin loss vs base price
        discount_distribution: histogram in 5% bands
        strategic_signal:      "price_negotiation" | "bulk_opportunity" | "normal"
    """
    try:
        if not _DB_AVAILABLE:
            return jsonify({"status": "error",
                            "message": "DB layer not available"}), 503
        stats = _db.get_override_stats(service.strip())

        # Derive strategic signal from data
        freq        = stats.get("override_frequency", 0.0)
        avg_disc    = stats.get("avg_discount_pct")
        signal      = "normal"
        if freq >= 0.30:
            signal  = "price_negotiation"   # >30% of transactions use override
        elif avg_disc is not None and avg_disc >= 15:
            signal  = "bulk_opportunity"    # customers consistently want bulk rate

        return jsonify({
            "status":      "ok",
            **stats,
            "strategic_signal": signal,
        })
    except Exception as e:
        log.error("override-insights: %s", e)
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/recent-transactions", methods=["GET"])
def recent_transactions():
    """Return the N most recent transactions from DB for the sidebar feed."""
    try:
        if not _DB_AVAILABLE:
            return jsonify([]), 200
        limit = min(int(request.args.get("limit", 8)), 50)
        conn  = _db.get_db()
        rows  = conn.execute(
            "SELECT timestamp, service_name, qty, revenue, profit, payment_mode, customer_phone "
            "FROM transactions ORDER BY id DESC LIMIT ?",
            (limit,)
        ).fetchall()
        result = [
            {
                "timestamp":     row["timestamp"],
                "service_name":  row["service_name"],
                "qty":           row["qty"],
                "revenue":       round(float(row["revenue"] or 0), 2),
                "profit":        round(float(row["profit"]  or 0), 2),
                "payment_mode":  row["payment_mode"],
                "customer_phone":row["customer_phone"],
            }
            for row in rows
        ]
        return jsonify(result)
    except Exception as e:
        log.error("recent-transactions: %s", e)
        return jsonify([]), 200   # always return array — UI doesn't need to crash


@app.route("/db-stats", methods=["GET"])
def db_stats():
    """Live DB health — services, transactions, customers, today revenue/profit."""
    try:
        if not _DB_AVAILABLE:
            return jsonify({"status": "error", "message": "DB not available"}), 503
        conn  = _db.get_db()
        today = datetime.now().strftime("%Y-%m-%d")
        svc_count  = conn.execute("SELECT COUNT(*) FROM services").fetchone()[0]
        tx_count   = conn.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]
        cust_count = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
        row = conn.execute(
            "SELECT COALESCE(SUM(revenue),0), COALESCE(SUM(profit),0), COUNT(*) "
            "FROM transactions WHERE timestamp >= ?", (today,)
        ).fetchone()
        return jsonify({
            "status":        "ok",
            "use_db_read":   True,
            "auto_pricing":  AUTO_PRICING,
            "failsafe":      _is_failsafe(),
            "services":      svc_count,
            "transactions":  tx_count,
            "customers":     cust_count,
            "today_revenue": round(float(row[0]), 2),
            "today_profit":  round(float(row[1]), 2),
            "today_count":   int(row[2]),
        })
    except Exception as e:
        log.error("db-stats: %s", e)
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/validate-migration", methods=["GET"])
def validate_migration():
    """Cross-check DB service count. No Excel dependency."""
    try:
        if not _DB_AVAILABLE:
            return jsonify({"ok": False, "message": "DB not available"}), 503
        db_svcs  = _db.get_services_from_db()
        db_count = len(db_svcs)
        zero_price = [n for n, s in db_svcs.items()
                      if s.get("price", 0) == 0 and n != "Other"]
        neg_margin = [n for n, s in db_svcs.items()
                      if s.get("price", 0) > 0 and s.get("cost", 0) > s.get("price", 0)]
        ok = db_count > 0 and not neg_margin
        return jsonify({
            "ok":            ok,
            "db_count":      db_count,
            "zero_price":    zero_price,
            "neg_margin":    neg_margin,
            "use_db_read":   True,
            "message":       (f"PASS — {db_count} services in DB" if ok
                              else f"WARN — issues detected"),
        })
    except Exception as e:
        log.error("validate-migration: %s", e)
        return jsonify({"ok": False, "message": str(e)}), 500


@app.route("/export-excel", methods=["POST"])
def export_excel():
    """Rebuild MASTER sheet in data.xlsx from DB. Excel is export-only."""
    try:
        if not _DB_AVAILABLE:
            return jsonify({"ok": False, "error": "DB not available"}), 503
        svc = _db.get_services_from_db()
        if not svc:
            return jsonify({"ok": False, "error": "No services in DB"}), 500
        import shutil as _shutil
        ts_bak = datetime.now().strftime("%Y%m%d_%H%M%S")
        if os.path.exists(EXCEL_PATH):
            _shutil.copy2(EXCEL_PATH, EXCEL_PATH + f".bak_{ts_bak}")
        try:
            wb = load_workbook(EXCEL_PATH)
        except Exception:
            from openpyxl import Workbook as _WB
            wb = _WB()
        sheet_name = MASTER_SHEET
        if sheet_name not in wb.sheetnames:
            wb.create_sheet(sheet_name)
        ws = wb[sheet_name]
        hdr_row = MASTER_DATA_START - 1
        for col, h in enumerate(
            ["SERVICE NAME","CATEGORY","SELL RATE ₹","COST ₹","MARGIN ₹","MARGIN %","ROLE"],
            start=2
        ):
            ws.cell(row=hdr_row, column=col, value=h)
        for row_idx, (name, info) in enumerate(svc.items(), start=MASTER_DATA_START):
            sell = info.get("price", 0); cost = info.get("cost", 0)
            ws.cell(row=row_idx, column=2, value=name)
            ws.cell(row=row_idx, column=3, value=info.get("category", ""))
            ws.cell(row=row_idx, column=4, value=sell)
            ws.cell(row=row_idx, column=5, value=cost)
            ws.cell(row=row_idx, column=6, value=round(sell - cost, 2))
            ws.cell(row=row_idx, column=7, value=round((sell-cost)/sell*100, 2) if sell else 0)
            ws.cell(row=row_idx, column=8, value=info.get("role", ""))
        with _write_lock:
            wb.save(EXCEL_PATH)
        wb.close()
        log.info("export-excel: wrote %d services to %s", len(svc), EXCEL_PATH)
        return jsonify({"ok": True, "services_exported": len(svc)})
    except Exception as e:
        log.exception("export-excel failed")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/ping", methods=["GET"])
def ping():
    """Lightweight liveness check — used by UI topbar to show server status."""
    return jsonify({
        "status":  "ok",
        "uptime":  round(time.time() - _health.get("uptime_start", time.time()), 0),
        "failsafe": _is_failsafe(),
        "db":       _DB_AVAILABLE,
    }), 200


# ═══════════════════════════════════════════════════════════════════════════════
# UDHAAR (CREDIT LEDGER) API
# ═══════════════════════════════════════════════════════════════════════════════

@app.route("/udhaar/add", methods=["POST"])
def udhaar_add():
    """
    Record udhaar given to a customer (standalone debit entry, outside a sale).
    Body: { phone, amount, note?, name? }
    """
    if not _DB_AVAILABLE:
        return jsonify({"error": "Database unavailable"}), 503
    data   = request.get_json(force=True) or {}
    phone  = str(data.get("phone") or "").strip()
    amount = data.get("amount")
    note   = data.get("note") or None
    name   = str(data.get("name") or "").strip() or None

    if not phone or len(phone) < 10:
        return jsonify({"error": "Valid phone number required (min 10 digits)"}), 400
    try:
        amount = float(amount)
        if amount <= 0:
            raise ValueError
    except (TypeError, ValueError):
        return jsonify({"error": "amount must be a positive number"}), 400

    # Credit limit enforcement: reject if this debit would exceed the customer's limit
    try:
        _bal   = _db.get_customer_balance(phone)
        _cust  = _db.get_customer_profile(phone)
        _limit = float((_cust or {}).get("credit_limit") or 0)
        _cur_bal = float(_bal.get("balance", 0))
        if _limit > 0 and _cur_bal + amount > _limit:
            return jsonify({
                "error":           "Credit limit exceeded",
                "current_balance": round(_cur_bal, 2),
                "credit_limit":    round(_limit,   2),
                "requested":       round(amount,   2),
                "shortfall":       round((_cur_bal + amount) - _limit, 2),
            }), 400
    except Exception as _cle:
        log.warning("Credit limit check failed (non-critical, allowing): %s", _cle)

    entry_id = _db.add_udhaar_entry(phone, amount, "debit", note=note)
    if entry_id is None:
        return jsonify({"error": "Failed to record udhaar entry"}), 500

    # Log credit action to audit trail
    try:
        _db.log_system_event("credit_action", "customer", phone,
                             old_value=None, new_value=amount, operator="operator")
    except Exception:
        pass

    # Upsert customer record with name if provided
    try:
        _db.upsert_customer(phone, 0.0, name=name)
    except Exception as _ce:
        log.warning("upsert_customer in udhaar/add failed (non-critical): %s", _ce)

    balance_info = _db.get_customer_balance(phone)
    log.info("udhaar/add: phone=%s amount=%.2f id=%s", phone, amount, entry_id)
    return jsonify({
        "status":   "ok",
        "entry_id": entry_id,
        "phone":    phone,
        "name":     name,
        "amount":   amount,
        "type":     "debit",
        "balance":  balance_info,
    })


@app.route("/udhaar/pay", methods=["POST"])
def udhaar_pay():
    """
    Record payment received from a customer (credit entry).
    Body: { phone, amount, note? }
    """
    if not _DB_AVAILABLE:
        return jsonify({"error": "Database unavailable"}), 503
    data   = request.get_json(force=True) or {}
    phone  = str(data.get("phone") or "").strip()
    amount = data.get("amount")
    note   = data.get("note") or "Payment received"

    if not phone or len(phone) < 10:
        return jsonify({"error": "Valid phone number required (min 10 digits)"}), 400
    try:
        amount = float(amount)
        if amount <= 0:
            raise ValueError
    except (TypeError, ValueError):
        return jsonify({"error": "amount must be a positive number"}), 400

    balance_info = _db.get_customer_balance(phone)
    if balance_info.get("balance", 0) <= 0:
        # Warn but don't block — advance payment is valid in some cases
        log.warning("udhaar/pay: phone=%s has no outstanding balance, recording anyway", phone)

    entry_id = _db.add_udhaar_entry(phone, amount, "credit", note=note)
    if entry_id is None:
        return jsonify({"error": "Failed to record payment entry"}), 500

    new_balance = _db.get_customer_balance(phone)
    log.info("udhaar/pay: phone=%s amount=%.2f id=%s new_balance=%.2f",
             phone, amount, entry_id, new_balance.get("balance", 0))
    return jsonify({
        "status":   "ok",
        "entry_id": entry_id,
        "phone":    phone,
        "amount":   amount,
        "type":     "credit",
        "balance":  new_balance,
    })


@app.route("/udhaar/balance/<phone>", methods=["GET"])
def udhaar_balance(phone):
    """Return current udhaar balance for a customer."""
    if not _DB_AVAILABLE:
        return jsonify({"error": "Database unavailable"}), 503
    phone = str(phone).strip()
    if not phone or len(phone) < 10:
        return jsonify({"error": "Valid phone number required"}), 400
    return jsonify(_db.get_customer_balance(phone))


@app.route("/udhaar/history/<phone>", methods=["GET"])
def udhaar_history(phone):
    """Return full udhaar ledger history for a customer."""
    if not _DB_AVAILABLE:
        return jsonify({"error": "Database unavailable"}), 503
    phone = str(phone).strip()
    if not phone or len(phone) < 10:
        return jsonify({"error": "Valid phone number required"}), 400
    limit   = min(int(request.args.get("limit", 50)), 200)
    history = _db.get_udhaar_history(phone, limit=limit)
    balance = _db.get_customer_balance(phone)
    return jsonify({
        "phone":   phone,
        "balance": balance,
        "history": history,
    })


@app.route("/udhaar/debtors", methods=["GET"])
def udhaar_debtors():
    """Return top debtors list with aging analysis."""
    if not _DB_AVAILABLE:
        return jsonify({"error": "Database unavailable"}), 503
    limit   = min(int(request.args.get("limit", 20)), 100)
    debtors = _db.get_top_debtors(limit=limit)
    summary = _db.get_udhaar_summary()
    return jsonify({
        "summary": summary,
        "debtors": debtors,
    })


@app.route("/udhaar/summary", methods=["GET"])
def udhaar_summary_route():
    """Global udhaar overview: total outstanding, today's activity."""
    if not _DB_AVAILABLE:
        return jsonify({"error": "Database unavailable"}), 503
    return jsonify(_db.get_udhaar_summary())


@app.route("/export-udhaar-sheet", methods=["POST"])
def export_udhaar_sheet():
    """
    Rebuild the '💳 UDHAAR LEDGER' sheet in data.xlsx from live DB data.
    Called automatically after export-excel, or manually via POST.
    Thread-safe: uses the existing _write_lock.
    """
    if not _DB_AVAILABLE:
        return jsonify({"ok": False, "error": "Database unavailable"}), 503
    try:
        from openpyxl.styles import Font as _Font, Alignment as _Align, PatternFill as _Fill, Border as _Border, Side as _Side
        from openpyxl.utils import get_column_letter as _gcl

        # ── Pull data from DB ────────────────────────────────────────────────
        summary  = _db.get_udhaar_summary()
        debtors  = _db.get_top_debtors(limit=200)
        conn     = _db.get_db()
        all_rows = conn.execute(
            "SELECT u.timestamp, u.customer_phone, c.name AS customer_name, "
            "u.type, u.amount, u.note, u.reference_txn_id "
            "FROM udhaar_ledger u "
            "LEFT JOIN customers c ON c.phone = u.customer_phone "
            "ORDER BY u.timestamp DESC LIMIT 500"
        ).fetchall()

        # ── Style helpers ────────────────────────────────────────────────────
        def _fill(h): return _Fill("solid", fgColor=h)
        def _font(color="E6EDF3", bold=False, size=10):
            return _Font(name="Segoe UI", size=size, bold=bold, color=color)
        def _align(h="left", v="center"):
            return _Align(horizontal=h, vertical=v)
        def _bdr(color="30363D"):
            s = _Side(style="thin", color=color)
            return _Border(left=s, right=s, top=s, bottom=s)

        CLR_UDHAAR = "E6A817"; CLR_GREEN = "39D353"; CLR_RED = "F85149"
        CLR_MUTED  = "8B949E"; CLR_TEXT  = "E6EDF3"; CLR_CYAN = "39C5CF"
        BG_SURF    = "161B22"; BG_SURF2  = "21262D"; BG_UDHAAR = "2D2006"

        with _write_lock:
            wb = _load_wb(data_only=False)
            SHEET = "💳 UDHAAR LEDGER"

            # Create or clear the sheet
            if SHEET in wb.sheetnames:
                del wb[SHEET]
            ws = wb.create_sheet(SHEET)
            ws.sheet_properties.tabColor = "E6A817"

            # ── Column widths ────────────────────────────────────────────────
            for i, w in enumerate([3,22,18,14,14,14,14,12,20,32,20], 1):
                ws.column_dimensions[_gcl(i)].width = w

            # ── Title ────────────────────────────────────────────────────────
            ws.row_dimensions[2].height = 28
            ws.merge_cells("B2:J2")
            c = ws["B2"]
            c.value     = "⚡  CITYCYBER — UDHAAR LEDGER"
            c.font      = _font(CLR_UDHAAR, bold=True, size=15)
            c.fill      = _fill(BG_SURF)
            c.alignment = _align()

            ws.row_dimensions[3].height = 16
            ws.merge_cells("B3:J3")
            c = ws["B3"]
            c.value     = f"Exported: {datetime.now().strftime('%d %b %Y %H:%M')}  ·  {len(debtors)} customers with outstanding balance"
            c.font      = _font(CLR_MUTED, size=9)
            c.fill      = _fill(BG_SURF)
            c.alignment = _align()

            # ── Summary block (rows 5-8) ──────────────────────────────────
            for r in range(5, 10):
                ws.row_dimensions[r].height = 20

            pairs = [
                (5, "B", "TOTAL OUTSTANDING",  "D", f"₹{summary['total_outstanding']:,.2f}",  CLR_UDHAAR, "E", "TOTAL GIVEN",      "G", f"₹{summary['total_given']:,.2f}",     CLR_RED,    "H", "TOTAL RECOVERED", "J", f"₹{summary['total_recovered']:,.2f}", CLR_GREEN),
                (7, "B", "TODAY GIVEN",         "D", f"₹{summary['today_given']:,.2f}",        CLR_MUTED,  "E", "TODAY RECOVERED",  "G", f"₹{summary['today_recovered']:,.2f}", CLR_MUTED,  "H", "CUSTOMERS",       "J", str(summary['total_customers']),        CLR_MUTED),
            ]
            for rn, l1,h1,l2,v1,c1, l3,h3,l4,v3,c3, l5,h5,l6,v5,c5 in pairs:
                for (lbl_col, hdr, val_col, val, clr) in [(l1,h1,l2,v1,c1),(l3,h3,l4,v3,c3),(l5,h5,l6,v5,c5)]:
                    ws.merge_cells(f"{lbl_col}{rn}:{val_col}{rn}")
                    c = ws[f"{lbl_col}{rn}"]
                    c.value = hdr; c.font = _font(CLR_MUTED, size=8); c.fill = _fill(BG_SURF2); c.alignment = _align()
                    r2 = rn + 1
                    ws.merge_cells(f"{lbl_col}{r2}:{val_col}{r2}")
                    c = ws[f"{lbl_col}{r2}"]
                    c.value = val; c.font = _font(clr, bold=True, size=14); c.fill = _fill(BG_SURF2); c.alignment = _align()

            # ── Customer balances header (row 11) ────────────────────────
            ws.row_dimensions[11].height = 18
            hdrs = ["","PHONE","NAME","TOTAL GIVEN ₹","RECOVERED ₹","OUTSTANDING ₹","ENTRIES","RISK","LAST ACTIVITY","FIRST DEBIT","AGE (DAYS)"]
            for ci, h in enumerate(hdrs, 1):
                c = ws.cell(11, ci)
                c.value     = h
                c.font      = _font(CLR_MUTED, bold=True, size=9)
                c.fill      = _fill(BG_SURF2)
                c.border    = _bdr()
                c.alignment = _align("center" if ci > 3 else "left")

            # ── Customer balance data rows ────────────────────────────────
            R = 12
            for d in debtors:
                ws.row_dimensions[R].height = 17
                risk_clr = {"high": CLR_RED, "medium": "D29922", "low": CLR_GREEN}.get(d["risk"], CLR_MUTED)
                row_data = [
                    ("", None),
                    (d["phone"],                  CLR_TEXT),
                    (d.get("name") or "—",        CLR_MUTED),
                    (round(d["total_debit"],  2),  CLR_RED),
                    (round(d["total_credit"], 2),  CLR_GREEN),
                    (round(d["balance"],      2),  CLR_UDHAAR),
                    (d["entry_count"],             CLR_MUTED),
                    (d["risk"].upper(),            risk_clr),
                    (str(d.get("last_entry_ts",""))[:16], CLR_MUTED),
                    (str(d.get("first_debit_ts",""))[:16], CLR_MUTED),
                    (d.get("age_days", 0),         CLR_MUTED),
                ]
                for ci, (val, clr) in enumerate(row_data, 1):
                    c = ws.cell(R, ci)
                    c.value     = val
                    c.font      = _font(clr or CLR_MUTED, size=9)
                    c.fill      = _fill("1A2233" if R % 2 == 0 else BG_SURF2)
                    c.border    = _bdr()
                    c.alignment = _align("center" if ci > 3 else "left")
                    if isinstance(val, float) and ci in (4,5,6):
                        c.number_format = '₹#,##0.00'
                R += 1

            # ── Full ledger header ────────────────────────────────────────
            LEDGER_START = R + 3
            ws.row_dimensions[LEDGER_START].height = 28
            ws.merge_cells(f"B{LEDGER_START}:J{LEDGER_START}")
            c = ws.cell(LEDGER_START, 2)
            c.value     = "FULL TRANSACTION LEDGER  (latest 500 entries)"
            c.font      = _font(CLR_UDHAAR, bold=True, size=12)
            c.fill      = _fill(BG_SURF)
            c.alignment = _align()

            LH = LEDGER_START + 1
            ws.row_dimensions[LH].height = 16
            for ci, h in enumerate(["","TIMESTAMP","PHONE","NAME","TYPE","AMOUNT ₹","NOTE","","REF TXN",""],1):
                c = ws.cell(LH, ci)
                c.value = h; c.font = _font(CLR_MUTED, bold=True, size=9)
                c.fill = _fill(BG_SURF2); c.border = _bdr()
                c.alignment = _align("center" if ci in (5,6) else "left")

            # ── Ledger data rows ──────────────────────────────────────────
            LD = LH + 1
            for row in all_rows:
                ws.row_dimensions[LD].height = 15
                t = row["type"]
                amt_clr = CLR_RED if t == "debit" else CLR_GREEN
                type_lbl = "↑ DEBIT" if t == "debit" else "↓ PAID"
                vals = [
                    ("", None),
                    (str(row["timestamp"])[:16],          CLR_MUTED),
                    (row["customer_phone"],                CLR_TEXT),
                    (row["customer_name"] or "—",         CLR_MUTED),
                    (type_lbl,                            amt_clr),
                    (round(float(row["amount"]),2),       amt_clr),
                    (row["note"] or "—",                  CLR_MUTED),
                    ("", None),
                    (row["reference_txn_id"] or "—", CLR_MUTED),
                    ("", None),
                ]
                for ci, (val, clr) in enumerate(vals, 1):
                    c = ws.cell(LD, ci)
                    c.value     = val
                    c.font      = _font(clr or CLR_MUTED, size=9)
                    c.fill      = _fill("1A2233" if LD % 2 == 0 else BG_SURF2)
                    c.border    = _bdr()
                    c.alignment = _align("center" if ci == 4 else "left")
                    if isinstance(val, float):
                        c.number_format = '₹#,##0.00'
                LD += 1

            ws.freeze_panes = "B12"
            wb.save(EXCEL_PATH)
            wb.close()

        log.info("export-udhaar-sheet: wrote %d debtors, %d ledger rows", len(debtors), len(all_rows))
        return jsonify({"ok": True, "debtors": len(debtors), "ledger_rows": len(all_rows)})

    except Exception as e:
        log.error("export-udhaar-sheet: %s", e)
        return jsonify({"ok": False, "error": str(e)}), 500


# ═══════════════════════════════════════════════════════════════════════════════
# CONTROL LAYER API ENDPOINTS  (v6.0)
# ═══════════════════════════════════════════════════════════════════════════════

@app.route("/control/strategy", methods=["GET"])
def get_strategy_route():
    """Return current pricing strategy config."""
    if not _CTRL_AVAILABLE:
        return jsonify({"error": "Control layer not available"}), 503
    return jsonify(_ctrl.get_strategy())


@app.route("/control/strategy", methods=["POST"])
def set_strategy_route():
    """Update strategy config. Body: full or partial strategy JSON."""
    if not _CTRL_AVAILABLE:
        return jsonify({"error": "Control layer not available"}), 503
    data = request.get_json(force=True)
    if not data:
        return jsonify({"error": "Empty body"}), 400
    ok = _ctrl.save_strategy(data)
    return jsonify({"status": "ok" if ok else "error", "strategy": _ctrl.get_strategy()})


@app.route("/control/mode/<mode>", methods=["POST"])
def set_mode_route(mode):
    """Quick mode switch: extraction / expansion / balanced."""
    if not _CTRL_AVAILABLE:
        return jsonify({"error": "Control layer not available"}), 503
    return jsonify(_ctrl.set_mode(mode))


@app.route("/control/decisions", methods=["GET"])
def get_decisions_route():
    """Return recent pricing decision log."""
    if not _CTRL_AVAILABLE:
        return jsonify({"error": "Control layer not available"}), 503
    limit = min(int(request.args.get("limit", 50)), 200)
    return jsonify({
        "decisions": _ctrl.get_decision_log(limit),
        "stats":     _ctrl.get_decision_stats(),
    })


@app.route("/control/simulate/<path:service>", methods=["GET"])
def simulate_service_route(service):
    """Simulate pricing candidates for a single service. Read-only."""
    if not _CTRL_AVAILABLE:
        return jsonify({"error": "Control layer not available"}), 503
    svc_map = _get_services()
    svc = svc_map.get(service)
    if not svc:
        return jsonify({"error": f"Service not found: {service}"}), 404
    demand   = _ensure_demand_cache()
    d        = demand.get(service, {})
    elasticity = _estimate_elasticity(service, demand, svc)
    strategy = _ctrl.get_strategy()
    decision = _ctrl.choose_price(
        service_name=service,
        current_price=svc["price"],
        cost=svc["cost"],
        role=svc.get("role", "filler"),
        demand_data=d,
        elasticity=elasticity,
        strategy=strategy,
    )
    return jsonify({
        "service":       service,
        "elasticity":    elasticity,
        "demand":        d,
        "decision":      decision,
        "strategy_mode": strategy["mode"],
    })


@app.route("/control/run", methods=["POST"])
def run_control_route():
    """Manually trigger one control cycle. Returns summary."""
    if not _CTRL_AVAILABLE:
        return jsonify({"error": "Control layer not available"}), 503
    summary = _ctrl.run_control_cycle(
        svc_map=_get_services(),
        demand_cache=_ensure_demand_cache(),
        estimate_elasticity_fn=_estimate_elasticity,
        is_in_cooldown_fn=_is_in_cooldown,
        is_in_hysteresis_fn=_is_in_hysteresis,
        update_price_fn=update_price,
        is_failsafe_fn=_is_failsafe,
        auto_pricing=AUTO_PRICING,
    )
    return jsonify(summary)


# ═══════════════════════════════════════════════════════════════════════════════
# CREDIT INTELLIGENCE & CASHFLOW ENDPOINTS  (v6.0 upgrade)
# ═══════════════════════════════════════════════════════════════════════════════

@app.route("/metrics/cashflow", methods=["GET"])
def cashflow_metrics():
    """
    Real cashflow truth: booked vs collected vs outstanding credit.
    Returns today's figures + all-time udhaar summary.
    """
    if not _DB_AVAILABLE:
        return jsonify({"error": "Database unavailable"}), 503
    return jsonify(_db.get_cashflow_metrics())


@app.route("/udhaar/overdue", methods=["GET"])
def udhaar_overdue():
    """
    Return customers with outstanding udhaar older than `days` days (default 7).
    Query param: ?days=N
    """
    if not _DB_AVAILABLE:
        return jsonify({"error": "Database unavailable"}), 503
    days = max(1, min(int(request.args.get("days", 7)), 365))
    return jsonify({
        "days_threshold": days,
        "overdue":        _db.get_overdue_entries(days),
    })


@app.route("/customer/set-limit", methods=["POST"])
def set_credit_limit():
    """
    Set a credit limit for a customer.
    Body: { phone, limit }
    limit=0 means no limit enforced.
    """
    if not _DB_AVAILABLE:
        return jsonify({"error": "Database unavailable"}), 503
    data  = request.get_json(force=True) or {}
    phone = str(data.get("phone", "")).strip()
    limit = data.get("limit")
    if not phone or len(phone) < 10:
        return jsonify({"error": "Valid phone required (min 10 digits)"}), 400
    try:
        limit = float(limit)
        if limit < 0:
            raise ValueError
    except (TypeError, ValueError):
        return jsonify({"error": "limit must be a number >= 0"}), 400
    ok = _db.set_customer_credit_limit(phone, limit, operator="operator")
    return jsonify({
        "status":       "ok" if ok else "error",
        "phone":        phone,
        "credit_limit": limit,
    })


@app.route("/customer/risk/<phone>", methods=["GET"])
def customer_risk_route(phone):
    """
    Recompute and return risk level for a customer.
    Also returns current balance and credit limit.
    """
    if not _DB_AVAILABLE:
        return jsonify({"error": "Database unavailable"}), 503
    phone = str(phone).strip()
    if not phone or len(phone) < 10:
        return jsonify({"error": "Valid phone required"}), 400
    level   = _db.update_customer_risk(phone)
    profile = _db.get_customer_profile(phone) or {}
    balance = _db.get_customer_balance(phone)
    return jsonify({
        "phone":        phone,
        "risk_level":   level,
        "credit_limit": profile.get("credit_limit", 0),
        "balance":      balance,
    })


@app.route("/system/events", methods=["GET"])
def system_events_route():
    """
    Return system audit events.
    Query params: ?type=credit_action&entity=9876543210&limit=50
    """
    if not _DB_AVAILABLE:
        return jsonify({"error": "Database unavailable"}), 503
    limit  = min(int(request.args.get("limit", 50)), 500)
    etype  = request.args.get("type")  or None
    entity = request.args.get("entity") or None
    events = _db.get_system_events(limit=limit, event_type=etype, entity_id=entity)
    return jsonify({"count": len(events), "events": events})


@app.route("/customer/all-risk", methods=["GET"])
def all_customer_risk():
    """
    Bulk risk refresh: recompute risk_level for all customers with any udhaar history.
    Returns summary of risk distribution.
    """
    if not _DB_AVAILABLE:
        return jsonify({"error": "Database unavailable"}), 503
    try:
        conn  = _db.get_db()
        phones = [
            r["customer_phone"]
            for r in conn.execute(
                "SELECT DISTINCT customer_phone FROM udhaar_ledger"
            ).fetchall()
        ]
        distribution = {"low": 0, "medium": 0, "high": 0}
        for p in phones:
            lvl = _db.update_customer_risk(p)
            distribution[lvl] = distribution.get(lvl, 0) + 1
        return jsonify({
            "total_customers":   len(phones),
            "risk_distribution": distribution,
        })
    except Exception as e:
        log.error("all-customer-risk failed: %s", e)
        return jsonify({"error": str(e)}), 500


# ═══════════════════════════════════════════════════════════════════════════════
# MULTI-AGENT ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

@app.route("/agent/command", methods=["POST"])
def agent_command():
    if not _AGENTS_AVAILABLE:
        return jsonify({"error": "Agent system unavailable"}), 503
    body = request.get_json(silent=True) or {}
    text = (body.get("text") or "").strip()
    if not text:
        return jsonify({"error": "text required"}), 400
    import uuid
    sid = body.get("session_id") or str(uuid.uuid4())
    _commander.handle_command(text, session_id=sid)
    return jsonify({"ok": True, "session_id": sid})


@app.route("/agent/stream/<session_id>", methods=["GET"])
def agent_stream(session_id: str):
    if not _AGENTS_AVAILABLE:
        return jsonify({"error": "Agent system unavailable"}), 503
    bus = get_bus()
    q   = bus.subscribe(session_id)

    @stream_with_context
    def _generate():
        import queue as _queue
        yield ": connected\n\n"
        while True:
            try:
                msg = q.get(timeout=25)
                yield msg.to_sse()
                if msg.type in ("complete", "error"):
                    break
            except _queue.Empty:
                yield ": keepalive\n\n"

    resp = Response(_generate(), mimetype="text/event-stream")
    resp.headers["Cache-Control"]               = "no-cache"
    resp.headers["X-Accel-Buffering"]           = "no"
    resp.headers["Access-Control-Allow-Origin"] = "*"
    return resp


@app.route("/agent/agents", methods=["GET"])
def agent_list():
    agents = []
    for name, w in _agent_workers.items():
        agents.append({
            "name":        w.name,
            "description": w.description,
            "emoji":       w.emoji,
            "color":       w.color,
        })
    # add commander
    if _AGENTS_AVAILABLE and _commander:
        agents.insert(0, {
            "name":        _commander.name,
            "description": _commander.description,
            "emoji":       _commander.emoji,
            "color":       _commander.color,
        })
    return jsonify({"agents": agents, "available": _AGENTS_AVAILABLE})


@app.route("/agent/history", methods=["GET"])
def agent_history():
    if not _AGENTS_AVAILABLE:
        return jsonify({"history": []}), 200
    limit = min(int(request.args.get("limit", 100)), 600)
    return jsonify({"history": get_bus().history(limit)})


@app.route("/agent/clear", methods=["POST"])
def agent_clear():
    if _AGENTS_AVAILABLE:
        get_bus().clear()
    return jsonify({"ok": True})


# ═══════════════════════════════════════════════════════════════════════════════
# ENTRYPOINT
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    log.info("CityCyber POS v4 starting — Excel: %s", EXCEL_PATH)
    _startup_validation()
    app.run(host="0.0.0.0", port=5000, debug=False)