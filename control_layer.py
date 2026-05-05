"""
CityCyber POS — Strategic Pricing Control Layer  (control_layer.py)
====================================================================
Closed-loop control engine: DATA → MODEL → SIMULATE → DECIDE → ACT → LOG → LEARN

Replaces reactive "suggest price" with intentional objective-function optimisation.
Plugs into app.py's background loop via optimize_system_v2().

Dependencies: db.py (same process), app.py exposes _get_services, _ensure_demand_cache,
              _estimate_elasticity, _is_in_cooldown, _is_in_hysteresis, update_price.
"""

import json
import math
import os
import logging
import threading
import time
from datetime import datetime, timedelta

log = logging.getLogger("citycyber.control")

# ═══════════════════════════════════════════════════════════════════════════════
# 1. STRATEGY CONFIG
# ═══════════════════════════════════════════════════════════════════════════════

STRATEGY_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "strategy.json")

_DEFAULT_STRATEGY = {
    "mode": "balanced",
    "weights": {
        "profit":   0.40,
        "revenue":  0.20,
        "growth":   0.20,
        "capture":  0.20,
    },
    "constraints": {
        "min_margin_pct":     10.0,      # never price below this margin
        "max_change_pct":     10.0,      # max single-step change
        "confidence_gate":    0.55,      # minimum confidence to act
        "cooldown_seconds":   300,       # reuse from app.py
        "max_daily_changes":  20,        # circuit breaker: max price changes per day
    },
    "policy_overrides": {
        "traffic": {
            "allow_low_margin": True,
            "min_margin_pct":   5.0,
            "bias": "volume",
        },
        "profit": {
            "allow_low_margin": False,
            "min_margin_pct":   20.0,
            "bias": "margin",
        },
        "bundle": {
            "allow_low_margin": True,
            "min_margin_pct":   8.0,
            "bias": "pair_optimise",
        },
    },
}

_strategy_lock = threading.Lock()
_strategy: dict = {}
_daily_change_count: int = 0
_daily_change_date: str = ""

# Decision log — ring buffer of last 200 decisions
_decision_log: list = []
_decision_lock = threading.Lock()
_MAX_DECISIONS = 200


def load_strategy() -> dict:
    """Load strategy from disk, falling back to defaults."""
    global _strategy
    with _strategy_lock:
        if _strategy:
            return dict(_strategy)
    try:
        if os.path.exists(STRATEGY_PATH):
            with open(STRATEGY_PATH, "r") as f:
                loaded = json.load(f)
            # Merge with defaults for any missing keys
            merged = _deep_merge(_DEFAULT_STRATEGY, loaded)
            with _strategy_lock:
                _strategy = merged
            log.info("Strategy loaded from %s: mode=%s", STRATEGY_PATH, merged["mode"])
            return dict(merged)
    except Exception as e:
        log.warning("Strategy load failed, using defaults: %s", e)
    with _strategy_lock:
        _strategy = dict(_DEFAULT_STRATEGY)
    return dict(_DEFAULT_STRATEGY)


def save_strategy(new_strategy: dict) -> bool:
    """Persist strategy to disk. Returns True on success."""
    global _strategy
    merged = _deep_merge(_DEFAULT_STRATEGY, new_strategy)
    try:
        tmp = STRATEGY_PATH + ".tmp"
        with open(tmp, "w") as f:
            json.dump(merged, f, indent=2)
        os.replace(tmp, STRATEGY_PATH)
        with _strategy_lock:
            _strategy = merged
        log.info("Strategy saved: mode=%s", merged["mode"])
        return True
    except Exception as e:
        log.error("Strategy save failed: %s", e)
        return False


def get_strategy() -> dict:
    """Return current strategy (loads from disk on first call)."""
    with _strategy_lock:
        if _strategy:
            return dict(_strategy)
    return load_strategy()


def set_mode(mode: str) -> dict:
    """Quick mode switch: extraction / expansion / balanced / custom."""
    presets = {
        "extraction": {"profit": 0.70, "revenue": 0.10, "growth": 0.10, "capture": 0.10},
        "expansion":  {"profit": 0.10, "revenue": 0.30, "growth": 0.35, "capture": 0.25},
        "balanced":   {"profit": 0.40, "revenue": 0.20, "growth": 0.20, "capture": 0.20},
    }
    weights = presets.get(mode)
    if not weights:
        return {"error": f"Unknown mode: {mode}. Valid: extraction, expansion, balanced"}
    s = get_strategy()
    s["mode"] = mode
    s["weights"] = weights
    save_strategy(s)
    return {"status": "ok", "mode": mode, "weights": weights}


def _deep_merge(base: dict, override: dict) -> dict:
    """Recursively merge override into base."""
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


# ═══════════════════════════════════════════════════════════════════════════════
# 2. SIMULATION ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

def generate_candidates(current_price: float, cost: float, role: str,
                        strategy: dict) -> list[float]:
    """
    Generate candidate prices for evaluation.
    Returns sorted list of unique prices including current.
    """
    max_change = strategy["constraints"]["max_change_pct"] / 100.0
    min_margin = strategy["constraints"]["min_margin_pct"] / 100.0

    # Role-specific policy override
    policy = strategy.get("policy_overrides", {}).get(role, {})
    if policy:
        min_margin = policy.get("min_margin_pct", min_margin * 100) / 100.0

    floor = cost * (1 + min_margin) if cost > 0 else current_price * 0.80

    steps = [-0.10, -0.07, -0.05, -0.03, 0.0, +0.03, +0.05, +0.07, +0.10]
    candidates = set()
    for s in steps:
        if abs(s) <= max_change:
            p = round(current_price * (1 + s), 1)
            if p >= floor:
                candidates.add(p)

    # Always include current price
    candidates.add(current_price)

    return sorted(candidates)


def simulate_price(service_name: str, candidate_price: float,
                   current_price: float, cost: float,
                   current_demand: float, elasticity: float,
                   trend: str = "stable") -> dict:
    """
    Forward-simulate outcome of setting candidate_price.
    Returns projected demand, revenue, profit, growth signal.
    """
    if current_price <= 0 or current_demand <= 0:
        return {
            "price": candidate_price,
            "demand": current_demand,
            "revenue": 0.0,
            "profit": 0.0,
            "growth": 0.0,
        }

    # Price-elasticity demand projection
    ratio = candidate_price / current_price
    ratio = max(0.70, min(1.50, ratio))
    projected_demand = current_demand * (ratio ** elasticity)

    # Trend modifier: rising trend boosts projected demand slightly
    trend_mult = {"rising": 1.05, "stable": 1.0, "falling": 0.95, "new": 1.0}.get(trend, 1.0)
    projected_demand *= trend_mult

    projected_demand = max(0.1, projected_demand)

    revenue = candidate_price * projected_demand
    profit = (candidate_price - cost) * projected_demand
    growth = (projected_demand - current_demand) / max(current_demand, 0.1)

    return {
        "price":   candidate_price,
        "demand":  round(projected_demand, 3),
        "revenue": round(revenue, 2),
        "profit":  round(profit, 2),
        "growth":  round(growth, 4),
    }


# ═══════════════════════════════════════════════════════════════════════════════
# 3. OBJECTIVE FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════

def compute_objective(
    sim: dict,
    strategy: dict,
    customer_signal: float = 0.0,
    credit_context: dict | None = None,
) -> float:
    """
    Weighted objective score for a simulated outcome.

    Σ [ w1 × normalised_profit
      + w2 × normalised_revenue
      + w3 × growth_signal
      + w4 × customer_capture
      - credit_penalty ]

    credit_context (optional):
        { outstanding: float, credit_limit: float, risk_level: 'low'|'medium'|'high' }
    When provided, high-credit-exposure customers reduce the score so the engine
    avoids aggressive upselling to customers who already owe money.
    """
    w = strategy["weights"]

    # Normalise profit and revenue to [0, 1] range via sigmoid-like scaling
    profit_score  = math.tanh(max(sim["profit"], 0) / 100.0)     # saturates around ₹300
    revenue_score = math.tanh(max(sim["revenue"], 0) / 200.0)    # saturates around ₹600
    growth_score  = max(-1.0, min(1.0, sim["growth"]))            # already [-1, 1]
    capture_score = max(0.0, min(1.0, customer_signal))           # [0, 1]

    score = (
        w["profit"]  * profit_score
        + w["revenue"] * revenue_score
        + w["growth"]  * growth_score
        + w["capture"] * capture_score
    )

    # Credit penalty: discount score when customer has significant outstanding balance
    if credit_context:
        risk_multipliers = {"low": 0.0, "medium": 0.05, "high": 0.20}
        limit       = max(float(credit_context.get("credit_limit") or 1000), 1.0)
        outstanding = float(credit_context.get("outstanding") or 0)
        risk_level  = credit_context.get("risk_level", "low")
        outstanding_ratio = min(outstanding / limit, 1.0)
        risk_mult         = risk_multipliers.get(risk_level, 0.0)
        credit_penalty    = outstanding_ratio * risk_mult
        score = max(0.0, score - credit_penalty)

    return round(score, 6)


# ═══════════════════════════════════════════════════════════════════════════════
# 4. DECISION ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

def choose_price(
    service_name: str,
    current_price: float,
    cost: float,
    role: str,
    demand_data: dict,
    elasticity: float,
    strategy: dict,
    bundle_partners: list | None = None,
    credit_context: dict | None = None,
) -> dict:
    """
    Core decision function. Generates candidates, simulates each,
    scores via objective function, selects best.

    credit_context (optional): { outstanding, credit_limit, risk_level }
      When provided, the objective function applies a credit penalty so the
      pricing engine is less aggressive toward customers with outstanding debt.

    Returns:
        {
            "action": "change" | "hold",
            "best_price": float,
            "current_price": float,
            "score": float,
            "confidence": float,
            "reason": str,
            "expected": { demand, revenue, profit, growth },
            "candidates_evaluated": int,
            "all_scores": [ { price, score } ... ],
        }
    """
    current_demand = max(
        demand_data.get("decay_score", 0),
        demand_data.get("demand_score", 0),
        0.1,
    )
    trend = demand_data.get("trend", "stable")
    dem_conf = demand_data.get("confidence", 0.0)
    total_tx = demand_data.get("total_tx", 0)

    # Customer capture signal: higher for rising trend + repeat customers
    customer_signal = 0.0
    if trend == "rising":
        customer_signal += 0.3
    if total_tx >= 20:
        customer_signal += 0.3
    if dem_conf >= 0.5:
        customer_signal += 0.2

    candidates = generate_candidates(current_price, cost, role, strategy)
    if not candidates:
        return _hold_decision(current_price, "no_candidates", 0.0)

    scored = []
    for p in candidates:
        sim = simulate_price(
            service_name, p, current_price, cost,
            current_demand, elasticity, trend,
        )
        obj = compute_objective(sim, strategy, customer_signal, credit_context)
        scored.append({
            "price": p,
            "score": obj,
            "sim": sim,
        })

    scored.sort(key=lambda x: -x["score"])
    best = scored[0]
    current_entry = next((s for s in scored if s["price"] == current_price), scored[-1])

    # Decision logic
    score_improvement = best["score"] - current_entry["score"]
    confidence = _compute_decision_confidence(
        score_improvement, dem_conf, total_tx, len(scored),
    )

    gate = strategy["constraints"]["confidence_gate"]

    if best["price"] == current_price or confidence < gate:
        return {
            "action": "hold",
            "best_price": current_price,
            "current_price": current_price,
            "score": current_entry["score"],
            "score_improvement": 0.0,
            "confidence": round(confidence, 3),
            "reason": "hold_optimal" if best["price"] == current_price else "low_confidence",
            "expected": current_entry["sim"],
            "candidates_evaluated": len(scored),
            "all_scores": [{"price": s["price"], "score": s["score"]} for s in scored[:5]],
        }

    return {
        "action": "change",
        "best_price": best["price"],
        "current_price": current_price,
        "score": best["score"],
        "score_improvement": round(score_improvement, 6),
        "confidence": round(confidence, 3),
        "reason": _classify_reason(best["price"], current_price, role, trend),
        "expected": best["sim"],
        "candidates_evaluated": len(scored),
        "all_scores": [{"price": s["price"], "score": s["score"]} for s in scored[:5]],
    }


def _hold_decision(price: float, reason: str, conf: float) -> dict:
    return {
        "action": "hold",
        "best_price": price,
        "current_price": price,
        "score": 0.0,
        "score_improvement": 0.0,
        "confidence": conf,
        "reason": reason,
        "expected": {"price": price, "demand": 0, "revenue": 0, "profit": 0, "growth": 0},
        "candidates_evaluated": 0,
        "all_scores": [],
    }


def _compute_decision_confidence(
    score_improvement: float,
    dem_conf: float,
    total_tx: int,
    num_candidates: int,
) -> float:
    """
    Decision confidence = f(score_improvement, demand_confidence, data_depth).
    Higher score improvement + more data → higher confidence.
    """
    # Base: how much better is the best vs current
    improvement_signal = min(1.0, score_improvement * 10)  # saturates at 0.1 improvement

    # Data quality
    data_signal = min(1.0, math.log1p(total_tx) / math.log1p(50))

    # Demand model confidence
    model_signal = dem_conf

    # Weighted combination
    confidence = (
        0.40 * improvement_signal
        + 0.35 * data_signal
        + 0.25 * model_signal
    )
    return max(0.0, min(1.0, confidence))


def _classify_reason(new_price: float, old_price: float, role: str, trend: str) -> str:
    direction = "uplift" if new_price > old_price else "reduction"
    return f"control_{role}_{direction}_{trend}"


# ═══════════════════════════════════════════════════════════════════════════════
# 5. DECISION LOGGING
# ═══════════════════════════════════════════════════════════════════════════════

def log_decision(service_name: str, decision: dict):
    """Append decision to ring buffer with timestamp."""
    entry = {
        "ts": datetime.now().isoformat(),
        "service": service_name,
        "action": decision["action"],
        "price": decision["best_price"],
        "current": decision["current_price"],
        "confidence": decision["confidence"],
        "reason": decision["reason"],
        "score": decision["score"],
        "score_improvement": decision.get("score_improvement", 0),
        "expected_profit": decision["expected"].get("profit", 0),
        "expected_revenue": decision["expected"].get("revenue", 0),
        "expected_demand": decision["expected"].get("demand", 0),
    }
    with _decision_lock:
        _decision_log.append(entry)
        if len(_decision_log) > _MAX_DECISIONS:
            _decision_log[:] = _decision_log[-_MAX_DECISIONS:]


def get_decision_log(limit: int = 50) -> list:
    """Return most recent decisions."""
    with _decision_lock:
        return list(_decision_log[-limit:])


def get_decision_stats() -> dict:
    """Aggregate stats from decision log."""
    with _decision_lock:
        log_copy = list(_decision_log)

    if not log_copy:
        return {"total": 0, "changes": 0, "holds": 0, "avg_confidence": 0}

    changes = [d for d in log_copy if d["action"] == "change"]
    holds = [d for d in log_copy if d["action"] == "hold"]
    avg_conf = sum(d["confidence"] for d in log_copy) / len(log_copy)
    avg_improvement = (
        sum(d.get("score_improvement", 0) for d in changes) / len(changes)
        if changes else 0
    )

    return {
        "total": len(log_copy),
        "changes": len(changes),
        "holds": len(holds),
        "avg_confidence": round(avg_conf, 3),
        "avg_score_improvement": round(avg_improvement, 4),
        "change_rate": round(len(changes) / len(log_copy), 3) if log_copy else 0,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# 6. DAILY CIRCUIT BREAKER
# ═══════════════════════════════════════════════════════════════════════════════

def _check_daily_limit(strategy: dict) -> bool:
    """Returns True if we can still make changes today."""
    global _daily_change_count, _daily_change_date
    today = datetime.now().strftime("%Y-%m-%d")
    if _daily_change_date != today:
        _daily_change_count = 0
        _daily_change_date = today
    return _daily_change_count < strategy["constraints"].get("max_daily_changes", 20)


def _increment_daily_count():
    global _daily_change_count
    _daily_change_count += 1


# ═══════════════════════════════════════════════════════════════════════════════
# 7. ORCHESTRATOR — plugs into app.py's background loop
# ═══════════════════════════════════════════════════════════════════════════════

def run_control_cycle(
    svc_map: dict,
    demand_cache: dict,
    estimate_elasticity_fn,
    is_in_cooldown_fn,
    is_in_hysteresis_fn,
    update_price_fn,
    is_failsafe_fn,
    auto_pricing: bool = True,
) -> dict:
    """
    Full control cycle. Called from app.py's _refresh_all().

    Sequence:
      1. Load strategy
      2. For each service: generate → simulate → score → decide
      3. Apply high-confidence changes (if auto_pricing)
      4. Log all decisions
      5. Return summary

    Returns:
        {
            "cycle_ts": str,
            "mode": str,
            "evaluated": int,
            "changes_proposed": int,
            "changes_applied": int,
            "changes_blocked": int,
            "decisions": [ ... ],
        }
    """
    if is_failsafe_fn():
        return {
            "cycle_ts": datetime.now().isoformat(),
            "mode": "failsafe",
            "evaluated": 0,
            "changes_proposed": 0,
            "changes_applied": 0,
            "changes_blocked": 0,
            "decisions": [],
        }

    strategy = get_strategy()
    decisions = []
    applied = 0
    blocked = 0
    proposed = 0

    for name, svc in svc_map.items():
        price = svc.get("price", 0)
        cost = svc.get("cost", 0)
        role = svc.get("role", "filler")

        if price <= 0:
            continue

        d = demand_cache.get(name, {})
        if d.get("total_tx", 0) < 3:
            # Not enough data — skip
            continue

        elasticity = estimate_elasticity_fn(name, demand_cache, svc)

        # Credit context: fetch per-customer data if DB is available
        # This makes pricing decisions aware of outstanding credit exposure
        _credit_ctx: dict | None = None
        try:
            import db as _ctrl_db
            _top_debtor = _ctrl_db.get_top_debtors(limit=1)  # cheapest way to get structure
            # For service-level decisions we use aggregate outstanding as context
            _cf = _ctrl_db.get_cashflow_metrics()
            if _cf.get("outstanding_udhaar", 0) > 0 and _cf.get("booked_revenue", 0) > 0:
                _credit_ctx = {
                    "outstanding":  _cf["outstanding_udhaar"],
                    "credit_limit": max(_cf["booked_revenue"] * 0.2, 1000),
                    "risk_level":   "high" if _cf["collection_rate"] < 0.7
                                    else ("medium" if _cf["collection_rate"] < 0.9 else "low"),
                }
        except Exception:
            _credit_ctx = None

        decision = choose_price(
            service_name=name,
            current_price=price,
            cost=cost,
            role=role,
            demand_data=d,
            elasticity=elasticity,
            strategy=strategy,
            credit_context=_credit_ctx,
        )

        log_decision(name, decision)
        decisions.append({"service": name, **decision})

        if decision["action"] == "change":
            proposed += 1

            if not auto_pricing:
                blocked += 1
                continue

            if is_in_cooldown_fn(name):
                blocked += 1
                decision["reason"] += "|cooldown"
                continue

            direction = "up" if decision["best_price"] > price else "down"
            if is_in_hysteresis_fn(name, direction):
                blocked += 1
                decision["reason"] += "|hysteresis"
                continue

            if not _check_daily_limit(strategy):
                blocked += 1
                decision["reason"] += "|daily_limit"
                continue

            # Execute
            result = update_price_fn(
                name, decision["best_price"], cost, source="control"
            )
            if result.get("status") == "ok":
                applied += 1
                _increment_daily_count()
                log.info(
                    "CONTROL applied %s: ₹%.1f→₹%.1f conf=%.2f reason=%s",
                    name, price, decision["best_price"],
                    decision["confidence"], decision["reason"],
                )
            else:
                blocked += 1
                log.warning(
                    "CONTROL blocked %s: %s", name, result.get("message", "unknown"),
                )

    summary = {
        "cycle_ts": datetime.now().isoformat(),
        "mode": strategy["mode"],
        "evaluated": len(decisions),
        "changes_proposed": proposed,
        "changes_applied": applied,
        "changes_blocked": blocked,
        "daily_changes_remaining": max(
            0, strategy["constraints"].get("max_daily_changes", 20) - _daily_change_count
        ),
    }
    log.info(
        "CONTROL CYCLE: mode=%s eval=%d proposed=%d applied=%d blocked=%d",
        strategy["mode"], len(decisions), proposed, applied, blocked,
    )
    return summary
