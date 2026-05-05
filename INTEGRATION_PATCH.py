# Control Layer Integration — Exact Changes to app.py
# ====================================================
# Apply these 4 patches to wire control_layer.py into your system.
# All changes are additive. Zero breaking changes.

# ─────────────────────────────────────────────────────────────────────
# PATCH 1: Import control layer (add after line ~111)
# ─────────────────────────────────────────────────────────────────────
# Find this block:
#   try:
#       import db as _db
#       _DB_AVAILABLE = True
#   except ImportError:
#       _DB_AVAILABLE = False
#
# Add AFTER it:

try:
    import control_layer as _ctrl
    _CTRL_AVAILABLE = True
except ImportError:
    _CTRL_AVAILABLE = False


# ─────────────────────────────────────────────────────────────────────
# PATCH 2: Replace optimize_system() call in _refresh_all() (~line 1483)
# ─────────────────────────────────────────────────────────────────────
# Find:
#     try:
#         optimize_system()
#     except Exception as _oe:
#         log.warning("optimize_system in refresh loop: %s", _oe)
#
# Replace with:

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


# ─────────────────────────────────────────────────────────────────────
# PATCH 3: Add control layer API endpoints (add before ENTRYPOINT)
# ─────────────────────────────────────────────────────────────────────

@app.route("/control/strategy", methods=["GET"])
def get_strategy_route():
    """Return current strategy config."""
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
    result = _ctrl.set_mode(mode)
    return jsonify(result)


@app.route("/control/decisions", methods=["GET"])
def get_decisions_route():
    """Return recent decision log."""
    if not _CTRL_AVAILABLE:
        return jsonify({"error": "Control layer not available"}), 503
    limit = min(int(request.args.get("limit", 50)), 200)
    return jsonify({
        "decisions": _ctrl.get_decision_log(limit),
        "stats": _ctrl.get_decision_stats(),
    })


@app.route("/control/simulate/<service>", methods=["GET"])
def simulate_service_route(service):
    """Simulate pricing candidates for a single service. Read-only."""
    if not _CTRL_AVAILABLE:
        return jsonify({"error": "Control layer not available"}), 503
    svc_map = _get_services()
    svc = svc_map.get(service)
    if not svc:
        return jsonify({"error": f"Service not found: {service}"}), 404
    demand = _ensure_demand_cache()
    d = demand.get(service, {})
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
        "service": service,
        "elasticity": elasticity,
        "demand": d,
        "decision": decision,
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
