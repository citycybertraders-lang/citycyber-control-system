# ═══════════════════════════════════════════════════════════════════════════════
# UDHAAR (CREDIT LEDGER) API  — Paste ALL blocks into app.py
# ═══════════════════════════════════════════════════════════════════════════════
# ADD THESE ENDPOINTS before the "if __name__ == '__main__':" line in app.py.
# Requires db.py with udhaar functions already added.
# Zero breaking changes — purely additive.
# ═══════════════════════════════════════════════════════════════════════════════


# ─────────────────────────────────────────────────────────────────────────────
# PATCH A: Modify add_transaction() to handle UDHAAR payment mode
# ─────────────────────────────────────────────────────────────────────────────
#
# In your add_transaction() function, FIND the block where DB writes happen
# after a successful Excel write (marked "# DB INTEGRATION: Insert transactions").
# REPLACE that entire try/except block with the one below.
#
# BEFORE (current):
# ─────────────────
#     try:
#         dt_obj = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
#         customer_phone = ...
#         for item in resolved:
#             db.insert_transaction(...)
#         if customer_phone:
#             db.upsert_customer(...)
#     except Exception as e:
#         log.warning("DB write failed (non-critical): %s", e)
#
# AFTER (replace with):
# ─────────────────────
#     try:
#         customer_phone = data.get("customer_phone") or None
#         for item in resolved:
#             txn_id = db.insert_transaction(
#                 timestamp=timestamp, service_name=item["name"],
#                 qty=item["qty"], revenue=item["revenue"],
#                 cost=item["cost"], profit=item["profit"],
#                 payment_mode=payment_mode, customer_phone=customer_phone,
#             )
#             # UDHAAR INTEGRATION: if payment is udhaar, create ledger entry
#             if payment_mode.upper() == "UDHAAR":
#                 if customer_phone:
#                     db.add_udhaar_entry(
#                         phone=customer_phone,
#                         amount=item["revenue"],
#                         entry_type="debit",
#                         note=f"Sale: {item['name']} ×{item['qty']}",
#                         reference_txn_id=txn_id,
#                     )
#                 else:
#                     log.warning("UDHAAR transaction recorded without customer phone")
#         if customer_phone:
#             db.upsert_customer(customer_phone, sum(i["revenue"] for i in resolved))
#     except Exception as e:
#         log.warning("DB write failed (non-critical): %s", e)
#
# NOTE: Also add "UDHAAR" to the allowed payment modes list if you have one.
# ─────────────────────────────────────────────────────────────────────────────


# ─────────────────────────────────────────────────────────────────────────────
# PATCH B: New API endpoints (add to app.py)
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/udhaar/add", methods=["POST"])
def udhaar_add():
    """
    Record udhaar given to a customer (debit entry).
    Body: { phone, amount, note? }
    """
    data   = request.get_json(force=True) or {}
    phone  = str(data.get("phone") or "").strip()
    amount = data.get("amount")
    note   = data.get("note") or None

    if not phone or len(phone) < 10:
        return jsonify({"error": "Valid phone number required (min 10 digits)"}), 400
    try:
        amount = float(amount)
        if amount <= 0:
            raise ValueError
    except (TypeError, ValueError):
        return jsonify({"error": "amount must be a positive number"}), 400

    entry_id = db.add_udhaar_entry(phone, amount, "debit", note=note)
    if entry_id is None:
        return jsonify({"error": "Failed to record udhaar entry"}), 500

    balance_info = db.get_customer_balance(phone)
    log.info("udhaar/add: phone=%s amount=%.2f id=%s", phone, amount, entry_id)
    return jsonify({
        "status":   "ok",
        "entry_id": entry_id,
        "phone":    phone,
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

    # Check current balance before accepting payment
    balance_info = db.get_customer_balance(phone)
    if balance_info["balance"] <= 0:
        return jsonify({
            "error": f"No outstanding balance for {phone}. Current balance: ₹{balance_info['balance']}",
            "balance": balance_info,
        }), 400

    entry_id = db.add_udhaar_entry(phone, amount, "credit", note=note)
    if entry_id is None:
        return jsonify({"error": "Failed to record payment entry"}), 500

    new_balance = db.get_customer_balance(phone)
    log.info("udhaar/pay: phone=%s amount=%.2f id=%s new_balance=%.2f",
             phone, amount, entry_id, new_balance["balance"])
    return jsonify({
        "status":      "ok",
        "entry_id":    entry_id,
        "phone":       phone,
        "amount":      amount,
        "type":        "credit",
        "balance":     new_balance,
    })


@app.route("/udhaar/balance/<phone>", methods=["GET"])
def udhaar_balance(phone):
    """Return current udhaar balance for a customer."""
    phone = str(phone).strip()
    if not phone or len(phone) < 10:
        return jsonify({"error": "Valid phone number required"}), 400
    return jsonify(db.get_customer_balance(phone))


@app.route("/udhaar/history/<phone>", methods=["GET"])
def udhaar_history(phone):
    """Return full udhaar ledger history for a customer."""
    phone = str(phone).strip()
    if not phone or len(phone) < 10:
        return jsonify({"error": "Valid phone number required"}), 400
    limit   = min(int(request.args.get("limit", 50)), 200)
    history = db.get_udhaar_history(phone, limit=limit)
    balance = db.get_customer_balance(phone)
    return jsonify({
        "phone":   phone,
        "balance": balance,
        "history": history,
    })


@app.route("/udhaar/debtors", methods=["GET"])
def udhaar_debtors():
    """Return top debtors list with aging analysis."""
    limit   = min(int(request.args.get("limit", 20)), 100)
    debtors = db.get_top_debtors(limit=limit)
    summary = db.get_udhaar_summary()
    return jsonify({
        "summary": summary,
        "debtors": debtors,
    })


@app.route("/udhaar/summary", methods=["GET"])
def udhaar_summary_route():
    """Global udhaar overview: total outstanding, today's activity."""
    return jsonify(db.get_udhaar_summary())
