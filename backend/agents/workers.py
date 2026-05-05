"""
CityCyber POS — Worker Agents
Six specialised agents that each own a domain and query real POS data.
"""
from __future__ import annotations

import time
from datetime import date, timedelta
from typing import Any

from .base import BaseAgent


def _db():
    """Lazy import to avoid circular imports at module load time."""
    import db as _db_module
    return _db_module


# ── helpers ──────────────────────────────────────────────────────────────────
def _safe_query(sql: str, params: tuple = ()) -> list[dict]:
    try:
        conn = _db().get_db()
        cur = conn.execute(sql, params)
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()]
    except Exception:
        return []


def _scalar(sql: str, params: tuple = (), default: Any = None) -> Any:
    rows = _safe_query(sql, params)
    if rows:
        return list(rows[0].values())[0]
    return default


# ─────────────────────────────────────────────────────────────────────────────
class PricingAgent(BaseAgent):
    name        = "pricing"
    description = "Manages service pricing, GST rates, and cost analysis"
    emoji       = "💰"
    color       = "#2ecc71"

    def process(self, task: dict) -> dict:
        sid = task.get("session_id", "")
        self.thinking("Fetching current service pricing data…", sid)

        rows = _safe_query(
            "SELECT name, base_price, gst_rate, category FROM services ORDER BY category, name"
        )
        self.working(f"Found {len(rows)} services", session_id=sid)

        if not rows:
            self.response("No services found in the database.", session_id=sid)
            return {"summary": "No services found."}

        total_revenue_today = _scalar(
            "SELECT COALESCE(SUM(total),0) FROM sales WHERE date(created_at)=date('now')",
            default=0,
        )

        # group by category
        cats: dict[str, list] = {}
        for r in rows:
            cats.setdefault(r["category"] or "General", []).append(r)

        lines = [f"**Pricing Overview** ({len(rows)} services, today ₹{total_revenue_today:,.0f})"]
        for cat, svcs in cats.items():
            lines.append(f"\n*{cat}*")
            for s in svcs:
                gst = f" +{s['gst_rate']}% GST" if s.get("gst_rate") else ""
                lines.append(f"  • {s['name']}: ₹{s['base_price']:,.2f}{gst}")

        summary = "\n".join(lines)
        self.response(summary, data={"services": rows, "today_revenue": total_revenue_today}, session_id=sid)
        return {"summary": f"{len(rows)} services loaded, today ₹{total_revenue_today:,.0f} revenue"}


# ─────────────────────────────────────────────────────────────────────────────
class AnalyticsAgent(BaseAgent):
    name        = "analytics"
    description = "Revenue trends, top services, payment splits, hourly patterns"
    emoji       = "📊"
    color       = "#3498db"

    def process(self, task: dict) -> dict:
        sid = task.get("session_id", "")
        self.thinking("Running analytics queries…", sid)

        today = date.today().isoformat()
        week_ago = (date.today() - timedelta(days=7)).isoformat()

        today_rev = _scalar(
            "SELECT COALESCE(SUM(total),0) FROM sales WHERE date(created_at)=?",
            (today,), 0,
        )
        week_rev = _scalar(
            "SELECT COALESCE(SUM(total),0) FROM sales WHERE date(created_at)>=?",
            (week_ago,), 0,
        )
        total_txns = _scalar("SELECT COUNT(*) FROM sales", default=0)

        self.working(
            f"Today ₹{today_rev:,.0f} | Week ₹{week_rev:,.0f} | Total txns {total_txns}",
            data={"today": today_rev, "week": week_rev, "txns": total_txns},
            session_id=sid,
        )

        top = _safe_query(
            """SELECT si.service_name, SUM(si.qty) as units, SUM(si.line_total) as revenue
               FROM sale_items si JOIN sales s ON si.sale_id=s.id
               WHERE date(s.created_at)>=?
               GROUP BY si.service_name ORDER BY revenue DESC LIMIT 5""",
            (week_ago,),
        )
        self.working("Top services computed", data={"top": top}, session_id=sid)

        top_lines = "\n".join(
            f"  {i+1}. {r['service_name']}: ₹{r['revenue']:,.0f} ({r['units']} units)"
            for i, r in enumerate(top)
        ) or "  No data"

        summary = (
            f"**Analytics (7-day window)**\n"
            f"Today: ₹{today_rev:,.0f} | Week: ₹{week_rev:,.0f} | All-time txns: {total_txns}\n\n"
            f"**Top Services this week:**\n{top_lines}"
        )
        self.response(summary, data={"today": today_rev, "week": week_rev, "top": top}, session_id=sid)
        return {"summary": f"Today ₹{today_rev:,.0f}, week ₹{week_rev:,.0f}"}


# ─────────────────────────────────────────────────────────────────────────────
class InventoryAgent(BaseAgent):
    name        = "inventory"
    description = "Stock levels, low-stock alerts, reorder suggestions"
    emoji       = "📦"
    color       = "#e67e22"

    def process(self, task: dict) -> dict:
        sid = task.get("session_id", "")
        self.thinking("Checking inventory levels…", sid)

        items = _safe_query(
            "SELECT name, quantity, unit, low_stock_threshold FROM inventory ORDER BY quantity ASC"
        )
        if not items:
            self.response("No inventory items found.", session_id=sid)
            return {"summary": "No inventory data."}

        low = [i for i in items if i["quantity"] <= (i.get("low_stock_threshold") or 5)]
        self.working(
            f"{len(items)} items total, {len(low)} low-stock alerts",
            data={"low_count": len(low)},
            session_id=sid,
        )

        alert_lines = "\n".join(
            f"  ⚠️ {i['name']}: {i['quantity']} {i['unit'] or 'units'} (threshold {i.get('low_stock_threshold',5)})"
            for i in low
        ) or "  ✅ All items sufficiently stocked"

        summary = (
            f"**Inventory Status** ({len(items)} items)\n\n"
            f"**Low Stock Alerts ({len(low)}):**\n{alert_lines}"
        )
        self.response(summary, data={"items": items, "low": low}, session_id=sid)
        return {"summary": f"{len(items)} items, {len(low)} low-stock"}


# ─────────────────────────────────────────────────────────────────────────────
class CustomerAgent(BaseAgent):
    name        = "customer"
    description = "Customer visit patterns, loyalty insights, top spenders"
    emoji       = "👥"
    color       = "#9b59b6"

    def process(self, task: dict) -> dict:
        sid = task.get("session_id", "")
        self.thinking("Analysing customer patterns…", sid)

        week_ago = (date.today() - timedelta(days=7)).isoformat()

        # Customers table might not exist in all deployments — guard it
        top_customers = _safe_query(
            """SELECT customer_name, COUNT(*) as visits, SUM(total) as spent
               FROM sales WHERE customer_name IS NOT NULL AND customer_name!=''
               AND date(created_at)>=?
               GROUP BY customer_name ORDER BY spent DESC LIMIT 10""",
            (week_ago,),
        )

        unique_today = _scalar(
            """SELECT COUNT(DISTINCT customer_name) FROM sales
               WHERE date(created_at)=date('now') AND customer_name IS NOT NULL AND customer_name!=''""",
            default=0,
        )

        self.working(
            f"{unique_today} unique customers today, {len(top_customers)} tracked this week",
            session_id=sid,
        )

        top_lines = "\n".join(
            f"  {i+1}. {r['customer_name']}: ₹{r['spent']:,.0f} ({r['visits']} visits)"
            for i, r in enumerate(top_customers)
        ) or "  No named customers this week"

        summary = (
            f"**Customer Insights (7 days)**\n"
            f"Unique customers today: {unique_today}\n\n"
            f"**Top Spenders:**\n{top_lines}"
        )
        self.response(summary, data={"today_unique": unique_today, "top": top_customers}, session_id=sid)
        return {"summary": f"{unique_today} unique customers today"}


# ─────────────────────────────────────────────────────────────────────────────
class HealthAgent(BaseAgent):
    name        = "health"
    description = "System health, DB integrity, performance metrics"
    emoji       = "🩺"
    color       = "#1abc9c"

    def process(self, task: dict) -> dict:
        sid = task.get("session_id", "")
        self.thinking("Running system health checks…", sid)

        checks = {}

        # DB integrity
        integrity = _scalar("PRAGMA integrity_check", default="error")
        checks["db_integrity"] = integrity == "ok"
        self.working(f"DB integrity: {integrity}", session_id=sid)

        # WAL mode
        wal = _scalar("PRAGMA journal_mode", default="unknown")
        checks["wal_mode"] = wal == "wal"

        # Row counts
        for tbl in ("sales", "sale_items", "services", "inventory"):
            cnt = _scalar(f"SELECT COUNT(*) FROM {tbl}", default=-1)
            checks[f"table_{tbl}"] = cnt

        self.working("Row counts fetched", data=checks, session_id=sid)

        status_icon = "✅" if checks["db_integrity"] else "❌"
        wal_icon    = "✅" if checks["wal_mode"] else "⚠️"

        lines = [
            f"**System Health Report**",
            f"  {status_icon} DB Integrity: {integrity}",
            f"  {wal_icon} Journal Mode: {wal}",
            f"\n**Table Counts:**",
        ]
        for tbl in ("sales", "sale_items", "services", "inventory"):
            lines.append(f"  • {tbl}: {checks.get(f'table_{tbl}', '?')} rows")

        summary = "\n".join(lines)
        self.response(summary, data=checks, session_id=sid)
        return {"summary": f"DB {integrity}, WAL={wal}"}


# ─────────────────────────────────────────────────────────────────────────────
class ReportAgent(BaseAgent):
    name        = "report"
    description = "Generates daily/weekly summary reports"
    emoji       = "📋"
    color       = "#f39c12"

    def process(self, task: dict) -> dict:
        sid = task.get("session_id", "")
        self.thinking("Compiling report…", sid)

        today = date.today().isoformat()
        week_ago = (date.today() - timedelta(days=7)).isoformat()

        # Daily summary
        daily_rows = _safe_query(
            """SELECT date(created_at) as day,
                      COUNT(*) as txns,
                      SUM(total) as revenue
               FROM sales WHERE date(created_at)>=?
               GROUP BY day ORDER BY day DESC""",
            (week_ago,),
        )
        self.working(f"Compiled {len(daily_rows)} days of data", session_id=sid)

        avg_daily = (
            sum(r["revenue"] for r in daily_rows) / len(daily_rows)
            if daily_rows else 0
        )

        day_lines = "\n".join(
            f"  {r['day']}: ₹{r['revenue']:,.0f} ({r['txns']} txns)"
            for r in daily_rows
        ) or "  No sales data"

        summary = (
            f"**Weekly Report**\n"
            f"Avg daily revenue: ₹{avg_daily:,.0f}\n\n"
            f"**Daily Breakdown:**\n{day_lines}"
        )
        self.response(summary, data={"daily": daily_rows, "avg": avg_daily}, session_id=sid)
        return {"summary": f"7-day avg ₹{avg_daily:,.0f}/day"}
