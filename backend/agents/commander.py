"""
CityCyber POS — CommanderAgent
Top-level orchestrator. Parses user intent, routes tasks to workers,
aggregates results, and streams a final coherent reply to the user.
"""
from __future__ import annotations

import threading
import time
import uuid
from typing import Dict, List

from .base import BaseAgent
from .bus import get_bus, AgentMessage


# Intent → worker name + extracted params
_INTENTS: List[dict] = [
    {
        "keywords": ["price", "pricing", "gst", "cost", "rate", "charge"],
        "agent": "pricing",
        "label": "Pricing",
    },
    {
        "keywords": ["analytics", "revenue", "sales", "report", "trend", "top", "earning"],
        "agent": "analytics",
        "label": "Analytics",
    },
    {
        "keywords": ["stock", "inventory", "item", "product", "low", "restock"],
        "agent": "inventory",
        "label": "Inventory",
    },
    {
        "keywords": ["customer", "client", "visit", "loyalty", "frequent"],
        "agent": "customer",
        "label": "Customer",
    },
    {
        "keywords": ["health", "system", "status", "database", "performance", "check"],
        "agent": "health",
        "label": "Health",
    },
]


def _detect_intent(text: str) -> List[str]:
    """Return list of agent names that match keywords in text."""
    low = text.lower()
    matched = []
    for intent in _INTENTS:
        if any(kw in low for kw in intent["keywords"]):
            matched.append(intent["agent"])
    return matched or ["analytics"]  # default to analytics if nothing matched


class CommanderAgent(BaseAgent):
    name        = "commander"
    description = "Master orchestrator — routes tasks and aggregates results"
    emoji       = "🎖️"
    color       = "#e63946"

    def __init__(self, worker_registry: dict):
        super().__init__()
        self._workers: Dict[str, BaseAgent] = worker_registry

    # ── public entry ─────────────────────────────────────────────────────
    def handle_command(self, user_text: str, session_id: str = "") -> None:
        """Non-blocking: spawns a thread to process the command."""
        sid = session_id or str(uuid.uuid4())
        t = threading.Thread(
            target=self._run, args=(user_text, sid), daemon=True
        )
        t.start()

    def process(self, task: dict) -> dict:
        """Synchronous version (used internally)."""
        self._run(task.get("text", ""), task.get("session_id", ""))
        return {"ok": True}

    # ── internal orchestration ───────────────────────────────────────────
    def _run(self, user_text: str, session_id: str) -> None:
        bus = get_bus()

        # 1. Acknowledge
        bus.publish(AgentMessage(
            agent="commander", type="user",
            message=user_text, to="broadcast", session_id=session_id,
        ))

        self.thinking(f"Analysing request: \"{user_text[:80]}\"", session_id)
        time.sleep(0.2)

        # 2. Detect which agents to call
        targets = _detect_intent(user_text)
        agent_labels = [t.title() for t in targets]
        self.working(
            f"Routing to: {', '.join(agent_labels)}",
            data={"agents": targets},
            session_id=session_id,
        )

        # 3. Dispatch to each worker (parallel threads)
        results: Dict[str, dict] = {}
        lock = threading.Lock()
        threads = []

        def _call(name: str):
            worker = self._workers.get(name)
            if not worker:
                return
            self.talk_to(name, f"Handle: {user_text}", session_id=session_id)
            try:
                result = worker.process({"text": user_text, "session_id": session_id})
                with lock:
                    results[name] = result
            except Exception as exc:
                with lock:
                    results[name] = {"error": str(exc)}

        for agent_name in targets:
            th = threading.Thread(target=_call, args=(agent_name,), daemon=True)
            threads.append(th)
            th.start()

        for th in threads:
            th.join(timeout=15)

        # 4. Build consolidated answer
        parts = []
        for name, res in results.items():
            if "error" in res:
                parts.append(f"**{name.title()}**: ⚠️ {res['error']}")
            elif "summary" in res:
                parts.append(f"**{name.title()}**: {res['summary']}")

        final = "\n\n".join(parts) if parts else "All agents completed. No data to display."
        self.complete(final, data={"results": results}, session_id=session_id)
