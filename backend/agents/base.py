"""
CityCyber POS — BaseAgent
Every worker agent inherits from this class.
"""
from __future__ import annotations

import time
from abc import ABC, abstractmethod
from typing import Any

from .bus import AgentMessage, get_bus


class BaseAgent(ABC):
    name:        str = "agent"
    description: str = ""
    emoji:       str = "🤖"
    color:       str = "#6c757d"

    def __init__(self):
        self._bus = get_bus()

    # ── emit helpers ─────────────────────────────────────────────────────
    def _emit(self, type_: str, message: str, to: str = "broadcast",
              data: dict | None = None, session_id: str = "") -> None:
        self._bus.publish(AgentMessage(
            agent=self.name,
            type=type_,
            message=message,
            to=to,
            data=data or {},
            session_id=session_id,
        ))

    def thinking(self, msg: str, session_id: str = "") -> None:
        self._emit("thinking", msg, session_id=session_id)

    def working(self, msg: str, data: dict | None = None, session_id: str = "") -> None:
        self._emit("working", msg, data=data, session_id=session_id)

    def response(self, msg: str, data: dict | None = None,
                 to: str = "user", session_id: str = "") -> None:
        self._emit("response", msg, to=to, data=data or {}, session_id=session_id)

    def error(self, msg: str, session_id: str = "") -> None:
        self._emit("error", msg, session_id=session_id)

    def complete(self, msg: str, data: dict | None = None, session_id: str = "") -> None:
        self._emit("complete", msg, data=data or {}, session_id=session_id)

    def talk_to(self, agent_name: str, msg: str, data: dict | None = None,
                session_id: str = "") -> None:
        self._emit("inter_agent", msg, to=agent_name, data=data or {}, session_id=session_id)

    # ── main entry point ─────────────────────────────────────────────────
    @abstractmethod
    def process(self, task: dict) -> dict:
        """Execute a task and return a result dict."""
