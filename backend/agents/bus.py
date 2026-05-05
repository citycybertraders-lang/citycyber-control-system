"""
CityCyber POS — Multi-Agent Message Bus
Thread-safe pub/sub with SSE streaming support.
Every agent message flows through here so the frontend can watch live.
"""
from __future__ import annotations

import json
import queue
import threading
import time
from dataclasses import dataclass, asdict, field
from typing import Dict, List, Optional


@dataclass
class AgentMessage:
    agent:      str            # sender name
    type:       str            # thinking|working|inter_agent|response|complete|error|info|user
    message:    str
    to:         str  = "broadcast"   # target agent OR "broadcast" OR "user"
    data:       dict = field(default_factory=dict)
    session_id: str  = ""
    timestamp:  float = field(default_factory=time.time)

    def to_sse(self) -> str:
        d = asdict(self)
        d["ts"] = time.strftime("%H:%M:%S", time.localtime(d["timestamp"]))
        return f"data: {json.dumps(d, ensure_ascii=False)}\n\n"


class MessageBus:
    """Central nervous system of the agent network."""

    def __init__(self, max_history: int = 600):
        self._lock        = threading.Lock()
        self._subs: Dict[str, queue.Queue] = {}
        self._history: List[AgentMessage]  = []
        self._max_history  = max_history

    # ── subscription ────────────────────────────────────────────────────
    def subscribe(self, session_id: str) -> queue.Queue:
        q = queue.Queue(maxsize=300)
        with self._lock:
            self._subs[session_id] = q
            # replay last 30 messages for late joiners
            for msg in self._history[-30:]:
                try:
                    q.put_nowait(msg)
                except queue.Full:
                    pass
        return q

    def unsubscribe(self, session_id: str) -> None:
        with self._lock:
            self._subs.pop(session_id, None)

    # ── publish ──────────────────────────────────────────────────────────
    def publish(self, msg: AgentMessage) -> None:
        with self._lock:
            self._history.append(msg)
            if len(self._history) > self._max_history:
                self._history = self._history[-self._max_history:]
            for q in list(self._subs.values()):
                try:
                    q.put_nowait(msg)
                except queue.Full:
                    pass

    # ── history ──────────────────────────────────────────────────────────
    def history(self, limit: int = 100) -> List[dict]:
        with self._lock:
            return [asdict(m) for m in self._history[-limit:]]

    def clear(self) -> None:
        with self._lock:
            self._history.clear()

    # ── agent registry ───────────────────────────────────────────────────
    def active_subscribers(self) -> List[str]:
        with self._lock:
            return list(self._subs.keys())


# ── singleton ────────────────────────────────────────────────────────────────
_bus = MessageBus()


def get_bus() -> MessageBus:
    return _bus
