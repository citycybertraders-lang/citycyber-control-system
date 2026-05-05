import { useState, useEffect, useCallback } from "react";

const API = "http://localhost:5000";

const MODES = {
  extraction: { label: "Extraction", desc: "Max profit, min growth", color: "#dc2626", weights: { profit: 0.70, revenue: 0.10, growth: 0.10, capture: 0.10 } },
  balanced:   { label: "Balanced",   desc: "Equal weight all axes",  color: "#2563eb", weights: { profit: 0.40, revenue: 0.20, growth: 0.20, capture: 0.20 } },
  expansion:  { label: "Expansion",  desc: "Max growth & capture",   color: "#16a34a", weights: { profit: 0.10, revenue: 0.30, growth: 0.35, capture: 0.25 } },
};

function Badge({ children, variant = "default" }) {
  const colors = {
    default: "background: #1e293b; color: #94a3b8;",
    change: "background: #1e3a5f; color: #38bdf8;",
    hold: "background: #1a1a2e; color: #64748b;",
    up: "background: #052e16; color: #4ade80;",
    down: "background: #450a0a; color: #f87171;",
    active: "background: #1e3a5f; color: #38bdf8; border: 1px solid #38bdf8;",
  };
  return <span style={{ cssText: colors[variant] || colors.default }} className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium" dangerouslySetInnerHTML={{ __html: undefined }}>{children}</span>;
}

function WeightBar({ label, value, color }) {
  return (
    <div style={{ marginBottom: 8 }}>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, color: "#94a3b8", marginBottom: 3 }}>
        <span>{label}</span>
        <span style={{ fontFamily: "JetBrains Mono, monospace" }}>{(value * 100).toFixed(0)}%</span>
      </div>
      <div style={{ height: 6, background: "#1e293b", borderRadius: 3, overflow: "hidden" }}>
        <div style={{ width: `${value * 100}%`, height: "100%", background: color, borderRadius: 3, transition: "width 0.6s cubic-bezier(0.4,0,0.2,1)" }} />
      </div>
    </div>
  );
}

function StatCard({ label, value, sub, accent }) {
  return (
    <div style={{ background: "#0f172a", border: "1px solid #1e293b", borderRadius: 8, padding: "14px 16px", flex: 1, minWidth: 120 }}>
      <div style={{ fontSize: 11, color: "#64748b", textTransform: "uppercase", letterSpacing: 1, marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: 24, fontWeight: 700, color: accent || "#e2e8f0", fontFamily: "JetBrains Mono, monospace" }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: "#475569", marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function DecisionRow({ d }) {
  const isChange = d.action === "change";
  const dir = d.price > d.current ? "up" : d.price < d.current ? "down" : "hold";
  return (
    <tr style={{ borderBottom: "1px solid #1e293b" }}>
      <td style={{ padding: "8px 10px", fontSize: 12, color: "#cbd5e1", maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{d.service}</td>
      <td style={{ padding: "8px 6px", textAlign: "center" }}>
        <span style={{ fontSize: 11, padding: "2px 8px", borderRadius: 4, background: isChange ? "#1e3a5f" : "#1a1a2e", color: isChange ? "#38bdf8" : "#64748b" }}>{d.action}</span>
      </td>
      <td style={{ padding: "8px 6px", textAlign: "right", fontFamily: "JetBrains Mono, monospace", fontSize: 12, color: "#94a3b8" }}>₹{d.current?.toFixed(1)}</td>
      <td style={{ padding: "8px 6px", textAlign: "right", fontFamily: "JetBrains Mono, monospace", fontSize: 12, color: dir === "up" ? "#4ade80" : dir === "down" ? "#f87171" : "#64748b" }}>
        {isChange ? `₹${d.price?.toFixed(1)}` : "—"}
      </td>
      <td style={{ padding: "8px 6px", textAlign: "right", fontFamily: "JetBrains Mono, monospace", fontSize: 12, color: "#94a3b8" }}>{(d.confidence * 100).toFixed(0)}%</td>
      <td style={{ padding: "8px 6px", fontSize: 11, color: "#475569", maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{d.reason}</td>
      <td style={{ padding: "8px 6px", fontSize: 10, color: "#334155" }}>{d.ts?.slice(11, 19)}</td>
    </tr>
  );
}

export default function ControlDashboard() {
  const [strategy, setStrategy] = useState(null);
  const [decisions, setDecisions] = useState({ decisions: [], stats: {} });
  const [dbStats, setDbStats] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [simService, setSimService] = useState("");
  const [simResult, setSimResult] = useState(null);
  const [simLoading, setSimLoading] = useState(false);
  const [runResult, setRunResult] = useState(null);
  const [activeTab, setActiveTab] = useState("overview");

  const fetchAll = useCallback(async () => {
    try {
      const [sRes, dRes, dbRes] = await Promise.all([
        fetch(`${API}/control/strategy`).then(r => r.ok ? r.json() : null).catch(() => null),
        fetch(`${API}/control/decisions?limit=100`).then(r => r.ok ? r.json() : null).catch(() => null),
        fetch(`${API}/db-stats`).then(r => r.ok ? r.json() : null).catch(() => null),
      ]);
      if (sRes) setStrategy(sRes);
      if (dRes) setDecisions(dRes);
      if (dbRes) setDbStats(dbRes);
      setError(null);
    } catch (e) {
      setError("Cannot reach server at " + API);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchAll();
    const iv = setInterval(fetchAll, 15000);
    return () => clearInterval(iv);
  }, [fetchAll]);

  const switchMode = async (mode) => {
    try {
      const res = await fetch(`${API}/control/mode/${mode}`, { method: "POST" });
      if (res.ok) {
        const data = await res.json();
        setStrategy(prev => ({ ...prev, mode, weights: data.weights }));
      }
    } catch {}
  };

  const runCycle = async () => {
    try {
      const res = await fetch(`${API}/control/run`, { method: "POST" });
      if (res.ok) {
        const data = await res.json();
        setRunResult(data);
        setTimeout(fetchAll, 1000);
      }
    } catch {}
  };

  const simulate = async () => {
    if (!simService.trim()) return;
    setSimLoading(true);
    try {
      const res = await fetch(`${API}/control/simulate/${encodeURIComponent(simService.trim())}`);
      if (res.ok) setSimResult(await res.json());
      else setSimResult({ error: "Service not found" });
    } catch { setSimResult({ error: "Request failed" }); }
    finally { setSimLoading(false); }
  };

  const stats = decisions.stats || {};
  const recentDecisions = (decisions.decisions || []).slice().reverse().slice(0, 50);
  const changesOnly = recentDecisions.filter(d => d.action === "change");

  if (loading) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#030712", color: "#64748b", fontFamily: "'JetBrains Mono', monospace" }}>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 14, marginBottom: 8 }}>CONNECTING TO CONTROL LAYER...</div>
          <div style={{ width: 40, height: 2, background: "#1e293b", borderRadius: 1, margin: "0 auto", overflow: "hidden" }}>
            <div style={{ width: "60%", height: "100%", background: "#38bdf8", animation: "pulse 1s infinite" }} />
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: "100vh", background: "#030712", color: "#e2e8f0", fontFamily: "'Inter', -apple-system, sans-serif" }}>
      <link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet" />
      <style>{`
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: #0f172a; }
        ::-webkit-scrollbar-thumb { background: #1e293b; border-radius: 3px; }
        @keyframes pulse { 0%, 100% { opacity: 0.4; } 50% { opacity: 1; } }
        button { cursor: pointer; border: none; font-family: inherit; }
        button:active { transform: scale(0.97); }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 8px 10px; font-size: 10px; color: #475569; text-transform: uppercase; letter-spacing: 1px; border-bottom: 1px solid #1e293b; }
      `}</style>

      {/* HEADER */}
      <div style={{ borderBottom: "1px solid #1e293b", padding: "16px 24px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 8, height: 8, borderRadius: "50%", background: error ? "#ef4444" : "#4ade80", boxShadow: error ? "0 0 8px #ef4444" : "0 0 8px #4ade80" }} />
          <div>
            <div style={{ fontSize: 15, fontWeight: 700, letterSpacing: -0.5 }}>CITYCYBER CONTROL LAYER</div>
            <div style={{ fontSize: 11, color: "#475569" }}>Strategic Pricing Engine v5.0</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          {["overview", "decisions", "simulate"].map(tab => (
            <button key={tab} onClick={() => setActiveTab(tab)} style={{
              padding: "6px 14px", fontSize: 12, fontWeight: 500, borderRadius: 6,
              background: activeTab === tab ? "#1e3a5f" : "transparent",
              color: activeTab === tab ? "#38bdf8" : "#64748b",
              border: activeTab === tab ? "1px solid #1e3a5f" : "1px solid transparent",
            }}>
              {tab.toUpperCase()}
            </button>
          ))}
        </div>
      </div>

      {error && (
        <div style={{ margin: "16px 24px", padding: "10px 16px", background: "#450a0a", border: "1px solid #7f1d1d", borderRadius: 6, fontSize: 12, color: "#fca5a5" }}>
          {error}. Ensure app.py is running with control_layer.py integrated.
        </div>
      )}

      <div style={{ padding: "20px 24px" }}>
        {/* OVERVIEW TAB */}
        {activeTab === "overview" && (
          <div>
            {/* Stats Row */}
            <div style={{ display: "flex", gap: 12, marginBottom: 20, flexWrap: "wrap" }}>
              <StatCard label="Mode" value={strategy?.mode?.toUpperCase() || "—"} accent={MODES[strategy?.mode]?.color || "#94a3b8"} />
              <StatCard label="Decisions" value={stats.total || 0} sub={`${stats.changes || 0} changes / ${stats.holds || 0} holds`} />
              <StatCard label="Change Rate" value={`${((stats.change_rate || 0) * 100).toFixed(0)}%`} accent="#f59e0b" />
              <StatCard label="Avg Confidence" value={`${((stats.avg_confidence || 0) * 100).toFixed(0)}%`} accent="#8b5cf6" />
              <StatCard label="DB Services" value={dbStats?.services || "—"} sub={`${dbStats?.transactions || 0} tx total`} />
              <StatCard label="Today" value={`₹${dbStats?.today_revenue?.toFixed(0) || 0}`} sub={`₹${dbStats?.today_profit?.toFixed(0) || 0} profit`} accent="#4ade80" />
            </div>

            {/* Strategy Mode Selector */}
            <div style={{ background: "#0f172a", border: "1px solid #1e293b", borderRadius: 10, padding: 20, marginBottom: 20 }}>
              <div style={{ fontSize: 11, color: "#475569", textTransform: "uppercase", letterSpacing: 1.5, marginBottom: 14 }}>STRATEGY MODE</div>
              <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
                {Object.entries(MODES).map(([key, m]) => (
                  <button key={key} onClick={() => switchMode(key)} style={{
                    flex: 1, padding: "14px 16px", borderRadius: 8, textAlign: "left",
                    background: strategy?.mode === key ? `${m.color}15` : "#0a0a1a",
                    border: strategy?.mode === key ? `1.5px solid ${m.color}` : "1.5px solid #1e293b",
                    transition: "all 0.2s",
                  }}>
                    <div style={{ fontSize: 13, fontWeight: 600, color: strategy?.mode === key ? m.color : "#94a3b8", marginBottom: 4 }}>{m.label}</div>
                    <div style={{ fontSize: 11, color: "#475569" }}>{m.desc}</div>
                  </button>
                ))}
              </div>

              {/* Weight Bars */}
              {strategy?.weights && (
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px 24px" }}>
                  <WeightBar label="Profit" value={strategy.weights.profit} color="#dc2626" />
                  <WeightBar label="Revenue" value={strategy.weights.revenue} color="#f59e0b" />
                  <WeightBar label="Growth" value={strategy.weights.growth} color="#16a34a" />
                  <WeightBar label="Capture" value={strategy.weights.capture} color="#8b5cf6" />
                </div>
              )}
            </div>

            {/* Manual Trigger */}
            <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
              <button onClick={runCycle} style={{
                padding: "10px 20px", borderRadius: 6, fontSize: 12, fontWeight: 600,
                background: "linear-gradient(135deg, #1e3a5f, #0f2440)", color: "#38bdf8",
                border: "1px solid #1e3a5f",
              }}>
                ▶ RUN CONTROL CYCLE
              </button>
              <button onClick={fetchAll} style={{
                padding: "10px 16px", borderRadius: 6, fontSize: 12, fontWeight: 500,
                background: "#0f172a", color: "#64748b", border: "1px solid #1e293b",
              }}>
                ↻ REFRESH
              </button>
              {runResult && (
                <span style={{ fontSize: 11, color: "#4ade80", fontFamily: "JetBrains Mono, monospace" }}>
                  eval={runResult.evaluated} proposed={runResult.changes_proposed} applied={runResult.changes_applied} blocked={runResult.changes_blocked}
                </span>
              )}
            </div>
          </div>
        )}

        {/* DECISIONS TAB */}
        {activeTab === "decisions" && (
          <div>
            <div style={{ display: "flex", gap: 12, marginBottom: 16 }}>
              <StatCard label="Total Decisions" value={stats.total || 0} />
              <StatCard label="Changes" value={stats.changes || 0} accent="#38bdf8" />
              <StatCard label="Holds" value={stats.holds || 0} />
              <StatCard label="Avg Improvement" value={`${((stats.avg_score_improvement || 0) * 100).toFixed(2)}%`} accent="#4ade80" />
            </div>

            <div style={{ background: "#0f172a", border: "1px solid #1e293b", borderRadius: 10, overflow: "hidden" }}>
              <div style={{ padding: "12px 16px", borderBottom: "1px solid #1e293b", fontSize: 11, color: "#475569", textTransform: "uppercase", letterSpacing: 1.5 }}>
                RECENT DECISIONS ({recentDecisions.length})
              </div>
              <div style={{ maxHeight: 440, overflowY: "auto" }}>
                <table>
                  <thead>
                    <tr>
                      <th>Service</th><th style={{ textAlign: "center" }}>Action</th>
                      <th style={{ textAlign: "right" }}>Current</th><th style={{ textAlign: "right" }}>Target</th>
                      <th style={{ textAlign: "right" }}>Conf</th><th>Reason</th><th>Time</th>
                    </tr>
                  </thead>
                  <tbody>
                    {recentDecisions.length === 0 ? (
                      <tr><td colSpan={7} style={{ padding: 20, textAlign: "center", color: "#334155", fontSize: 12 }}>No decisions yet. Run a control cycle.</td></tr>
                    ) : recentDecisions.map((d, i) => <DecisionRow key={i} d={d} />)}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* SIMULATE TAB */}
        {activeTab === "simulate" && (
          <div>
            <div style={{ background: "#0f172a", border: "1px solid #1e293b", borderRadius: 10, padding: 20, marginBottom: 20 }}>
              <div style={{ fontSize: 11, color: "#475569", textTransform: "uppercase", letterSpacing: 1.5, marginBottom: 12 }}>SERVICE SIMULATION</div>
              <div style={{ display: "flex", gap: 10 }}>
                <input
                  value={simService}
                  onChange={e => setSimService(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && simulate()}
                  placeholder="Enter exact service name..."
                  style={{
                    flex: 1, padding: "10px 14px", borderRadius: 6, fontSize: 13,
                    background: "#0a0a1a", color: "#e2e8f0", border: "1px solid #1e293b",
                    fontFamily: "JetBrains Mono, monospace", outline: "none",
                  }}
                />
                <button onClick={simulate} disabled={simLoading} style={{
                  padding: "10px 20px", borderRadius: 6, fontSize: 12, fontWeight: 600,
                  background: "#1e3a5f", color: "#38bdf8", border: "1px solid #1e3a5f",
                  opacity: simLoading ? 0.5 : 1,
                }}>
                  {simLoading ? "..." : "SIMULATE"}
                </button>
              </div>
            </div>

            {simResult && !simResult.error && (
              <div style={{ background: "#0f172a", border: "1px solid #1e293b", borderRadius: 10, padding: 20 }}>
                <div style={{ display: "flex", gap: 12, marginBottom: 16, flexWrap: "wrap" }}>
                  <StatCard label="Service" value={simResult.service} />
                  <StatCard label="Decision" value={simResult.decision?.action?.toUpperCase()} accent={simResult.decision?.action === "change" ? "#38bdf8" : "#64748b"} />
                  <StatCard label="Best Price" value={`₹${simResult.decision?.best_price?.toFixed(1)}`} accent="#4ade80" />
                  <StatCard label="Confidence" value={`${((simResult.decision?.confidence || 0) * 100).toFixed(0)}%`} accent="#8b5cf6" />
                  <StatCard label="Elasticity" value={simResult.elasticity?.toFixed(2)} />
                </div>

                <div style={{ fontSize: 11, color: "#475569", textTransform: "uppercase", letterSpacing: 1.5, marginBottom: 10 }}>CANDIDATE SCORES</div>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {(simResult.decision?.all_scores || []).map((s, i) => (
                    <div key={i} style={{
                      padding: "8px 12px", borderRadius: 6, fontSize: 12,
                      fontFamily: "JetBrains Mono, monospace",
                      background: s.price === simResult.decision?.best_price ? "#1e3a5f" : "#0a0a1a",
                      border: s.price === simResult.decision?.best_price ? "1px solid #38bdf8" : "1px solid #1e293b",
                      color: s.price === simResult.decision?.best_price ? "#38bdf8" : "#94a3b8",
                    }}>
                      ₹{s.price?.toFixed(1)} → {s.score?.toFixed(4)}
                    </div>
                  ))}
                </div>

                {simResult.decision?.expected && (
                  <div style={{ marginTop: 16, display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10 }}>
                    <StatCard label="Proj. Demand" value={simResult.decision.expected.demand?.toFixed(2)} />
                    <StatCard label="Proj. Revenue" value={`₹${simResult.decision.expected.revenue?.toFixed(1)}`} />
                    <StatCard label="Proj. Profit" value={`₹${simResult.decision.expected.profit?.toFixed(1)}`} accent="#4ade80" />
                    <StatCard label="Growth" value={`${((simResult.decision.expected.growth || 0) * 100).toFixed(1)}%`} accent={simResult.decision.expected.growth > 0 ? "#4ade80" : "#f87171"} />
                  </div>
                )}

                <div style={{ marginTop: 12, fontSize: 11, color: "#475569" }}>
                  Reason: <span style={{ color: "#94a3b8" }}>{simResult.decision?.reason}</span>
                  {" · "}Mode: <span style={{ color: MODES[simResult.strategy_mode]?.color || "#94a3b8" }}>{simResult.strategy_mode}</span>
                </div>
              </div>
            )}

            {simResult?.error && (
              <div style={{ padding: 16, background: "#450a0a", border: "1px solid #7f1d1d", borderRadius: 8, fontSize: 13, color: "#fca5a5" }}>
                {simResult.error}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
