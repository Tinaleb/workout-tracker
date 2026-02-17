import { useState, useEffect, useRef, useCallback } from "react";

// ─── Constants ────────────────────────────────────────────────────────────────
const LS_CONSENT  = "wt_consent_v1";
const LS_SESSIONS = "wt_sessions_v1";
const LS_HISTORY  = "wt_history_v1";

// ─── Responsive hook ──────────────────────────────────────────────────────────
function useIsMobile() {
  const [isMobile, setIsMobile] = useState(() => window.innerWidth < 768);
  useEffect(() => {
    const handler = () => setIsMobile(window.innerWidth < 768);
    window.addEventListener("resize", handler);
    return () => window.removeEventListener("resize", handler);
  }, []);
  return isMobile;
}

function useIsTablet() {
  const [isTablet, setIsTablet] = useState(() => window.innerWidth >= 768 && window.innerWidth < 1024);
  useEffect(() => {
    const handler = () => setIsTablet(window.innerWidth >= 768 && window.innerWidth < 1024);
    window.addEventListener("resize", handler);
    return () => window.removeEventListener("resize", handler);
  }, []);
  return isTablet;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
const pad   = (n) => String(n).padStart(2, "0");
const fmt24 = (d) => d ? `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}` : "";

const calcDuration = (start, end) => {
  if (!start || !end) return null;
  const diff = Math.round((end - start) / 1000);
  if (diff <= 0) return null;
  const h = Math.floor(diff / 3600), m = Math.floor((diff % 3600) / 60), s = diff % 60;
  return h > 0 ? `${h}h ${pad(m)}m ${pad(s)}s` : `${pad(m)}m ${pad(s)}s`;
};

const newExercise = () => ({
  id: crypto.randomUUID(), type: "", sets: "", reps: "",
  weight: "", weightUnit: "lbs", duration: "", notes: "", isPR: false,
});

const newSession = () => ({
  id: crypto.randomUUID(),
  date: new Date().toISOString().split("T")[0],
  name: "", startTime: null, endTime: null,
  exercises: [newExercise()],
});

const serialize = (sessions) =>
  JSON.stringify(sessions, (_, v) => (v instanceof Date ? { __date: v.toISOString() } : v));

const deserialize = (str) => {
  if (!str) return null;
  try {
    return JSON.parse(str, (_, v) =>
      v && typeof v === "object" && v.__date ? new Date(v.__date) : v);
  } catch { return null; }
};

const lsGet = (key) => { try { return localStorage.getItem(key); } catch { return null; } };
const lsSet = (key, val) => { try { localStorage.setItem(key, val); } catch {} };

// ─── SheetJS loader ───────────────────────────────────────────────────────────
function parseUploadedFile(file, onDone) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const run = () => {
      try {
        const wb = window.XLSX.read(e.target.result, { type: "array" });
        onDone(window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" }), null);
      } catch { onDone(null, "Could not read file. Make sure it matches the exported format."); }
    };
    if (!window.XLSX) {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      s.onload = run; s.onerror = () => onDone(null, "Could not load spreadsheet parser.");
      document.head.appendChild(s);
    } else { run(); }
  };
  reader.readAsArrayBuffer(file);
}

// ─── Export CSV ───────────────────────────────────────────────────────────────
function exportCSV(sessions) {
  const hdrs = ["Session","Date","Session Name","Start Time","End Time","Total Duration","Exercise","Sets","Reps","Weight","Unit","Exercise Duration","Notes","PR"];
  const esc  = (v) => { const s = String(v ?? ""); return (s.includes(",") || s.includes('"')) ? `"${s.replace(/"/g, '""')}"` : s; };
  const rows = [hdrs];
  sessions.forEach((s, si) => {
    const dur = calcDuration(s.startTime, s.endTime);
    s.exercises.forEach((e) =>
      rows.push([`Session ${si+1}`, s.date, s.name||"", s.startTime?fmt24(s.startTime):"", s.endTime?fmt24(s.endTime):"", dur||"", e.type, e.sets, e.reps, e.weight, e.weight?e.weightUnit:"", e.duration, e.notes, e.isPR?"PR":""])
    );
  });
  const blob = new Blob([rows.map((r) => r.map(esc).join(",")).join("\n")], { type: "text/csv;charset=utf-8;" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a"); a.href = url; a.download = `workout-log-${new Date().toISOString().split("T")[0]}.csv`; a.click();
  URL.revokeObjectURL(url);
}

// ─── Export PDF ───────────────────────────────────────────────────────────────
function exportPDF(sessions) {
  const exHtml = sessions.map((s, si) => {
    const dur = calcDuration(s.startTime, s.endTime);
    const exRows = s.exercises.map((e) =>
      `<tr${e.isPR?' class="pr"':''}><td>${e.isPR?"🏆 ":""}${e.type||"—"}</td><td>${e.sets||"—"}</td><td>${e.reps||"—"}</td><td>${e.weight?`${e.weight} ${e.weightUnit}`:"—"}</td><td>${e.duration||"—"}</td><td>${e.notes||"—"}</td></tr>`
    ).join("");
    return `<div class="s"><div class="sh"><span class="sn">SESSION ${si+1}</span><span class="st">${s.name||"Untitled"}</span><span class="sd">${s.date}</span></div>${s.startTime||s.endTime?`<div class="tr">${s.startTime?`<span>Start: <b>${fmt24(s.startTime)}</b></span>`:""} ${s.endTime?`<span>End: <b>${fmt24(s.endTime)}</b></span>`:""} ${dur?`<span>Duration: <b>${dur}</b></span>`:""}</div>`:""}<table><thead><tr><th>Exercise</th><th>Sets</th><th>Reps</th><th>Weight</th><th>Duration</th><th>Notes</th></tr></thead><tbody>${exRows}</tbody></table></div>`;
  }).join("");
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width"/><title>Workout Log</title><style>@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600&family=DM+Serif+Display&display=swap');*{margin:0;padding:0;box-sizing:border-box}body{font-family:'DM Sans',sans-serif;font-size:11px;color:#1a1a2e;padding:24px}h1{font-family:'DM Serif Display',serif;font-size:24px;margin-bottom:4px}.sub{color:#666;margin-bottom:20px;font-size:10px}.s{margin-bottom:22px;page-break-inside:avoid}.sh{display:flex;align-items:baseline;gap:8px;margin-bottom:6px;padding-bottom:5px;border-bottom:2px solid #1a1a2e;flex-wrap:wrap}.sn{font-size:9px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#888;padding:2px 5px;background:#f0f0ec;border-radius:3px}.st{font-family:'DM Serif Display',serif;font-size:14px;flex:1}.sd{font-size:10px;color:#555}.tr{display:flex;gap:12px;margin-bottom:6px;font-size:10px;color:#555;flex-wrap:wrap}table{width:100%;border-collapse:collapse}th{background:#f0f0f0;text-align:left;padding:4px 6px;font-size:9px;text-transform:uppercase;letter-spacing:.06em;color:#555}td{padding:4px 6px;border-bottom:1px solid #eee;vertical-align:top;font-size:10px}.pr td{background:#fffbeb}.footer{margin-top:24px;font-size:9px;color:#aaa;border-top:1px solid #eee;padding-top:6px}</style></head><body><h1>Workout Log</h1><p class="sub">Exported ${new Date().toLocaleDateString("en-US",{weekday:"long",year:"numeric",month:"long",day:"numeric"})} · ${sessions.length} session${sessions.length!==1?"s":""}</p>${exHtml}<div class="footer">🏆 = Personal Record · Generated by Workout Tracker · All data stored locally on your device</div></body></html>`;
  const win = window.open("", "_blank");
  if (!win) { alert("Please allow pop-ups to export PDF."); return; }
  win.document.write(html); win.document.close();
  win.onload = () => { win.focus(); win.print(); };
}

// ─── Consent Modal ────────────────────────────────────────────────────────────
function ConsentModal({ onAccept, isMobile }) {
  return (
    <div style={R.overlay}>
      <div style={{ ...R.modal, ...(isMobile ? R.modalMobile : {}) }}>
        <div style={R.modalIcon}>💪</div>
        <h2 style={R.modalTitle}>Welcome to Workout Tracker</h2>
        <p style={R.modalSubtitle}>Before you begin, please review how this app works.</p>
        <div style={R.consentBox}>
          {[
            { icon: "✓", color: "#2d6a4f", text: <><strong>Your data stays on your device.</strong> All workout data is saved locally in your browser only. Nothing is sent to any server, collected, or shared.</> },
            { icon: "✓", color: "#2d6a4f", text: <><strong>No account required.</strong> Your progress is automatically saved in your browser and will be there when you return.</> },
            { icon: "✓", color: "#2d6a4f", text: <><strong>You are in control.</strong> Export your data at any time as PDF or CSV, and clear it from browser settings whenever you choose.</> },
            { icon: "⚠", color: "#d97706", text: <><strong>Clearing your browser data</strong> will erase your saved workout history. Export regularly to keep a permanent record.</> },
          ].map((item, i) => (
            <div key={i} style={R.consentItem}>
              <span style={{ ...R.consentIcon, color: item.color }}>{item.icon}</span>
              <div style={R.consentText}>{item.text}</div>
            </div>
          ))}
        </div>
        <p style={R.consentLegal}>By tapping <strong>I Understand — Let's Go</strong> you acknowledge that you have read how this app handles your data. No personal information is collected, stored, or transmitted.</p>
        <button style={R.acceptBtn} onClick={onAccept}>I Understand — Let's Go 🚀</button>
        <p style={R.modalNote}>You will only see this message once.</p>
      </div>
    </div>
  );
}

// ─── Privacy Badge ────────────────────────────────────────────────────────────
function PrivacyBadge() {
  const [open, setOpen] = useState(false);
  return (
    <div style={{ position: "relative" }}>
      <button style={R.privacyBadge} onClick={() => setOpen(!open)}>🔒 Local Only</button>
      {open && (
        <div style={R.privacyPopover}>
          <strong>Your privacy is protected.</strong><br />
          All data is stored in your browser only. No accounts, no servers, no tracking. Export regularly for a permanent backup.
          <button style={R.privacyClose} onClick={() => setOpen(false)}>✕</button>
        </div>
      )}
    </div>
  );
}

// ─── Save Indicator ───────────────────────────────────────────────────────────
function SaveIndicator({ status }) {
  const map = { saved: { text: "✓ Saved", color: "#2d6a4f" }, saving: { text: "Saving…", color: "#888" }, error: { text: "⚠ Save failed", color: "#ef4444" } };
  const info = map[status] || map.saved;
  return <span style={{ fontSize: 11, color: info.color, fontWeight: 600 }}>{info.text}</span>;
}

// ─── Rest Timer ───────────────────────────────────────────────────────────────
function RestTimer({ isMobile }) {
  const PRESETS = [30, 60, 90, 120];
  const [selected, setSelected]   = useState(60);
  const [remaining, setRemaining] = useState(null);
  const [running, setRunning]     = useState(false);
  const [done, setDone]           = useState(false);
  const itvRef = useRef(null);

  const playBeep = () => {
    try {
      const ctx = new (window.AudioContext || window.webkitAudioContext)();
      [0, 0.18, 0.36].forEach((delay) => {
        const osc = ctx.createOscillator(), gain = ctx.createGain();
        osc.connect(gain); gain.connect(ctx.destination);
        osc.frequency.value = 880;
        gain.gain.setValueAtTime(0.4, ctx.currentTime + delay);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + delay + 0.14);
        osc.start(ctx.currentTime + delay); osc.stop(ctx.currentTime + delay + 0.15);
      });
    } catch {}
  };

  const start = () => { setRemaining(selected); setRunning(true); setDone(false); };
  const stop  = () => { clearInterval(itvRef.current); setRunning(false); setRemaining(null); setDone(false); };

  useEffect(() => {
    if (!running) return;
    itvRef.current = setInterval(() => {
      setRemaining((r) => {
        if (r <= 1) { clearInterval(itvRef.current); setRunning(false); setDone(true); playBeep(); return 0; }
        return r - 1;
      });
    }, 1000);
    return () => clearInterval(itvRef.current);
  }, [running]);

  const pct    = remaining !== null ? (remaining / selected) * 100 : 100;
  const radius = isMobile ? 30 : 22;
  const circ   = 2 * Math.PI * radius;
  const svgSize = isMobile ? 80 : 64;

  return (
    <div style={{ ...R.timerCard, ...(done ? R.timerCardDone : {}), ...(isMobile ? R.timerCardMobile : {}) }}>
      <div style={R.timerLabel}>REST TIMER</div>
      <div style={R.timerPresets}>
        {PRESETS.map((p) => (
          <button key={p}
            style={{ ...R.presetBtn, ...(selected === p && !running ? R.presetActive : {}), ...(isMobile ? R.presetBtnMobile : {}) }}
            onClick={() => { if (!running) { setSelected(p); setDone(false); } }} disabled={running}>
            {p}s
          </button>
        ))}
      </div>
      <div style={{ ...R.timerRing, width: svgSize, height: svgSize }}>
        <svg width={svgSize} height={svgSize}>
          <circle cx={svgSize/2} cy={svgSize/2} r={radius} fill="none" stroke="#e8e8e4" strokeWidth={isMobile ? 5 : 4} />
          <circle cx={svgSize/2} cy={svgSize/2} r={radius} fill="none"
            stroke={done ? "#f59e0b" : running ? "#2d6a4f" : "#1a1a2e"}
            strokeWidth={isMobile ? 5 : 4} strokeDasharray={circ}
            strokeDashoffset={circ - (pct / 100) * circ}
            strokeLinecap="round"
            style={{ transform: `rotate(-90deg)`, transformOrigin: `${svgSize/2}px ${svgSize/2}px`, transition: "stroke-dashoffset 1s linear" }} />
        </svg>
        <div style={{ ...R.timerNum, fontSize: isMobile ? 20 : 16 }}>
          {done ? "✓" : remaining !== null ? `${remaining}` : `${selected}`}
        </div>
      </div>
      <div style={R.timerBtns}>
        {!running
          ? <button style={{ ...R.timerStartBtn, ...(isMobile ? R.timerBtnMobile : {}) }} onClick={start}>{done ? "Again" : "Start"}</button>
          : <button style={{ ...R.timerStopBtn, ...(isMobile ? R.timerBtnMobile : {}) }} onClick={stop}>Stop</button>}
      </div>
      {done && <div style={R.timerDoneMsg}>Rest up — time to go 💪</div>}
    </div>
  );
}

// ─── Mobile Timer Drawer ──────────────────────────────────────────────────────
function MobileTimerDrawer() {
  const [open, setOpen] = useState(false);
  return (
    <>
      <button style={R.timerFab} onClick={() => setOpen(!open)} aria-label="Rest Timer">
        ⏱
      </button>
      {open && (
        <div style={R.timerDrawer}>
          <div style={R.timerDrawerHeader}>
            <span style={R.timerDrawerTitle}>Rest Timer</span>
            <button style={R.timerDrawerClose} onClick={() => setOpen(false)}>✕</button>
          </div>
          <RestTimer isMobile={true} />
        </div>
      )}
    </>
  );
}

// ─── Summary Banner ───────────────────────────────────────────────────────────
function SummaryBanner({ sessions, isMobile }) {
  let totalSets = 0, totalReps = 0, totalVol = 0, prCount = 0;
  const types = new Set();
  sessions.forEach((s) => s.exercises.forEach((e) => {
    if (e.type) types.add(e.type.toLowerCase());
    const sets = parseInt(e.sets)||0, reps = parseInt(e.reps)||0, w = parseFloat(e.weight)||0;
    totalSets += sets; totalReps += sets * reps; totalVol += sets * reps * w;
    if (e.isPR) prCount++;
  }));
  const stats = [
    { label: "Sessions",   val: sessions.length,                                   icon: "📋" },
    { label: "Exercises",  val: types.size,                                         icon: "🏋️" },
    { label: "Sets",       val: totalSets,                                           icon: "🔢" },
    { label: "Reps",       val: totalReps,                                           icon: "🔄" },
    { label: "Volume",     val: totalVol > 0 ? totalVol.toLocaleString()+"lbs" : "—", icon: "⚖️" },
    { label: "PRs",        val: prCount,                                             icon: "🏆" },
  ];
  return (
    <div style={R.banner}>
      <div style={R.bannerTitle}>Today's Summary</div>
      <div style={{ ...R.bannerGrid, ...(isMobile ? R.bannerGridMobile : {}) }}>
        {stats.map((st) => (
          <div key={st.label} style={{ ...R.bannerStat, ...(isMobile ? R.bannerStatMobile : {}) }}>
            <div style={R.bannerIcon}>{st.icon}</div>
            <div style={{ ...R.bannerVal, ...(isMobile ? R.bannerValMobile : {}) }}>{st.val}</div>
            <div style={R.bannerStatLabel}>{st.label}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── History Panel ────────────────────────────────────────────────────────────
function HistoryPanel({ historyData, currentSessions, isMobile }) {
  const [tab, setTab] = useState("compare");

  const histMap = {};
  historyData.forEach((row) => {
    const key = (row["Exercise"] || row["exercise"] || "").trim().toLowerCase();
    if (!key) return;
    const w = parseFloat(row["Weight"] || row["weight"]) || 0;
    if (!histMap[key] || w > histMap[key].weight)
      histMap[key] = { weight: w, sets: parseInt(row["Sets"]||row["sets"])||0, reps: parseInt(row["Reps"]||row["reps"])||0, unit: row["Unit"]||row["unit"]||"lbs", date: row["Date"]||row["date"]||"" };
  });

  const currMap = {};
  currentSessions.forEach((s) => s.exercises.forEach((e) => {
    const key = e.type.trim().toLowerCase(); if (!key) return;
    const w = parseFloat(e.weight) || 0;
    if (!currMap[key] || w > currMap[key].weight)
      currMap[key] = { weight: w, sets: parseInt(e.sets)||0, reps: parseInt(e.reps)||0, unit: e.weightUnit, isPR: e.isPR };
  }));

  const allKeys = Array.from(new Set([...Object.keys(histMap), ...Object.keys(currMap)]));

  const byDate = {};
  historyData.forEach((row) => {
    const date = row["Date"]||row["date"]||"Unknown";
    const sname = row["Session Name"]||row["session name"]||"";
    const key = `${date}|${sname}`;
    if (!byDate[key]) byDate[key] = { date, sname, exercises: [] };
    if (row["Exercise"]||row["exercise"]) byDate[key].exercises.push(row);
  });
  const pastSessions = Object.values(byDate).sort((a, b) => b.date.localeCompare(a.date));

  return (
    <div style={R.histPanel}>
      <div style={R.histHeader}>
        <span style={R.histTitle}>📂 History &amp; Comparison</span>
        <div style={R.histTabs}>
          {["compare", "sessions"].map((t) => (
            <button key={t} style={{ ...R.histTab, ...(tab === t ? R.histTabActive : {}), ...(isMobile ? R.histTabMobile : {}) }}
              onClick={() => setTab(t)}>
              {t === "compare" ? "Compare" : "Past Sessions"}
            </button>
          ))}
        </div>
      </div>

      {tab === "compare" && (
        <div style={R.compareWrap}>
          {allKeys.length === 0 ? <p style={R.emptyMsg}>No exercises to compare yet.</p> : (
            isMobile ? (
              // Mobile: card layout instead of table
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {allKeys.map((key) => {
                  const h = histMap[key], c = currMap[key];
                  const hW = h?.weight||0, cW = c?.weight||0, diff = cW - hW;
                  const up = diff > 0, flat = diff === 0 && hW > 0 && cW > 0;
                  return (
                    <div key={key} style={R.compareCard}>
                      <div style={R.compareCardTitle}>
                        <span style={{ textTransform: "capitalize", fontWeight: 700 }}>{key}</span>
                        {c?.isPR && <span style={R.prChip}>🏆 PR</span>}
                        {h && c && hW > 0 && cW > 0 && (
                          <span style={{ marginLeft: "auto", color: up?"#2d6a4f":flat?"#888":"#ef4444", fontWeight: 700, fontSize: 13 }}>
                            {up?"▲":flat?"—":"▼"} {Math.abs(diff).toFixed(1)} {c.unit}
                          </span>
                        )}
                      </div>
                      <div style={R.compareCardRow}>
                        <div style={R.compareCardCol}>
                          <div style={R.compareCardColLabel}>Previous Best</div>
                          <div style={R.compareCardColVal}>{h ? (hW > 0 ? `${hW} ${h.unit}` : "—") : "No history"}</div>
                          {h && <div style={R.subVal}>{h.sets}×{h.reps}{h.date ? ` · ${h.date}` : ""}</div>}
                        </div>
                        <div style={R.compareCardDivider} />
                        <div style={R.compareCardCol}>
                          <div style={R.compareCardColLabel}>Today</div>
                          <div style={R.compareCardColVal}>{c ? (cW > 0 ? `${cW} ${c.unit}` : "—") : "Not logged"}</div>
                          {c && <div style={R.subVal}>{c.sets}×{c.reps}</div>}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              <table style={R.compTable}>
                <thead>
                  <tr>
                    {["Exercise","Previous Best","Today","Change"].map((h) => <th key={h} style={R.compTh}>{h}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {allKeys.map((key) => {
                    const h = histMap[key], c = currMap[key];
                    const hW = h?.weight||0, cW = c?.weight||0, diff = cW - hW;
                    const up = diff > 0, flat = diff === 0 && hW > 0 && cW > 0;
                    return (
                      <tr key={key} style={R.compRow}>
                        <td style={R.compTd}><span style={R.exName}>{key}</span>{c?.isPR && <span style={R.prChip}>🏆 PR</span>}</td>
                        <td style={R.compTd}>{h ? <><span>{hW > 0 ? `${hW} ${h.unit}` : "—"}</span><br/><span style={R.subVal}>{h.sets}×{h.reps}{h.date?` · ${h.date}`:""}</span></> : <span style={R.noData}>No history</span>}</td>
                        <td style={R.compTd}>{c ? <><span>{cW > 0 ? `${cW} ${c.unit}` : "—"}</span><br/><span style={R.subVal}>{c.sets}×{c.reps}</span></> : <span style={R.noData}>Not logged</span>}</td>
                        <td style={R.compTd}>{h && c && hW > 0 && cW > 0 ? <span style={{ color: up?"#2d6a4f":flat?"#888":"#ef4444", fontWeight: 600 }}>{up?"▲ ":flat?"— ":"▼ "}{Math.abs(diff).toFixed(1)} {c.unit}</span> : <span style={R.noData}>—</span>}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )
          )}
        </div>
      )}

      {tab === "sessions" && (
        <div style={R.pastWrap}>
          {pastSessions.length === 0 ? <p style={R.emptyMsg}>No past sessions found.</p> :
            pastSessions.map((ps, i) => (
              <div key={i} style={R.pastSession}>
                <div style={R.pastSessionHdr}>
                  <span style={R.pastDate}>{ps.date}</span>
                  {ps.sname && <span style={R.pastName}>{ps.sname}</span>}
                </div>
                <div style={R.pastExList}>
                  {ps.exercises.map((e, j) => (
                    <div key={j} style={R.pastEx}>
                      <span style={R.pastExName}>{e["Exercise"]||e["exercise"]}</span>
                      <span style={R.pastExDets}>{e["Sets"]||e["sets"]}×{e["Reps"]||e["reps"]}{(parseFloat(e["Weight"]||e["weight"])||0)>0?` · ${e["Weight"]||e["weight"]} ${e["Unit"]||e["unit"]||"lbs"}`:""}{(e["PR"]||e["pr"])==="PR"?" 🏆":""}</span>
                    </div>
                  ))}
                </div>
              </div>
            ))
          }
        </div>
      )}
    </div>
  );
}

// ─── Exercise Row ─────────────────────────────────────────────────────────────
function ExerciseRow({ ex, onChange, onRemove, canRemove, histMap, isMobile }) {
  const key  = ex.type.trim().toLowerCase();
  const prev = histMap?.[key];
  return (
    <div style={{ ...R.exRow, ...(ex.isPR ? R.exRowPR : {}) }}>
      {prev && ex.type && (
        <div style={R.prevHint}>
          📈 Previous best: <strong>{prev.weight > 0 ? `${prev.weight} ${prev.unit}` : "—"}</strong> · {prev.sets}×{prev.reps}{prev.date ? ` on ${prev.date}` : ""}
        </div>
      )}
      <div style={{ ...R.exGrid, ...(isMobile ? R.exGridMobile : {}) }}>
        <div style={{ ...R.field, gridColumn: "1 / -1" }}>
          <label style={R.label}>Exercise Type</label>
          <input style={{ ...R.input, ...(isMobile ? R.inputMobile : {}) }} placeholder="e.g. Bench Press, Squat, Run…" value={ex.type} onChange={(e) => onChange("type", e.target.value)} />
        </div>
        <div style={R.field}>
          <label style={R.label}>Sets</label>
          <input style={{ ...R.input, ...(isMobile ? R.inputMobile : {}) }} placeholder="e.g. 3" value={ex.sets} onChange={(e) => onChange("sets", e.target.value)} inputMode="numeric" />
        </div>
        <div style={R.field}>
          <label style={R.label}>Reps</label>
          <input style={{ ...R.input, ...(isMobile ? R.inputMobile : {}) }} placeholder="e.g. 12" value={ex.reps} onChange={(e) => onChange("reps", e.target.value)} inputMode="numeric" />
        </div>
        <div style={R.field}>
          <label style={R.label}>Weight</label>
          <div style={R.weightWrap}>
            <input style={{ ...R.input, ...(isMobile ? R.inputMobile : {}), flex: 1, minWidth: 0 }} placeholder="e.g. 135" value={ex.weight} onChange={(e) => onChange("weight", e.target.value)} inputMode="decimal" />
            <select style={{ ...R.unitSelect, ...(isMobile ? R.unitSelectMobile : {}) }} value={ex.weightUnit} onChange={(e) => onChange("weightUnit", e.target.value)}>
              <option>lbs</option><option>kg</option>
            </select>
          </div>
        </div>
        <div style={R.field}>
          <label style={R.label}>Duration (optional)</label>
          <input style={{ ...R.input, ...(isMobile ? R.inputMobile : {}) }} placeholder="e.g. 20 min" value={ex.duration} onChange={(e) => onChange("duration", e.target.value)} />
        </div>
        <div style={{ ...R.field, gridColumn: "1 / -1" }}>
          <label style={R.label}>Notes</label>
          <input style={{ ...R.input, ...(isMobile ? R.inputMobile : {}) }} placeholder="Any notes…" value={ex.notes} onChange={(e) => onChange("notes", e.target.value)} />
        </div>
      </div>
      <div style={R.exFooter}>
        <button style={{ ...R.prBtn, ...(ex.isPR ? R.prBtnActive : {}), ...(isMobile ? R.prBtnMobile : {}) }} onClick={() => onChange("isPR", !ex.isPR)}>
          {ex.isPR ? "🏆 PR!" : "☆ Mark PR"}
        </button>
        {canRemove && (
          <button style={{ ...R.removeExBtn, ...(isMobile ? R.removeExBtnMobile : {}) }} onClick={onRemove}>
            {isMobile ? "✕ Remove" : "✕ Remove"}
          </button>
        )}
      </div>
    </div>
  );
}

// ─── Session Card ─────────────────────────────────────────────────────────────
function SessionCard({ session, onUpdate, onRemove, canRemove, index, histMap, isMobile }) {
  const [collapsed, setCollapsed] = useState(false);
  const upd   = (f, v) => onUpdate({ ...session, [f]: v });
  const updEx = (id, f, v) => onUpdate({ ...session, exercises: session.exercises.map((e) => e.id === id ? { ...e, [f]: v } : e) });
  const addEx = () => onUpdate({ ...session, exercises: [...session.exercises, newExercise()] });
  const remEx = (id) => onUpdate({ ...session, exercises: session.exercises.filter((e) => e.id !== id) });
  const dur     = calcDuration(session.startTime, session.endTime);
  const prCount = session.exercises.filter((e) => e.isPR).length;

  return (
    <div style={R.card}>
      <div style={{ ...R.cardHeader, ...(isMobile ? R.cardHeaderMobile : {}) }}>
        <div style={R.cardHeaderLeft}>
          <span style={R.sessionBadge}>SESSION {index + 1}</span>
          {session.name && <span style={R.sessionNameBadge}>{session.name}</span>}
        </div>
        <div style={R.cardHeaderRight}>
          {prCount > 0 && <span style={R.prCountBadge}>🏆 {prCount}</span>}
          {dur && <span style={R.durBadge}>⏱ {dur}</span>}
          <button style={{ ...R.collapseBtn, ...(isMobile ? R.collapseBtnMobile : {}) }} onClick={() => setCollapsed(!collapsed)}>
            {collapsed ? "▶" : "▼"}
          </button>
          {canRemove && <button style={{ ...R.removeSessionBtn, ...(isMobile ? R.removeSessionBtnMobile : {}) }} onClick={onRemove}>Remove</button>}
        </div>
      </div>

      {!collapsed && (
        <>
          <div style={{ ...R.sessionMeta, ...(isMobile ? R.sessionMetaMobile : {}) }}>
            <div style={R.field}>
              <label style={R.label}>Session Name (optional)</label>
              <input style={{ ...R.input, ...(isMobile ? R.inputMobile : {}) }} placeholder="e.g. Leg Day, Morning Run…" value={session.name} onChange={(e) => upd("name", e.target.value)} />
            </div>
            <div style={R.field}>
              <label style={R.label}>Date</label>
              <input style={{ ...R.input, ...(isMobile ? R.inputMobile : {}) }} type="date" value={session.date} onChange={(e) => upd("date", e.target.value)} />
            </div>
            <div style={R.field}>
              <label style={R.label}>Start Time</label>
              <div style={{ ...R.timeWrap, ...(isMobile ? R.timeWrapMobile : {}) }}>
                <span style={{ ...R.timeDisplay, ...(isMobile ? R.timeDisplayMobile : {}) }}>{session.startTime ? fmt24(session.startTime) : "—"}</span>
                <button style={{ ...R.timeBtn, ...(isMobile ? R.timeBtnMobile : {}) }} onClick={() => upd("startTime", new Date())}>Mark Start</button>
                {session.startTime && <button style={{ ...R.timeBtn, ...R.timeBtnClear, ...(isMobile ? R.timeBtnMobile : {}) }} onClick={() => upd("startTime", null)}>Clear</button>}
              </div>
            </div>
            <div style={R.field}>
              <label style={R.label}>End Time</label>
              <div style={{ ...R.timeWrap, ...(isMobile ? R.timeWrapMobile : {}) }}>
                <span style={{ ...R.timeDisplay, ...(isMobile ? R.timeDisplayMobile : {}) }}>{session.endTime ? fmt24(session.endTime) : "—"}</span>
                <button style={{ ...R.timeBtn, ...(isMobile ? R.timeBtnMobile : {}) }} onClick={() => upd("endTime", new Date())}>Mark End</button>
                {session.endTime && <button style={{ ...R.timeBtn, ...R.timeBtnClear, ...(isMobile ? R.timeBtnMobile : {}) }} onClick={() => upd("endTime", null)}>Clear</button>}
              </div>
            </div>
          </div>
          <div style={R.exerciseList}>
            <div style={R.exListHeader}><span style={R.exListTitle}>Exercises</span></div>
            {session.exercises.map((ex) => (
              <ExerciseRow key={ex.id} ex={ex}
                onChange={(f, v) => updEx(ex.id, f, v)}
                onRemove={() => remEx(ex.id)}
                canRemove={session.exercises.length > 1}
                histMap={histMap}
                isMobile={isMobile} />
            ))}
            <button style={{ ...R.addExBtn, ...(isMobile ? R.addExBtnMobile : {}) }} onClick={addEx}>+ Add Exercise</button>
          </div>
        </>
      )}
    </div>
  );
}

// ─── Mobile Header Menu ───────────────────────────────────────────────────────
function MobileMenu({ onUpload, onExportPDF, onExportCSV, historyData, showHistory, setShowHistory, fileRef }) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{ position: "relative" }}>
      <button style={R.menuBtn} onClick={() => setOpen(!open)}>⋯ Menu</button>
      {open && (
        <div style={R.menuDropdown}>
          <button style={R.menuItem} onClick={() => { fileRef.current.click(); setOpen(false); }}>📂 Upload History</button>
          {historyData && (
            <button style={R.menuItem} onClick={() => { setShowHistory(!showHistory); setOpen(false); }}>
              {showHistory ? "Hide" : "Show"} Comparison
            </button>
          )}
          <div style={R.menuDivider} />
          <button style={R.menuItem} onClick={() => { onExportPDF(); setOpen(false); }}>⬇ Export PDF</button>
          <button style={R.menuItem} onClick={() => { onExportCSV(); setOpen(false); }}>⬇ Export Excel / CSV</button>
          <div style={R.menuDivider} />
          <button style={{ ...R.menuItem, fontSize: 11, color: "#888" }} onClick={() => setOpen(false)}>✕ Close</button>
        </div>
      )}
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function WorkoutTracker() {
  const isMobile = useIsMobile();
  const isTablet = useIsTablet();

  const [sessions,     setSessions]     = useState(null);
  const [historyData,  setHistoryData]  = useState(null);
  const [historyError, setHistoryError] = useState("");
  const [showHistory,  setShowHistory]  = useState(false);
  const [saveStatus,   setSaveStatus]   = useState("saved");
  const [hasConsent,   setHasConsent]   = useState(null);
  const [isDirty,      setIsDirty]      = useState(false);

  const fileRef     = useRef(null);
  const saveTimerRef = useRef(null);

  // Boot
  useEffect(() => {
    setHasConsent(lsGet(LS_CONSENT) === "true");
    const saved = deserialize(lsGet(LS_SESSIONS));
    setSessions(saved && saved.length ? saved : [newSession()]);
    const savedHistory = lsGet(LS_HISTORY);
    if (savedHistory) { try { setHistoryData(JSON.parse(savedHistory)); } catch {} }
  }, []);

  // beforeunload guard
  useEffect(() => {
    const handler = (e) => {
      if (!isDirty) return;
      e.preventDefault();
      e.returnValue = "You have workout data that hasn't been exported. Leave anyway?";
      return e.returnValue;
    };
    window.addEventListener("beforeunload", handler);
    return () => window.removeEventListener("beforeunload", handler);
  }, [isDirty]);

  // Auto-save
  useEffect(() => {
    if (!sessions) return;
    setSaveStatus("saving");
    clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => {
      try { lsSet(LS_SESSIONS, serialize(sessions)); setSaveStatus("saved"); setIsDirty(true); }
      catch { setSaveStatus("error"); }
    }, 800);
    return () => clearTimeout(saveTimerRef.current);
  }, [sessions]);

  const updateSession = useCallback((updated) => setSessions((p) => p.map((s) => s.id === updated.id ? updated : s)), []);
  const addSession    = () => setSessions((p) => [...p, newSession()]);
  const removeSession = (id) => setSessions((p) => p.filter((s) => s.id !== id));

  const handleUpload = (e) => {
    const file = e.target.files[0]; if (!file) return;
    setHistoryError("");
    parseUploadedFile(file, (rows, err) => {
      if (err) { setHistoryError(err); return; }
      setHistoryData(rows); lsSet(LS_HISTORY, JSON.stringify(rows)); setShowHistory(true);
    });
    e.target.value = "";
  };

  const handleExportCSV = () => { exportCSV(sessions); setIsDirty(false); };
  const handleExportPDF = () => { exportPDF(sessions); setIsDirty(false); };
  const handleAccept    = () => { lsSet(LS_CONSENT, "true"); setHasConsent(true); };

  const histMap = {};
  if (historyData) {
    historyData.forEach((row) => {
      const key = (row["Exercise"]||row["exercise"]||"").trim().toLowerCase();
      if (!key) return;
      const w = parseFloat(row["Weight"]||row["weight"]) || 0;
      if (!histMap[key] || w > histMap[key].weight)
        histMap[key] = { weight: w, sets: parseInt(row["Sets"]||row["sets"])||0, reps: parseInt(row["Reps"]||row["reps"])||0, unit: row["Unit"]||row["unit"]||"lbs", date: row["Date"]||row["date"]||"" };
    });
  }

  if (hasConsent === null || sessions === null) {
    return <div style={R.loading}><div style={{ fontSize: 48 }}>💪</div><p style={{ color: "#888", marginTop: 12 }}>Loading…</p></div>;
  }

  if (!hasConsent) return <ConsentModal onAccept={handleAccept} isMobile={isMobile} />;

  const totalEx = sessions.reduce((a, s) => a + s.exercises.length, 0);

  return (
    <div style={R.app}>
      {/* Header */}
      <header style={{ ...R.header, ...(isMobile ? R.headerMobile : {}) }}>
        <div style={{ ...R.headerInner, ...(isMobile ? R.headerInnerMobile : {}) }}>
          <div style={R.headerBrand}>
            <h1 style={{ ...R.title, ...(isMobile ? R.titleMobile : {}) }}>Workout Tracker</h1>
            <div style={R.headerMeta}>
              <span style={R.subtitle}>{sessions.length} session{sessions.length!==1?"s":""} · {totalEx} ex</span>
              <SaveIndicator status={saveStatus} />
            </div>
          </div>

          {isMobile ? (
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <PrivacyBadge />
              <MobileMenu onUpload={handleUpload} onExportPDF={handleExportPDF} onExportCSV={handleExportCSV}
                historyData={historyData} showHistory={showHistory} setShowHistory={setShowHistory} fileRef={fileRef} />
            </div>
          ) : (
            <div style={R.headerActions}>
              <PrivacyBadge />
              <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={handleUpload} />
              <button style={R.uploadBtn} onClick={() => fileRef.current.click()}>📂 Upload History</button>
              {historyData && (
                <button style={{ ...R.uploadBtn, background: showHistory?"#1a1a2e":"#f0f0ec", color: showHistory?"#fff":"#444" }}
                  onClick={() => setShowHistory(!showHistory)}>
                  {showHistory ? "Hide" : "Show"} Comparison
                </button>
              )}
              <button style={R.exportBtn} onClick={handleExportPDF}>⬇ PDF</button>
              <button style={{ ...R.exportBtn, ...R.exportBtnGreen }} onClick={handleExportCSV}>⬇ Excel / CSV</button>
            </div>
          )}
        </div>

        {/* Hidden file input for mobile menu */}
        {isMobile && <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={handleUpload} />}

        {historyError && <div style={R.errorBanner}>⚠ {historyError}</div>}
        {isDirty && (
          <div style={{ ...R.exportNudge, ...(isMobile ? R.exportNudgeMobile : {}) }}>
            <span>💡 Export when done to save a permanent copy.</span>
            <button style={R.exportNudgeBtn} onClick={handleExportCSV}>Export Now</button>
          </div>
        )}
      </header>

      <main style={{ ...R.main, ...(isMobile ? R.mainMobile : {}) }}>
        <SummaryBanner sessions={sessions} isMobile={isMobile} />

        {isMobile ? (
          // ── Mobile: single column ─────────────────────────────────────────
          <div style={R.sessionsCol}>
            {sessions.map((s, i) => (
              <SessionCard key={s.id} session={s} index={i}
                onUpdate={updateSession} onRemove={() => removeSession(s.id)}
                canRemove={sessions.length > 1} histMap={historyData ? histMap : null} isMobile={true} />
            ))}
            <button style={R.addSessionBtnMobile} onClick={addSession}>+ Add Session</button>
          </div>
        ) : (
          // ── Desktop / Tablet: two column ──────────────────────────────────
          <div style={{ ...R.layout, ...(isTablet ? R.layoutTablet : {}) }}>
            <div style={R.sessionsCol}>
              {sessions.map((s, i) => (
                <SessionCard key={s.id} session={s} index={i}
                  onUpdate={updateSession} onRemove={() => removeSession(s.id)}
                  canRemove={sessions.length > 1} histMap={historyData ? histMap : null} isMobile={false} />
              ))}
              <button style={R.addSessionBtn} onClick={addSession}>+ Add Session</button>
            </div>
            <div style={{ ...R.sidebar, ...(isTablet ? R.sidebarTablet : {}) }}>
              <RestTimer isMobile={false} />
              {!historyData && (
                <div style={R.sidebarHint}>
                  <div style={R.sidebarHintTitle}>Track Progress</div>
                  <p style={R.sidebarHintText}>Upload a previous workout export to compare your progress inline.</p>
                  <button style={R.sidebarUploadBtn} onClick={() => fileRef.current.click()}>📂 Upload History</button>
                </div>
              )}
            </div>
          </div>
        )}

        {historyData && showHistory && (
          <HistoryPanel historyData={historyData} currentSessions={sessions} isMobile={isMobile} />
        )}
      </main>

      {/* Mobile floating rest timer button */}
      {isMobile && <MobileTimerDrawer />}

      <footer style={{ ...R.footer, ...(isMobile ? R.footerMobile : {}) }}>
        🔒 All data stored locally · No accounts · No tracking · Export to keep a permanent copy
      </footer>
    </div>
  );
}

// ─── Responsive Styles ────────────────────────────────────────────────────────
const R = {
  app:   { fontFamily: "'DM Sans','Segoe UI',sans-serif", background: "#f7f7f5", minHeight: "100vh", color: "#1a1a2e" },
  loading: { display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100vh" },
  // Header
  header:        { background: "#fff", borderBottom: "1px solid #e8e8e4", position: "sticky", top: 0, zIndex: 100 },
  headerMobile:  { position: "sticky" },
  headerInner:   { maxWidth: 1080, margin: "0 auto", padding: "12px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" },
  headerInnerMobile: { padding: "12px 16px" },
  headerBrand:   { display: "flex", flexDirection: "column", gap: 2 },
  title:         { fontFamily: "'DM Serif Display',Georgia,serif", fontSize: 22, fontWeight: 400, letterSpacing: "-0.02em", margin: 0 },
  titleMobile:   { fontSize: 18 },
  headerMeta:    { display: "flex", alignItems: "center", gap: 10 },
  subtitle:      { fontSize: 11, color: "#888" },
  headerActions: { display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" },
  uploadBtn:     { padding: "7px 12px", background: "#f0f0ec", border: "1px solid #ddd", borderRadius: 8, fontSize: 11, fontWeight: 600, cursor: "pointer", color: "#444" },
  exportBtn:     { padding: "7px 12px", background: "#1a1a2e", color: "#fff", border: "none", borderRadius: 8, fontSize: 11, fontWeight: 600, cursor: "pointer" },
  exportBtnGreen:{ background: "#2d6a4f" },
  exportNudge:   { background: "#fffbeb", borderTop: "1px solid #fcd34d", fontSize: 11, padding: "8px 24px", color: "#92400e", display: "flex", alignItems: "center", gap: 10 },
  exportNudgeMobile: { padding: "8px 16px", fontSize: 12 },
  exportNudgeBtn:{ padding: "4px 10px", background: "#d97706", color: "#fff", border: "none", borderRadius: 6, fontSize: 11, cursor: "pointer", fontWeight: 600, marginLeft: "auto", whiteSpace: "nowrap" },
  errorBanner:   { background: "#fef2f2", color: "#ef4444", fontSize: 11, padding: "6px 24px", borderTop: "1px solid #fecaca" },
  // Mobile Menu
  menuBtn:       { padding: "8px 14px", background: "#1a1a2e", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" },
  menuDropdown:  { position: "absolute", right: 0, top: "calc(100% + 6px)", background: "#fff", border: "1px solid #e8e8e4", borderRadius: 12, boxShadow: "0 8px 32px rgba(0,0,0,.12)", zIndex: 200, minWidth: 200, overflow: "hidden" },
  menuItem:      { display: "block", width: "100%", padding: "13px 16px", background: "none", border: "none", textAlign: "left", fontSize: 14, cursor: "pointer", color: "#1a1a2e", fontWeight: 500 },
  menuDivider:   { height: 1, background: "#f0f0ec", margin: "4px 0" },
  // Privacy Badge
  privacyBadge:  { padding: "5px 10px", background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 20, fontSize: 10, fontWeight: 700, color: "#166534", cursor: "pointer", whiteSpace: "nowrap" },
  privacyPopover:{ position: "absolute", right: 0, top: "calc(100% + 6px)", background: "#fff", border: "1px solid #e8e8e4", borderRadius: 10, padding: "12px 14px 10px", width: 240, fontSize: 11, lineHeight: 1.5, color: "#333", boxShadow: "0 4px 20px rgba(0,0,0,.08)", zIndex: 200 },
  privacyClose:  { position: "absolute", top: 6, right: 8, background: "none", border: "none", cursor: "pointer", fontSize: 12, color: "#aaa" },
  // Consent Modal
  overlay:       { position: "fixed", inset: 0, background: "rgba(10,10,20,.65)", backdropFilter: "blur(4px)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16, overflowY: "auto" },
  modal:         { background: "#fff", borderRadius: 20, padding: "32px 28px 24px", maxWidth: 520, width: "100%", boxShadow: "0 20px 60px rgba(0,0,0,.2)", textAlign: "center" },
  modalMobile:   { borderRadius: 16, padding: "28px 20px 20px", margin: "auto" },
  modalIcon:     { fontSize: 44, marginBottom: 10 },
  modalTitle:    { fontFamily: "'DM Serif Display',Georgia,serif", fontSize: 24, fontWeight: 400, marginBottom: 6 },
  modalSubtitle: { fontSize: 13, color: "#666", marginBottom: 18 },
  consentBox:    { background: "#f7f7f5", borderRadius: 12, padding: "14px 16px", textAlign: "left", marginBottom: 16, display: "flex", flexDirection: "column", gap: 10 },
  consentItem:   { display: "flex", gap: 10, alignItems: "flex-start" },
  consentIcon:   { fontSize: 13, fontWeight: 700, marginTop: 2, flexShrink: 0 },
  consentText:   { fontSize: 13, lineHeight: 1.5, color: "#333" },
  consentLegal:  { fontSize: 11, color: "#888", lineHeight: 1.5, marginBottom: 16, textAlign: "left" },
  acceptBtn:     { width: "100%", padding: "16px", background: "#1a1a2e", color: "#fff", border: "none", borderRadius: 12, fontSize: 15, fontWeight: 700, cursor: "pointer" },
  modalNote:     { fontSize: 10, color: "#bbb", marginTop: 10 },
  // Main
  main:          { maxWidth: 1080, margin: "0 auto", padding: "22px 24px 80px" },
  mainMobile:    { padding: "14px 12px 100px" },
  layout:        { display: "flex", gap: 18, alignItems: "flex-start" },
  layoutTablet:  { gap: 14 },
  sessionsCol:   { flex: 1, minWidth: 0, display: "flex", flexDirection: "column", gap: 14 },
  sidebar:       { width: 200, flexShrink: 0, position: "sticky", top: 66, display: "flex", flexDirection: "column", gap: 14 },
  sidebarTablet: { width: 170 },
  sidebarHint:   { background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 12, padding: "13px" },
  sidebarHintTitle: { fontSize: 11, fontWeight: 700, color: "#0369a1", marginBottom: 5 },
  sidebarHintText:  { fontSize: 10, color: "#0369a1", lineHeight: 1.5, marginBottom: 10 },
  sidebarUploadBtn: { width: "100%", padding: "8px", background: "#0369a1", color: "#fff", border: "none", borderRadius: 7, fontSize: 12, fontWeight: 600, cursor: "pointer" },
  // Summary Banner
  banner:         { background: "#fff", border: "1px solid #e8e8e4", borderRadius: 12, padding: "14px 16px", marginBottom: 16 },
  bannerTitle:    { fontSize: 9, fontWeight: 700, letterSpacing: ".12em", textTransform: "uppercase", color: "#888", marginBottom: 10 },
  bannerGrid:     { display: "flex", flexWrap: "wrap" },
  bannerGridMobile:{ justifyContent: "space-around" },
  bannerStat:     { flex: "1 1 70px", textAlign: "center", padding: "6px 4px" },
  bannerStatMobile:{ flex: "1 1 30%" },
  bannerIcon:     { fontSize: 18, marginBottom: 3 },
  bannerVal:      { fontSize: 18, fontWeight: 700, letterSpacing: "-0.02em", fontFamily: "'DM Serif Display',serif" },
  bannerValMobile:{ fontSize: 16 },
  bannerStatLabel:{ fontSize: 9, color: "#888", textTransform: "uppercase", letterSpacing: ".07em", marginTop: 2 },
  // Rest Timer
  timerCard:       { background: "#fff", border: "1px solid #e8e8e4", borderRadius: 12, padding: "14px", textAlign: "center", transition: "background .3s,border-color .3s" },
  timerCardDone:   { background: "#fffbeb", borderColor: "#f59e0b" },
  timerCardMobile: { border: "none", padding: "8px 0 0" },
  timerLabel:      { fontSize: 9, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#888", marginBottom: 8 },
  timerPresets:    { display: "flex", gap: 6, justifyContent: "center", marginBottom: 12 },
  presetBtn:       { padding: "4px 8px", border: "1px solid #e0e0d8", borderRadius: 6, background: "#f7f7f5", fontSize: 11, cursor: "pointer", fontWeight: 500, color: "#555" },
  presetBtnMobile: { padding: "10px 14px", fontSize: 14, borderRadius: 8 },
  presetActive:    { background: "#1a1a2e", color: "#fff", borderColor: "#1a1a2e" },
  timerRing:       { position: "relative", margin: "0 auto 12px" },
  timerNum:        { position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontVariantNumeric: "tabular-nums" },
  timerBtns:       { display: "flex", gap: 8, justifyContent: "center" },
  timerStartBtn:   { padding: "7px 18px", background: "#1a1a2e", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" },
  timerStopBtn:    { padding: "7px 18px", background: "#ef4444", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" },
  timerBtnMobile:  { padding: "13px 28px", fontSize: 16, borderRadius: 10 },
  timerDoneMsg:    { marginTop: 10, fontSize: 13, color: "#d97706", fontWeight: 600 },
  // Mobile FAB timer
  timerFab:        { position: "fixed", bottom: 80, right: 16, width: 52, height: 52, borderRadius: "50%", background: "#1a1a2e", color: "#fff", border: "none", fontSize: 22, cursor: "pointer", boxShadow: "0 4px 16px rgba(0,0,0,.2)", zIndex: 50, display: "flex", alignItems: "center", justifyContent: "center" },
  timerDrawer:     { position: "fixed", bottom: 0, left: 0, right: 0, background: "#fff", borderRadius: "20px 20px 0 0", boxShadow: "0 -4px 30px rgba(0,0,0,.12)", padding: "0 24px 40px", zIndex: 99 },
  timerDrawerHeader:{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "16px 0 8px" },
  timerDrawerTitle: { fontSize: 16, fontWeight: 700 },
  timerDrawerClose: { background: "none", border: "none", fontSize: 20, cursor: "pointer", color: "#888", padding: "4px 8px" },
  // Cards
  card:            { background: "#fff", borderRadius: 12, border: "1px solid #e8e8e4", overflow: "hidden" },
  cardHeader:      { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 14px", background: "#fafaf8", borderBottom: "1px solid #e8e8e4", gap: 8, flexWrap: "wrap" },
  cardHeaderMobile:{ padding: "12px 14px" },
  cardHeaderLeft:  { display: "flex", alignItems: "center", gap: 8 },
  cardHeaderRight: { display: "flex", alignItems: "center", gap: 8 },
  sessionBadge:    { fontSize: 10, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#888", padding: "3px 7px", background: "#f0f0ec", borderRadius: 4 },
  sessionNameBadge:{ fontSize: 12, fontWeight: 600, color: "#1a1a2e" },
  collapseBtn:     { background: "none", border: "none", cursor: "pointer", fontSize: 11, color: "#aaa", padding: "4px" },
  collapseBtnMobile:{ fontSize: 16, padding: "6px 8px" },
  prCountBadge:    { fontSize: 11, fontWeight: 600, color: "#d97706", background: "#fffbeb", border: "1px solid #fcd34d", padding: "2px 7px", borderRadius: 4 },
  durBadge:        { fontSize: 11, fontWeight: 600, color: "#2d6a4f", background: "#d8f3dc", padding: "2px 7px", borderRadius: 4 },
  removeSessionBtn:{ background: "none", border: "1px solid #fca5a5", color: "#ef4444", borderRadius: 6, fontSize: 11, padding: "3px 8px", cursor: "pointer" },
  removeSessionBtnMobile:{ fontSize: 13, padding: "6px 12px" },
  sessionMeta:     { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, padding: "14px", borderBottom: "1px solid #f0f0ec" },
  sessionMetaMobile:{ gridTemplateColumns: "1fr", gap: 14, padding: "14px 14px" },
  field:           { display: "flex", flexDirection: "column", gap: 4 },
  label:           { fontSize: 10, fontWeight: 600, letterSpacing: ".08em", textTransform: "uppercase", color: "#888" },
  input:           { padding: "8px 10px", borderRadius: 7, border: "1px solid #e0e0d8", fontSize: 14, background: "#fff", color: "#1a1a2e", outline: "none", width: "100%", boxSizing: "border-box", fontFamily: "inherit" },
  inputMobile:     { padding: "13px 12px", fontSize: 16, borderRadius: 8 },   // 16px prevents iOS zoom
  timeWrap:        { display: "flex", alignItems: "center", gap: 6 },
  timeWrapMobile:  { gap: 8, flexWrap: "wrap" },
  timeDisplay:     { fontSize: 13, fontVariantNumeric: "tabular-nums", fontWeight: 500, minWidth: 56, color: "#333" },
  timeDisplayMobile:{ fontSize: 16, minWidth: 64 },
  timeBtn:         { padding: "7px 11px", background: "#1a1a2e", color: "#fff", border: "none", borderRadius: 7, fontSize: 11, cursor: "pointer", fontWeight: 600, whiteSpace: "nowrap" },
  timeBtnClear:    { background: "#f0f0ec", color: "#666" },
  timeBtnMobile:   { padding: "12px 16px", fontSize: 14, borderRadius: 8 },
  exerciseList:    { padding: "14px", display: "flex", flexDirection: "column", gap: 0 },
  exListHeader:    { marginBottom: 8 },
  exListTitle:     { fontSize: 10, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#888" },
  exRow:           { position: "relative", padding: "12px", background: "#fafaf8", borderRadius: 10, marginBottom: 8, border: "1px solid #efefeb" },
  exRowPR:         { background: "#fffbeb", borderColor: "#fcd34d" },
  prevHint:        { fontSize: 11, color: "#2d6a4f", background: "#d8f3dc", padding: "6px 10px", borderRadius: 6, marginBottom: 10 },
  exGrid:          { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 },
  exGridMobile:    { gridTemplateColumns: "1fr 1fr", gap: 12 },
  weightWrap:      { display: "flex", gap: 6 },
  unitSelect:      { padding: "8px 6px", border: "1px solid #e0e0d8", borderRadius: 7, fontSize: 13, background: "#fff", color: "#555", cursor: "pointer", fontFamily: "inherit" },
  unitSelectMobile:{ padding: "13px 8px", fontSize: 16 },
  exFooter:        { display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 10 },
  prBtn:           { padding: "6px 12px", border: "1px solid #e0e0d8", borderRadius: 7, background: "#fff", fontSize: 12, cursor: "pointer", color: "#888", fontWeight: 500 },
  prBtnMobile:     { padding: "11px 16px", fontSize: 15, borderRadius: 8 },
  prBtnActive:     { background: "#fffbeb", borderColor: "#fcd34d", color: "#d97706", fontWeight: 700 },
  removeExBtn:     { background: "none", border: "1px solid #fca5a5", borderRadius: 6, cursor: "pointer", fontSize: 11, color: "#ef4444", padding: "4px 8px" },
  removeExBtnMobile:{ padding: "10px 14px", fontSize: 14, borderRadius: 8 },
  addExBtn:        { marginTop: 6, padding: "9px 14px", background: "none", border: "1.5px dashed #ccc", borderRadius: 9, fontSize: 12, fontWeight: 600, color: "#888", cursor: "pointer", width: "100%", textAlign: "center" },
  addExBtnMobile:  { padding: "14px", fontSize: 15, borderRadius: 10 },
  addSessionBtn:   { padding: "12px", background: "none", border: "2px dashed #ccc", borderRadius: 12, fontSize: 13, fontWeight: 600, color: "#888", cursor: "pointer", width: "100%", textAlign: "center" },
  addSessionBtnMobile: { padding: "16px", background: "none", border: "2px dashed #ccc", borderRadius: 12, fontSize: 16, fontWeight: 600, color: "#888", cursor: "pointer", width: "100%", textAlign: "center" },
  // History Panel
  histPanel:       { background: "#fff", border: "1px solid #e8e8e4", borderRadius: 12, overflow: "hidden", marginTop: 18 },
  histHeader:      { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 16px", background: "#fafaf8", borderBottom: "1px solid #e8e8e4", flexWrap: "wrap", gap: 8 },
  histTitle:       { fontSize: 13, fontWeight: 600 },
  histTabs:        { display: "flex", gap: 4 },
  histTab:         { padding: "5px 12px", border: "1px solid #e0e0d8", borderRadius: 6, background: "#fff", fontSize: 12, cursor: "pointer", color: "#666" },
  histTabMobile:   { padding: "9px 14px", fontSize: 14 },
  histTabActive:   { background: "#1a1a2e", color: "#fff", borderColor: "#1a1a2e" },
  compareWrap:     { padding: "14px", overflowX: "auto" },
  compTable:       { width: "100%", borderCollapse: "collapse", fontSize: 12 },
  compTh:          { background: "#f7f7f5", padding: "6px 10px", textAlign: "left", fontSize: 9, textTransform: "uppercase", letterSpacing: ".06em", color: "#888", fontWeight: 600 },
  compRow:         { borderBottom: "1px solid #f0f0ec" },
  compTd:          { padding: "8px 10px", verticalAlign: "middle" },
  // Mobile compare cards
  compareCard:     { border: "1px solid #efefeb", borderRadius: 10, overflow: "hidden" },
  compareCardTitle:{ display: "flex", alignItems: "center", gap: 8, padding: "10px 12px", background: "#fafaf8", borderBottom: "1px solid #efefeb", fontSize: 13 },
  compareCardRow:  { display: "flex" },
  compareCardCol:  { flex: 1, padding: "10px 12px" },
  compareCardDivider:{ width: 1, background: "#efefeb" },
  compareCardColLabel:{ fontSize: 9, fontWeight: 600, letterSpacing: ".08em", textTransform: "uppercase", color: "#aaa", marginBottom: 4 },
  compareCardColVal:{ fontSize: 15, fontWeight: 700 },
  exName:          { fontWeight: 600, textTransform: "capitalize" },
  prChip:          { fontSize: 11, background: "#fffbeb", border: "1px solid #fcd34d", color: "#d97706", padding: "2px 6px", borderRadius: 4 },
  subVal:          { fontSize: 10, color: "#888" },
  noData:          { color: "#ccc", fontSize: 11 },
  pastWrap:        { padding: "14px", display: "flex", flexDirection: "column", gap: 10, maxHeight: 420, overflowY: "auto" },
  pastSession:     { border: "1px solid #efefeb", borderRadius: 8, overflow: "hidden" },
  pastSessionHdr:  { display: "flex", gap: 8, alignItems: "baseline", padding: "8px 10px", background: "#f7f7f5", borderBottom: "1px solid #efefeb" },
  pastDate:        { fontSize: 12, fontWeight: 600, color: "#555" },
  pastName:        { fontSize: 12, color: "#888" },
  pastExList:      { padding: "8px 10px", display: "flex", flexDirection: "column", gap: 5 },
  pastEx:          { display: "flex", justifyContent: "space-between", fontSize: 12 },
  pastExName:      { fontWeight: 500, textTransform: "capitalize", color: "#333" },
  pastExDets:      { color: "#888" },
  emptyMsg:        { color: "#aaa", fontSize: 13, textAlign: "center", padding: "20px 0" },
  // Footer
  footer:          { textAlign: "center", padding: "16px 24px", fontSize: 10, color: "#bbb", borderTop: "1px solid #eee", background: "#fff" },
  footerMobile:    { fontSize: 11, padding: "14px 16px" },
};
