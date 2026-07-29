import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx-js-style";

/* ============================================================
   Handball-Tracker
   Teams → Kader → Spiele → Live-Tracking → Nachbereitung → Statistik
   Persistenz: window.storage (account-gebunden)
   ============================================================ */

const KEY = "handball:v1";
const EMPTY = { teams: [], games: [] };

/* ---------- Design-Tokens ---------- */
const C = {
  ink: "#16202E",
  sub: "#5C6B7E",
  bg: "#F1F4F8",
  card: "#FFFFFF",
  line: "#DCE3EB",
  blue: "#2E5EAA",
  blueDark: "#1F3F76",
  blueSoft: "#E4EBF7",
  orange: "#E8622C",
  orangeSoft: "#FCEBE2",
  green: "#1F9D55",
  greenSoft: "#E3F4EA",
  red: "#D33F49",
  redSoft: "#FBE6E8",
  yellow: "#E0A82E",
  yellowSoft: "#FBF2DC",
  court: "#EAF0F7",
  courtHover: "#CFE0F4",
  navy: "#12233F",
};
const MONO = "ui-monospace, SFMono-Regular, Menlo, Consolas, monospace";
const SANS = "-apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif";

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const clone = (o) => JSON.parse(JSON.stringify(o));
const todayISO = () => new Date().toISOString().slice(0, 10);
const fmtDate = (iso) => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}.${m}.${y}`;
};
const fmtClock = (sec) => {
  const m = Math.floor(sec / 60), s = sec % 60;
  return `${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`;
};
const actMinute = (a) => Math.min(60, Math.floor((a.sec || 0) / 60) + 1);

/* ---------- Zonen & Ziele ---------- */
const ZONE_LABEL = {
  KREIS: "Kreis", AUSSEN_L: "Außen links", AUSSEN_R: "Außen rechts",
  DURCH_L: "Durchbruch links", DURCH_M: "Durchbruch Mitte", DURCH_R: "Durchbruch rechts",
  RUECK_L: "Rückraum links", RUECK_M: "Rückraum Mitte", RUECK_R: "Rückraum rechts",
  SIEBEN_M: "7 Meter", KONTER: "Konter", FREIWURF: "Freiwurf",
};
const TARGET_LABEL = {
  t1: "oben links", t2: "oben Mitte", t3: "oben rechts",
  t4: "halbhoch links", t5: "Mitte", t6: "halbhoch rechts",
  t7: "unten links", t8: "unten Mitte", t9: "unten rechts",
  POST: "Pfosten/Latte", WIDE: "daneben",
};
const RESULT_LABEL = { goal: "Tor", saved: "Gehalten", post: "Pfosten", wide: "Vorbei" };
const PENALTY_LABEL = { p2: "2 Minuten", yellow: "Gelbe Karte", red: "Rote Karte", blue: "Blaue Karte" };
const SIMPLE_LABEL = {
  assist: "Assist", steal: "Steal", block: "Block", tf: "Technischer Fehler",
  m7won: "7m herausgeholt", m7caused: "7m verursacht",
};

/* ---------- Zentrale Spalten-Definitionen (App-Tabelle + Excel-Export) ----------
   Regel: Neues Statistik-Feld = EIN Eintrag hier – Tabelle und Export ziehen mit.
   `short` = Spaltenkopf in der App, `label` = ausgeschriebener Kopf im Excel. */
const FIELD_STAT_COLS = [
  { key: "goals", label: "Tore", short: "Tore" },
  { key: "shots", label: "Würfe", short: "Würfe" },
  { key: "quote", label: "Quote", short: "Quote", percent: true },
  { key: "min", label: "Min.", short: "Min.", isMin: true },
  { key: "assist", label: "Assists", short: "Ass." },
  { key: "tf", label: "Techn. Fehler", short: "TF" },
  { key: "steal", label: "Steals", short: "St." },
  { key: "block", label: "Blocks", short: "Bl." },
  { key: "p2", label: "2 Minuten", short: "2min" },
  { key: "m7won", label: "7m herausgeholt", short: "7m+" },
  { key: "m7caused", label: "7m verursacht", short: "7m−" },
];
const KEEPER_STAT_COLS = [
  { key: "saves", label: "Paraden", short: "Paraden" },
  { key: "conceded", label: "Gegentore", short: "Gegentore" },
  { key: "quote", label: "Quote", short: "Quote", percent: true },
  { key: "min", label: "Min.", short: "Min.", isMin: true },
];
/* Gesamtzeile (App-Tabelle + Excel-Export): summiert alle Zähl-Spalten.
   Quote wird aus den Summen neu berechnet, Min. bleibt leer ("–"). */
const statTotals = (rows, cols) => {
  const t = {};
  for (const c of cols) {
    if (c.percent || c.isMin) continue;
    t[c.key] = rows.reduce((a, r) => a + (r[c.key] || 0), 0);
  }
  return t;
};
/* Karten erscheinen nur im Spieler-Detail (App) bzw. Spieler-Blatt (Excel). */
const PLAYER_CARD_ROWS = [
  { key: "yellow", label: "Gelbe Karten" },
  { key: "red", label: "Rote Karten" },
  { key: "blue", label: "Blaue Karten" },
];

/* ---------- Positionen (feste Slot-Reihenfolge) ---------- */
const POSITIONS = ["TW", "LA", "RL", "RM", "RR", "RA", "KL"];
const POS_LABEL = {
  TW: "Torwart", LA: "Linksaußen", RL: "Rückraum links", RM: "Rückraum Mitte",
  RR: "Rückraum rechts", RA: "Rechtsaußen", KL: "Kreisläufer",
};
const P2_SECONDS = 120;

/* ---------- Taktiktafel: Formationen ---------- */
const ATTACK_FORMATIONS = ["5:1", "4:2", "6:0"];
const DEFENSE_FORMATIONS = ["6:0", "5:1", "3:2:1"];
/* Angriff: rein visuelles Mapping der Lineup-Slots auf das Halbfeld.
   label = angezeigte Rolle (nur Anzeige – das Datenmodell bleibt unverändert).
   Der TW wird nicht im Angriff, sondern im Tor der Abwehrtafel dargestellt. */
const ATTACK_LAYOUT = {
  "5:1": {
    LA: { x: 38, y: 52, label: "LA" },
    RL: { x: 44, y: 218, label: "RL" },
    RM: { x: 200, y: 252, label: "RM" },
    RR: { x: 356, y: 218, label: "RR" },
    RA: { x: 362, y: 52, label: "RA" },
    KL: { x: 200, y: 140, label: "KL" },
  },
  "4:2": {
    LA: { x: 38, y: 52, label: "LA" },
    RL: { x: 44, y: 218, label: "RL" },
    RM: { x: 258, y: 140, label: "KL" },
    RR: { x: 356, y: 218, label: "RR" },
    RA: { x: 362, y: 52, label: "RA" },
    KL: { x: 142, y: 140, label: "KL" },
  },
  "6:0": {
    LA: { x: 38, y: 52, label: "LA" },
    RL: { x: 44, y: 205, label: "RL" },
    RM: { x: 162, y: 248, label: "RM" },
    KL: { x: 238, y: 248, label: "RM" },
    RR: { x: 356, y: 205, label: "RR" },
    RA: { x: 362, y: 52, label: "RA" },
  },
};
/* Abwehr: 6 Slots je Formation mit rollen-benannten Keys.
   DEFENSE_ORDER definiert die kanonische Reihenfolge (hinten links→rechts,
   dann Mitte, dann vorne). Beim Formationswechsel wird über diese Reihenfolge
   umgemappt – die Zuordnung bleibt erhalten, nur die Gruppierung ändert sich. */
const DEFENSE_LAYOUT = {
  "6:0": {
    hinten0: { x: 38, y: 52 }, hinten1: { x: 104, y: 126 }, hinten2: { x: 168, y: 142 },
    hinten3: { x: 232, y: 142 }, hinten4: { x: 296, y: 126 }, hinten5: { x: 362, y: 52 },
  },
  /* Halbspieler: Abstand zu „hinten Mitte" ×1,5, Höhe entlang des Abwehrbogens */
  "5:1": {
    hinten0: { x: 38, y: 52 }, hinten1: { x: 92, y: 119 }, hinten2: { x: 200, y: 142 },
    hinten3: { x: 308, y: 119 }, hinten4: { x: 362, y: 52 },
    vorne0: { x: 200, y: 228 },
  },
  "3:2:1": {
    hinten0: { x: 38, y: 52 }, hinten1: { x: 200, y: 142 }, hinten2: { x: 362, y: 52 },
    /* Halbspieler: gleiche Breite wie bei 5:1 (x=92/308), auf der 9m-Linie */
    mitte0: { x: 92, y: 162 }, mitte1: { x: 308, y: 162 },
    vorne0: { x: 200, y: 252 },
  },
};
const DEFENSE_ORDER = {
  "6:0": ["hinten0", "hinten1", "hinten2", "hinten3", "hinten4", "hinten5"],
  "5:1": ["hinten0", "hinten1", "hinten2", "hinten3", "hinten4", "vorne0"],
  "3:2:1": ["hinten0", "hinten1", "hinten2", "mitte0", "mitte1", "vorne0"],
};
/* Formationswechsel: Belegung über die kanonische Reihenfolge übernehmen. */
export function remapDefenseSlots(slots, fromF, toF) {
  const seq = DEFENSE_ORDER[fromF].map((k) => (slots || {})[k] || null);
  const out = {};
  DEFENSE_ORDER[toF].forEach((k, i) => { if (seq[i]) out[k] = seq[i]; });
  return out;
}
/* Vorbelegung der Abwehr-Slots aus der Angriffsaufstellung: LA, RL, RM, RR, RA, KL. */
const DEF_PREFILL_ORDER = ["LA", "RL", "RM", "RR", "RA", "KL"];
export function defaultDefenseSlots(lineup, formation) {
  const out = {};
  if (!lineup) return out;
  DEFENSE_ORDER[formation].forEach((k, i) => {
    const pid = lineup[DEF_PREFILL_ORDER[i]];
    if (pid) out[k] = pid;
  });
  return out;
}
/* Spielerwechsel: der neue Spieler übernimmt direkt die frei werdende
   Abwehrposition des ausgewechselten Spielers (in place). Hatte der
   ausgewechselte Spieler keinen Slot, bleibt der neue unzugeordnet. */
export function subDefenseSlots(slots, outId, inId) {
  if (!slots || !outId) return;
  // Veraltete Einträge des einwechselnden Spielers entfernen (keine Duplikate)
  if (inId) for (const k of Object.keys(slots)) if (slots[k] === inId && slots[k] !== outId) delete slots[k];
  for (const k of Object.keys(slots)) {
    if (slots[k] === outId) {
      if (inId) slots[k] = inId; else delete slots[k];
    }
  }
}

/* ---------- Spielfeld-Geometrie (1 m = 20 px, Tor oben) ---------- */
const yOn = (r, x) => {
  if (x <= 170) return Math.sqrt(Math.max(0, r * r - (x - 170) ** 2));
  if (x >= 230) return Math.sqrt(Math.max(0, r * r - (x - 230) ** 2));
  return r;
};
const seg = (r, x1, x2, n = 30) => {
  const pts = [];
  for (let i = 0; i <= n; i++) {
    const x = x1 + ((x2 - x1) * i) / n;
    pts.push([x, yOn(r, x)]);
  }
  return pts;
};
const path = (pts, close = true) =>
  pts.map((p, i) => `${i ? "L" : "M"}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join(" ") + (close ? " Z" : "");

const ZONE_PATHS = {
  AUSSEN_L: path([[0, 0], [80, 0], ...seg(180, 80, 0)]),
  AUSSEN_R: path([[400, 0], [320, 0], ...seg(180, 320, 400)]),
  KREIS: path([[80, 0], [320, 0], ...seg(120, 320, 80)]),
  DURCH_L: path([...seg(120, 80, 160), ...seg(180, 160, 80)]),
  DURCH_M: path([...seg(120, 160, 240), ...seg(180, 240, 160)]),
  DURCH_R: path([...seg(120, 240, 320), ...seg(180, 320, 240)]),
  RUECK_L: path([...seg(180, 0, 133.3), [133.3, 292], [0, 292]]),
  RUECK_M: path([...seg(180, 133.3, 266.7), [266.7, 292], [133.3, 292]]),
  RUECK_R: path([...seg(180, 266.7, 400), [400, 292], [266.7, 292]]),
};
const ZONE_SHORT = {
  AUSSEN_L: "Außen", AUSSEN_R: "Außen", KREIS: "Kreis",
  DURCH_L: "DB li", DURCH_M: "DB Mitte", DURCH_R: "DB re",
  RUECK_L: "RR links", RUECK_M: "RR Mitte", RUECK_R: "RR rechts",
};
const ZONE_LABEL_POS = {
  AUSSEN_L: [38, 44], AUSSEN_R: [362, 44], KREIS: [200, 58],
  DURCH_L: [116, 148], DURCH_M: [200, 154], DURCH_R: [284, 148],
  RUECK_L: [62, 238], RUECK_M: [200, 238], RUECK_R: [338, 238],
};

/* ---------- Spielfeld-SVG (Abwurfzone wählen) ---------- */
function CourtPicker({ onPick }) {
  const [hover, setHover] = useState(null);
  return (
    <div>
      <svg viewBox="-4 -8 408 306" style={{ width: "100%", display: "block", touchAction: "manipulation" }}>
        {Object.keys(ZONE_PATHS).map((z) => (
          <path key={z} d={ZONE_PATHS[z]}
            fill={hover === z ? C.courtHover : C.court}
            stroke="#fff" strokeWidth="2" style={{ cursor: "pointer" }}
            onMouseEnter={() => setHover(z)} onMouseLeave={() => setHover(null)}
            onClick={() => onPick(z)} />
        ))}
        {/* 6m-Linie */}
        <path d={path(seg(120, 50, 350), false)} fill="none" stroke={C.blue} strokeWidth="3" pointerEvents="none" />
        {/* 9m-Linie gestrichelt */}
        <path d={path(seg(180, 0, 400), false)} fill="none" stroke={C.blue} strokeWidth="2.5"
          strokeDasharray="12 9" pointerEvents="none" />
        {/* 7m- und 4m-Markierung */}
        <line x1="188" y1="140" x2="212" y2="140" stroke={C.ink} strokeWidth="3" pointerEvents="none" />
        <line x1="195" y1="80" x2="205" y2="80" stroke={C.ink} strokeWidth="3" pointerEvents="none" />
        {/* Torlinie + Tor */}
        <line x1="0" y1="0" x2="400" y2="0" stroke={C.sub} strokeWidth="2" pointerEvents="none" />
        <line x1="170" y1="-3" x2="230" y2="-3" stroke={C.red} strokeWidth="7" pointerEvents="none" />
        {Object.entries(ZONE_LABEL_POS).map(([z, [x, y]]) => (
          <text key={z} x={x} y={y} textAnchor="middle" pointerEvents="none"
            style={{ fontFamily: SANS, fontSize: 12, fontWeight: 700, fill: hover === z ? C.blueDark : C.sub }}>
            {ZONE_SHORT[z]}
          </text>
        ))}
      </svg>
      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
        {["SIEBEN_M", "KONTER", "FREIWURF"].map((z) => (
          <button key={z} onClick={() => onPick(z)} style={{
            flex: 1, padding: "14px 8px", borderRadius: 12, border: `2px solid ${C.blue}`,
            background: C.blueSoft, color: C.blueDark, fontFamily: SANS, fontSize: 15,
            fontWeight: 700, cursor: "pointer",
          }}>{ZONE_LABEL[z]}</button>
        ))}
      </div>
    </div>
  );
}

/* ---------- Tor-SVG (Zielzone wählen) ---------- */
function GoalPicker({ onPick }) {
  const [hover, setHover] = useState(null);
  const cells = [];
  for (let r = 0; r < 3; r++)
    for (let c = 0; c < 3; c++)
      cells.push({ id: `t${r * 3 + c + 1}`, x: 46 + c * 96, y: 46 + r * 64 });
  const cellFill = (id) => (hover === id ? C.courtHover : "rgba(234,240,247,0.65)");
  return (
    <svg viewBox="0 0 380 268" style={{ width: "100%", display: "block", touchAction: "manipulation" }}>
      <defs>
        <pattern id="rwH" width="24" height="14" patternUnits="userSpaceOnUse">
          <rect width="24" height="14" fill="#fff" />
          <rect width="12" height="14" fill={C.red} />
        </pattern>
        <pattern id="rwV" width="14" height="24" patternUnits="userSpaceOnUse">
          <rect width="14" height="24" fill="#fff" />
          <rect width="14" height="12" fill={C.red} />
        </pattern>
        <pattern id="net" width="16" height="16" patternUnits="userSpaceOnUse">
          <path d="M0 0 L16 0 M0 0 L0 16" stroke={C.line} strokeWidth="1" />
        </pattern>
      </defs>
      {/* daneben (außerhalb) */}
      <rect x="0" y="0" width="380" height="268" rx="14"
        fill={hover === "WIDE" ? "#E7DDD2" : "#F2ECE4"} style={{ cursor: "pointer" }}
        onMouseEnter={() => setHover("WIDE")} onMouseLeave={() => setHover(null)}
        onClick={() => onPick("WIDE")} />
      <text x="190" y="24" textAnchor="middle" pointerEvents="none"
        style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, fill: "#8A7B66" }}>daneben</text>
      {/* Netz + Rasterzellen */}
      <rect x="46" y="46" width="288" height="192" fill="url(#net)" pointerEvents="none" />
      {cells.map((cl) => (
        <rect key={cl.id} x={cl.x} y={cl.y} width="96" height="64"
          fill={cellFill(cl.id)} stroke="#fff" strokeWidth="2" style={{ cursor: "pointer" }}
          onMouseEnter={() => setHover(cl.id)} onMouseLeave={() => setHover(null)}
          onClick={() => onPick(cl.id)} />
      ))}
      {/* Pfosten/Latte (rot-weiß) */}
      <g style={{ cursor: "pointer" }}
        onMouseEnter={() => setHover("POST")} onMouseLeave={() => setHover(null)}
        onClick={() => onPick("POST")}>
        <rect x="32" y="32" width="316" height="14" fill="url(#rwH)" stroke={hover === "POST" ? C.ink : C.sub} strokeWidth={hover === "POST" ? 2.5 : 1} />
        <rect x="32" y="46" width="14" height="192" fill="url(#rwV)" stroke={hover === "POST" ? C.ink : C.sub} strokeWidth={hover === "POST" ? 2.5 : 1} />
        <rect x="334" y="46" width="14" height="192" fill="url(#rwV)" stroke={hover === "POST" ? C.ink : C.sub} strokeWidth={hover === "POST" ? 2.5 : 1} />
      </g>
      {/* Boden */}
      <line x1="20" y1="238" x2="360" y2="238" stroke={C.sub} strokeWidth="2.5" pointerEvents="none" />
    </svg>
  );
}

/* ---------- Taktiktafel: Halbfeld & Spieler-Marker ---------- */
const BOARD_VB = { x: -4, y: -8, w: 408, h: 306 };
function HalfCourt({ svgRef, style, onPointerDown, children }) {
  return (
    <svg ref={svgRef} viewBox="-4 -8 408 306" onPointerDown={onPointerDown}
      style={{ width: "100%", display: "block", ...style }}>
      <rect x="0" y="0" width="400" height="292" rx="4" fill={C.court} />
      {/* 6m-Linie */}
      <path d={path(seg(120, 50, 350), false)} fill="none" stroke={C.blue} strokeWidth="3" pointerEvents="none" />
      {/* 9m-Linie gestrichelt */}
      <path d={path(seg(180, 0, 400), false)} fill="none" stroke={C.blue} strokeWidth="2.5"
        strokeDasharray="12 9" pointerEvents="none" />
      {/* 7m-Markierung */}
      <line x1="188" y1="140" x2="212" y2="140" stroke={C.ink} strokeWidth="3" pointerEvents="none" />
      {/* Torlinie + Tor */}
      <line x1="0" y1="0" x2="400" y2="0" stroke={C.sub} strokeWidth="2" pointerEvents="none" />
      <line x1="170" y1="-3" x2="230" y2="-3" stroke={C.red} strokeWidth="7" pointerEvents="none" />
      {children}
    </svg>
  );
}

/* Marker: state = "on" (Spieler), "free" (leerer Slot), "locked" (Strafzeit).
   Bei "locked" steht der Countdown in `name`, bei "on" die Trikotnummer in `number`. */
function BoardMarker({ x, y, role, state, number, name, sub, isTW, highlight, dimmed, onClick, onPointerDown }) {
  return (
    <g onClick={onClick} onPointerDown={onPointerDown}
      style={{
        cursor: onPointerDown ? "grab" : onClick ? "pointer" : "default",
        touchAction: "none", opacity: dimmed ? 0.25 : 1,
      }}>
      <rect x={x - 32} y={y - 38} width="64" height="82" fill="transparent" />
      {highlight && <circle cx={x} cy={y} r="27" fill="none" stroke={C.blueDark} strokeWidth="3" strokeDasharray="6 4" pointerEvents="none" />}
      {role && (
        <text x={x} y={y - 27} textAnchor="middle" pointerEvents="none"
          style={{ fontFamily: SANS, fontSize: 10, fontWeight: 800, fill: C.sub, letterSpacing: "0.05em" }}>{role}</text>
      )}
      {state === "locked" && (
        <>
          <circle cx={x} cy={y} r="20" fill={C.redSoft} stroke={C.red} strokeWidth="2.5" />
          <text x={x} y={y + 4} textAnchor="middle" pointerEvents="none"
            style={{ fontFamily: MONO, fontSize: 11, fontWeight: 800, fill: C.red }}>{name}</text>
        </>
      )}
      {state === "free" && (
        <>
          <circle cx={x} cy={y} r="20" fill="#fff" stroke={C.sub} strokeWidth="2" strokeDasharray="5 4" />
          <text x={x} y={y + 4} textAnchor="middle" pointerEvents="none"
            style={{ fontFamily: SANS, fontSize: 10, fontWeight: 700, fill: C.sub }}>frei</text>
        </>
      )}
      {state === "on" && (
        <>
          <circle cx={x} cy={y} r="20" fill={isTW ? C.orange : C.blue} stroke="#fff" strokeWidth="2.5" />
          <text x={x} y={y + 5} textAnchor="middle" pointerEvents="none"
            style={{ fontFamily: MONO, fontSize: 15, fontWeight: 800, fill: "#fff" }}>{number}</text>
        </>
      )}
      {sub && (
        <text x={x < 60 ? x - 24 : x > 340 ? x + 24 : x} y={y + 34}
          textAnchor={x < 60 ? "start" : x > 340 ? "end" : "middle"} pointerEvents="none"
          style={{ fontFamily: SANS, fontSize: 11, fontWeight: 700, fill: state === "locked" ? C.red : C.ink }}>{sub}</text>
      )}
    </g>
  );
}
const firstName = (n) => (n || "").split(" ")[0].slice(0, 10);

/* Angriffs-Halbfeld: automatisches Mapping der Lineup-Slots auf die Formation.
   Tap auf Spieler = Aktions-Flow, Tap auf freien/gesperrten Slot = Wechsel-Modal. */
export function AttackBoard({ team, lineup, formation, penByPos, sec, onPlayer, onSlot }) {
  const layout = ATTACK_LAYOUT[formation] || ATTACK_LAYOUT["5:1"];
  const twPen = penByPos.TW || null;
  const keeper = lineup && lineup.TW ? team.players.find((x) => x.id === lineup.TW) : null;
  return (
    <HalfCourt style={{ touchAction: "manipulation" }}>
      {/* Torhüter im Tor (nur Anzeige; Aktions-Tap wie bei Feldspielern) */}
      {twPen ? (
        <BoardMarker x={200} y={20} state="locked"
          name={fmtClock(Math.max(0, twPen.startSec + P2_SECONDS - sec))}
          sub="TW gesperrt" onClick={() => onSlot("TW")} />
      ) : keeper ? (
        <BoardMarker x={200} y={20} state="on" isTW number={keeper.number}
          sub={firstName(keeper.name)} onClick={() => onPlayer(keeper.id)} />
      ) : (
        <BoardMarker x={200} y={20} state="free" onClick={() => onSlot("TW")} />
      )}
      {POSITIONS.filter((pos) => layout[pos]).map((pos) => {
        const c = layout[pos];
        const pen = penByPos[pos];
        const pid = lineup ? lineup[pos] : null;
        const pl = pid ? team.players.find((x) => x.id === pid) : null;
        if (pen) {
          const planned = pen.plannedInId ? team.players.find((x) => x.id === pen.plannedInId) : null;
          return <BoardMarker key={pos} x={c.x} y={c.y} role={c.label} state="locked"
            name={fmtClock(Math.max(0, pen.startSec + P2_SECONDS - sec))}
            sub={planned ? `→ ${firstName(planned.name)}` : "gesperrt"}
            onClick={() => onSlot(pos)} />;
        }
        if (!pl) return <BoardMarker key={pos} x={c.x} y={c.y} role={c.label} state="free" onClick={() => onSlot(pos)} />;
        return <BoardMarker key={pos} x={c.x} y={c.y} role={c.label} state="on" isTW={pl.pos === "TW"}
          number={pl.number} sub={firstName(pl.name)} onClick={() => onPlayer(pid)} />;
      })}
    </HalfCourt>
  );
}

/* Abwehr-Halbfeld: 6 Slots, Zuordnung per Drag & Drop (native Pointer-Events).
   Tap (ohne Ziehen) auf einen Spieler startet den normalen Aktions-Flow.
   Spieler mit Strafzeit bleiben verschiebbar, damit die Restabwehr angepasst
   werden kann. Der TW steht im Tor (nicht Teil der 6 Slots).
   onDrop(pid, fromKey|null, toKey|null): toKey=null → Zuordnung entfernen. */
export function DefenseBoard({ team, formation, slots, onFieldIds, penByPid, sec, unassigned, onDrop, onPlayer, keeper, keeperPen, onKeeperSlot }) {
  const layout = DEFENSE_LAYOUT[formation] || DEFENSE_LAYOUT["6:0"];
  const svgRef = useRef(null);
  const [drag, setDrag] = useState(null); // {pid, fromKey, tap, sx, sy, active, x, y}
  const dragRef = useRef(null);
  dragRef.current = drag;

  const toSvg = (cx, cy) => {
    const r = svgRef.current.getBoundingClientRect();
    return {
      x: BOARD_VB.x + ((cx - r.left) * BOARD_VB.w) / r.width,
      y: BOARD_VB.y + ((cy - r.top) * BOARD_VB.h) / r.height,
    };
  };
  const nearestKey = (x, y) => {
    let best = null, bd = Infinity;
    for (const [k, p] of Object.entries(layout)) {
      const d = Math.hypot(p.x - x, p.y - y);
      if (d < bd) { bd = d; best = k; }
    }
    return bd <= 46 ? best : null;
  };
  /* tap=false: reiner Drag-Marker (z. B. Strafzeit) ohne Tap-Aktion */
  const startDrag = (pid, tap = true) => (e) => {
    e.preventDefault(); e.stopPropagation();
    const fromKey = Object.keys(slots).find((k) => slots[k] === pid) || null;
    setDrag({ pid, fromKey, tap, sx: e.clientX, sy: e.clientY, active: false, x: 0, y: 0 });
  };
  useEffect(() => {
    if (!drag) return;
    const move = (e) => {
      const d = dragRef.current;
      if (!d || !svgRef.current) return;
      const active = d.active || Math.hypot(e.clientX - d.sx, e.clientY - d.sy) > 8;
      const p = toSvg(e.clientX, e.clientY);
      setDrag({ ...d, active, x: p.x, y: p.y });
    };
    const up = (e) => {
      const d = dragRef.current;
      if (d && svgRef.current) {
        if (d.active) {
          const p = toSvg(e.clientX, e.clientY);
          onDrop(d.pid, d.fromKey, nearestKey(p.x, p.y));
        } else if (d.tap && onPlayer) {
          onPlayer(d.pid); // Tap ohne Ziehen → Aktions-Flow
        }
      }
      setDrag(null);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
    window.addEventListener("pointercancel", up);
    return () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      window.removeEventListener("pointercancel", up);
    };
  }, [!!drag]);

  const target = drag && drag.active ? nearestKey(drag.x, drag.y) : null;
  const dragPl = drag ? team.players.find((x) => x.id === drag.pid) : null;

  return (
    <div>
      <HalfCourt svgRef={svgRef} style={{ touchAction: "none" }}>
        {/* Torhüter im Tor */}
        {keeperPen ? (
          <BoardMarker x={200} y={20} state="locked"
            name={fmtClock(Math.max(0, keeperPen.startSec + P2_SECONDS - sec))}
            sub="TW gesperrt" onClick={onKeeperSlot} />
        ) : keeper ? (
          <BoardMarker x={200} y={20} state="on" isTW number={keeper.number}
            sub={firstName(keeper.name)} onClick={onPlayer ? () => onPlayer(keeper.id) : onKeeperSlot} />
        ) : (
          <BoardMarker x={200} y={20} state="free" onClick={onKeeperSlot} />
        )}
        {Object.entries(layout).map(([k, p]) => {
          const pid = slots[k] || null;
          const pl = pid ? team.players.find((x) => x.id === pid) : null;
          const pen = pid ? penByPid[pid] : null;
          if (pen) {
            return <BoardMarker key={k} x={p.x} y={p.y} state="locked"
              name={fmtClock(Math.max(0, pen.startSec + P2_SECONDS - sec))}
              sub={firstName(pl?.name)} highlight={target === k}
              dimmed={drag && drag.active && drag.fromKey === k}
              onPointerDown={startDrag(pid, false)} />;
          }
          if (!pl || !onFieldIds.has(pid)) {
            return <BoardMarker key={k} x={p.x} y={p.y} state="free" highlight={target === k} />;
          }
          return <BoardMarker key={k} x={p.x} y={p.y} state="on" isTW={false}
            number={pl.number} sub={firstName(pl.name)} highlight={target === k}
            dimmed={drag && drag.active && drag.fromKey === k}
            onPointerDown={startDrag(pid)} />;
        })}
        {drag && drag.active && dragPl && (
          <g pointerEvents="none" opacity="0.9">
            <circle cx={drag.x} cy={drag.y} r="22" fill={C.blueDark} stroke="#fff" strokeWidth="2.5" />
            <text x={drag.x} y={drag.y + 5} textAnchor="middle"
              style={{ fontFamily: MONO, fontSize: 15, fontWeight: 800, fill: "#fff" }}>{dragPl.number}</text>
          </g>
        )}
      </HalfCourt>
      {unassigned.length > 0 && (
        <div style={{ marginTop: 8 }}>
          <div style={{ fontFamily: SANS, fontSize: 12, fontWeight: 700, color: C.sub, marginBottom: 6 }}>
            Nicht zugeordnet – auf einen Slot ziehen:
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {unassigned.map((p) => (
              <div key={p.id} onPointerDown={startDrag(p.id)} style={{
                fontFamily: SANS, fontSize: 14, fontWeight: 700, color: C.blueDark,
                background: C.blueSoft, border: `2px dashed ${C.blue}`, borderRadius: 10,
                padding: "9px 12px", touchAction: "none", cursor: "grab", userSelect: "none",
                opacity: drag && drag.active && drag.pid === p.id ? 0.35 : 1,
              }}>#{p.number} {p.name}</div>
            ))}
          </div>
        </div>
      )}
      <div style={{ fontFamily: SANS, fontSize: 11, color: C.sub, marginTop: 6 }}>
        Antippen erfasst eine Aktion · Ziehen tauscht Slots · außerhalb ablegen entfernt die Zuordnung.
      </div>
    </div>
  );
}
function FormationSelect({ value, options, onChange }) {
  return (
    <select value={value} onChange={(e) => onChange(e.target.value)}
      style={{ ...inputStyle, width: "auto", padding: "8px 12px", fontWeight: 700, fontFamily: MONO }}>
      {options.map((o) => <option key={o} value={o}>{o}</option>)}
    </select>
  );
}

/* ---------- UI-Bausteine ---------- */
const btnBase = {
  fontFamily: SANS, fontWeight: 700, border: "none", borderRadius: 12,
  cursor: "pointer", padding: "12px 18px", fontSize: 15,
};
function Btn({ children, kind = "primary", small, style, ...p }) {
  const kinds = {
    primary: { background: C.blue, color: "#fff" },
    accent: { background: C.orange, color: "#fff" },
    ghost: { background: "transparent", color: C.blueDark, border: `2px solid ${C.line}` },
    soft: { background: C.blueSoft, color: C.blueDark },
    danger: { background: C.redSoft, color: C.red },
    green: { background: C.green, color: "#fff" },
  };
  return (
    <button {...p} style={{
      ...btnBase, ...(kinds[kind] || kinds.primary),
      ...(small ? { padding: "8px 12px", fontSize: 13, borderRadius: 10 } : {}),
      ...style,
    }}>{children}</button>
  );
}
function ConfirmBtn({ label = "Löschen", confirmLabel = "Sicher?", onConfirm, small = true }) {
  const [armed, setArmed] = useState(false);
  useEffect(() => {
    if (!armed) return;
    const t = setTimeout(() => setArmed(false), 2500);
    return () => clearTimeout(t);
  }, [armed]);
  return (
    <Btn kind="danger" small={small} onClick={() => (armed ? (setArmed(false), onConfirm()) : setArmed(true))}
      style={armed ? { background: C.red, color: "#fff" } : {}}>
      {armed ? confirmLabel : label}
    </Btn>
  );
}
function Card({ children, style }) {
  return <div style={{ background: C.card, borderRadius: 16, border: `1px solid ${C.line}`, padding: 16, ...style }}>{children}</div>;
}
function Modal({ title, onClose, children, wide }) {
  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(18,35,63,0.55)", zIndex: 50,
      display: "flex", alignItems: "center", justifyContent: "center", padding: 12,
    }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{
        background: "#fff", borderRadius: 18, width: "100%", maxWidth: wide ? 620 : 480,
        maxHeight: "94vh", overflowY: "auto", padding: 18, boxShadow: "0 18px 50px rgba(0,0,0,0.25)",
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={{ fontFamily: SANS, fontWeight: 800, fontSize: 18, color: C.ink }}>{title}</div>
          <button onClick={onClose} style={{
            ...btnBase, padding: "6px 12px", background: C.bg, color: C.sub, fontSize: 16,
          }}>✕</button>
        </div>
        {children}
      </div>
    </div>
  );
}
function Field({ label, children }) {
  return (
    <label style={{ display: "block", marginBottom: 12 }}>
      <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: C.sub, marginBottom: 5 }}>{label}</div>
      {children}
    </label>
  );
}
const inputStyle = {
  width: "100%", boxSizing: "border-box", padding: "12px 12px", borderRadius: 10,
  border: `2px solid ${C.line}`, fontFamily: SANS, fontSize: 15, color: C.ink, background: "#fff",
};
function Empty({ children }) {
  return (
    <div style={{
      padding: "28px 16px", textAlign: "center", color: C.sub, fontFamily: SANS,
      fontSize: 14, border: `2px dashed ${C.line}`, borderRadius: 14,
    }}>{children}</div>
  );
}
function StatusChip({ status }) {
  const map = {
    open: { t: "Offen", bg: C.blueSoft, c: C.blueDark },
    live: { t: "Läuft", bg: C.orangeSoft, c: C.orange },
    finished: { t: "Beendet", bg: C.greenSoft, c: C.green },
  };
  const m = map[status];
  return <span style={{
    fontFamily: SANS, fontSize: 12, fontWeight: 800, padding: "4px 10px",
    borderRadius: 999, background: m.bg, color: m.c,
  }}>{m.t}</span>;
}

/* ---------- Spielzeit: Aufstellung & Minuten aus Wechselereignissen ---------- */
/* Ereignisse in game.subs:
   - Wechsel:        { id, sec, half, pos, outId|null, inId|null, reason? }
   - Positionstausch:{ id, sec, half, kind: "swap", posA, posB, aId, bId }
   reason: "sub" | "p2" | "card" | "p2in" – p2/card/p2in sind automatisch
   erzeugt (srcAction verweist auf die Strafen-Aktion). */
function lineupAndMinutes(game, endSec, startSec = 0) {
  if (!game.startLineup) return { lineup: null, minutes: null, minutesByPos: null };
  const time = {};
  const byPos = {}; // pid -> { pos: sec } – Spielzeit je Lineup-Position
  const on = {}; // pos -> { pid, since }
  for (const pos of POSITIONS) {
    const pid = game.startLineup[pos];
    if (pid) on[pos] = { pid, since: 0 };
  }
  // startSec: Zeit wird nur innerhalb [startSec, endSec] gutgeschrieben.
  // Die Wechsel-Historie wird trotzdem vollständig abgespielt, damit die
  // Aufstellung an der unteren Fenstergrenze korrekt rekonstruiert wird.
  const close = (pos, at) => {
    const c = on[pos];
    if (!c) return;
    const dur = Math.max(0, at - Math.max(c.since, startSec));
    if (dur > 0) {
      time[c.pid] = (time[c.pid] || 0) + dur;
      const t = (byPos[c.pid] ||= {});
      t[pos] = (t[pos] || 0) + dur;
    }
    delete on[pos];
  };
  const evts = [...(game.subs || [])].sort((a, b) => (a.sec || 0) - (b.sec || 0));
  for (const e of evts) {
    if (e.sec > endSec) break;
    if (e.kind === "swap") {
      // Intervalle schließen und neu öffnen, damit die Zeit vor dem Tausch
      // der alten Position gutgeschrieben wird (Gesamtsumme bleibt gleich).
      const a = on[e.posA], b = on[e.posB];
      close(e.posA, e.sec); close(e.posB, e.sec);
      if (b) on[e.posA] = { pid: b.pid, since: e.sec };
      if (a) on[e.posB] = { pid: a.pid, since: e.sec };
      continue;
    }
    if (e.outId != null) close(e.pos, e.sec);
    if (e.inId) {
      if (on[e.pos]) close(e.pos, e.sec);
      on[e.pos] = { pid: e.inId, since: e.sec };
    }
  }
  const lineup = {};
  for (const pos of POSITIONS) lineup[pos] = on[pos]?.pid || null;
  for (const pos of POSITIONS) close(pos, endSec);
  return { lineup, minutes: time, minutesByPos: byPos };
}
function gameEndSec(game) {
  let m = game.timerSec || 0;
  for (const a of game.actions || []) m = Math.max(m, a.sec || 0);
  for (const s of game.subs || []) m = Math.max(m, s.sec || 0);
  return m;
}
/* Aktion inkl. automatisch erzeugter Wechsel/Strafsperren entfernen */
function removeActionCascade(g, actionId) {
  g.actions = g.actions.filter((x) => x.id !== actionId);
  if (g.subs) g.subs = g.subs.filter((s) => s.srcAction !== actionId);
  if (g.penalties) g.penalties = g.penalties.filter((p) => p.srcAction !== actionId);
}

/* ---------- Statistik-Berechnung ---------- */
function computeScore(game) {
  let us = 0, them = 0;
  for (const a of game.actions || []) {
    if (a.type === "throw" && a.result === "goal") a.side === "us" ? us++ : them++;
  }
  return { us, them };
}
function playerName(team, id) {
  const p = team.players.find((x) => x.id === id);
  return p ? p.name : "Unbekannt";
}
function playerLabel(team, id) {
  const p = team.players.find((x) => x.id === id);
  return p ? `#${p.number} ${p.name}` : "Unbekannt";
}
/* opts:
   - fromSec/toSec: Zeitfenster – nur Aktionen mit fromSec <= sec < toSec zählen,
     Spielzeit wird auf das Fenster begrenzt (Wechsel-Historie bleibt vollständig).
   - endSec: explizites Spielende (z. B. aktueller Timer bei Live-Statistik). */
export function aggregate(team, games, opts = {}) {
  const fromSec = opts.fromSec || 0;
  const toSec = opts.toSec == null ? Infinity : opts.toSec;
  const inWin = (s) => (s || 0) >= fromSec && (s || 0) < toSec;
  const P = {}; // Feldstatistik pro Spieler
  const K = {}; // Torhüter
  const Z = {}; // Zonen (eigene Würfe)
  const M = {}; // Spielzeit in Sekunden
  const MP = {}; // Spielzeit je Position: pid -> { pos: sec }
  const R = {}; // Kader-Einsätze: Anzahl Spiele im Filter, in denen der Spieler im Kader stand
  let minutesTracked = 0;
  let oppP2 = 0; // gegnerische 2-min-Strafen (Überzahl-Situationen)
  const ensureP = (id) => (P[id] ||= { goals: 0, shots: 0, assist: 0, tf: 0, steal: 0, block: 0, p2: 0, yellow: 0, red: 0, blue: 0, m7won: 0, m7caused: 0 });
  const ensureK = (id) => (K[id] ||= { saves: 0, conceded: 0 });
  for (const g of games) {
    for (const id of g.roster || []) R[id] = (R[id] || 0) + 1;
    for (const a of g.actions || []) {
      if (!inWin(a.sec)) continue;
      if (a.type === "oppPenalty") {
        oppP2++;
      } else if (a.type === "throw") {
        if (a.side === "us") {
          const s = ensureP(a.playerId);
          s.shots++; if (a.result === "goal") s.goals++;
          if (a.assistId && a.result === "goal") ensureP(a.assistId).assist++;
          const z = (Z[a.zone] ||= { shots: 0, goals: 0 });
          z.shots++; if (a.result === "goal") z.goals++;
        } else if (a.keeperId) {
          const k = ensureK(a.keeperId);
          if (a.result === "saved") k.saves++;
          if (a.result === "goal") k.conceded++;
        }
      } else if (a.type === "penalty") {
        ensureP(a.playerId)[a.kind]++;
      } else if (SIMPLE_LABEL[a.type]) {
        ensureP(a.playerId)[a.type]++;
      }
    }
    if (g.startLineup) {
      minutesTracked++;
      const end = Math.min(opts.endSec != null ? opts.endSec : gameEndSec(g), toSec);
      const { minutes, minutesByPos } = lineupAndMinutes(g, end, fromSec);
      for (const [id, s] of Object.entries(minutes || {})) M[id] = (M[id] || 0) + s;
      for (const [id, per] of Object.entries(minutesByPos || {})) {
        const t = (MP[id] ||= {});
        for (const [pos, s] of Object.entries(per)) t[pos] = (t[pos] || 0) + s;
      }
    }
  }
  // Spieler mit Spielzeit, aber ohne Aktionen, trotzdem in die Tabellen aufnehmen
  for (const id of Object.keys(M)) {
    const pl = team.players.find((p) => p.id === id);
    if (pl?.pos === "TW") ensureK(id); else ensureP(id);
  }
  return { P, K, Z, M, MP, R, minutesTracked, gamesCount: games.length, oppP2 };
}

/* ---------- Aktion als Text (Spielverlauf) ---------- */
function actionText(team, a) {
  const m = `${actMinute(a)}'`;
  if (a.type === "throw") {
    const zone = ZONE_LABEL[a.zone] || "";
    if (a.side === "us") {
      const who = playerName(team, a.playerId);
      const ass = a.assistId ? `, Assist: ${playerName(team, a.assistId)}` : "";
      if (a.result === "goal") return `${m} Tor – ${who} (${zone}${ass})`;
      if (a.result === "saved") return `${m} Wurf gehalten – ${who} (${zone})`;
      if (a.result === "post") return `${m} Pfosten – ${who} (${zone})`;
      return `${m} Wurf vorbei – ${who} (${zone})`;
    }
    const kp = a.keeperId ? playerName(team, a.keeperId) : "Torhüter";
    if (a.result === "goal") return `${m} Gegentor (${zone})`;
    if (a.result === "saved") return `${m} Parade ${kp} (${zone})`;
    if (a.result === "post") return `${m} Gegner Pfosten (${zone})`;
    return `${m} Gegner vorbei (${zone})`;
  }
  if (a.type === "penalty") return `${m} ${PENALTY_LABEL[a.kind]} – ${playerName(team, a.playerId)}`;
  if (a.type === "oppPenalty") return `${m} 2 Minuten – Gegner (Überzahl)`;
  return `${m} ${SIMPLE_LABEL[a.type] || a.type} – ${playerName(team, a.playerId)}`;
}
/* Wechsel-Ereignis als Text */
function subText(team, e) {
  const m = `${Math.min(60, Math.floor((e.sec || 0) / 60) + 1)}'`;
  if (e.kind === "swap")
    return `${m} Positionstausch ${e.posA} ↔ ${e.posB}: ${playerName(team, e.aId)} / ${playerName(team, e.bId)}`;
  if (e.reason === "p2") return `${m} Strafzeit – ${playerName(team, e.outId)} runter (${e.pos})`;
  if (e.reason === "card") return `${m} Disqualifikation – ${playerName(team, e.outId)} runter (${e.pos})`;
  if (e.reason === "p2in") return `${m} Strafzeit vorbei – ${playerName(team, e.inId)} rein (${e.pos})`;
  if (e.outId && e.inId) return `${m} Wechsel ${e.pos}: ${playerName(team, e.inId)} für ${playerName(team, e.outId)}`;
  if (e.inId) return `${m} Einwechslung ${e.pos}: ${playerName(team, e.inId)}`;
  return `${m} Auswechslung ${e.pos}: ${playerName(team, e.outId)}`;
}

/* ============================================================
   Screens
   ============================================================ */

/* ---------- Team-Übersicht ---------- */
function TeamsScreen({ data, update, go }) {
  const [name, setName] = useState("");
  const add = () => {
    const n = name.trim();
    if (!n) return;
    update((d) => { d.teams.push({ id: uid(), name: n, players: [] }); });
    setName("");
  };
  return (
    <div>
      <PageTitle title="Teams" sub="Team anlegen oder auswählen" />
      <Card style={{ marginBottom: 14 }}>
        <div style={{ display: "flex", gap: 8 }}>
          <input style={inputStyle} placeholder="Teamname, z. B. HSG Usingen II"
            value={name} onChange={(e) => setName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && add()} />
          <Btn onClick={add} style={{ whiteSpace: "nowrap" }}>+ Team</Btn>
        </div>
      </Card>
      {data.teams.length === 0 && <Empty>Noch keine Teams. Lege oben dein erstes Team an.</Empty>}
      <div style={{ display: "grid", gap: 10 }}>
        {data.teams.map((t) => {
          const games = data.games.filter((g) => g.teamId === t.id);
          return (
            <Card key={t.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", cursor: "pointer" }}>
              <div onClick={() => go({ name: "team", teamId: t.id, tab: "kader" })} style={{ flex: 1 }}>
                <div style={{ fontFamily: SANS, fontWeight: 800, fontSize: 17, color: C.ink }}>{t.name}</div>
                <div style={{ fontFamily: SANS, fontSize: 13, color: C.sub, marginTop: 3 }}>
                  {t.players.length} Spieler · {games.length} Spiele
                </div>
              </div>
              <ConfirmBtn onConfirm={() => update((d) => {
                d.teams = d.teams.filter((x) => x.id !== t.id);
                d.games = d.games.filter((g) => g.teamId !== t.id);
              })} />
            </Card>
          );
        })}
      </div>
    </div>
  );
}

function PageTitle({ title, sub, onBack, right }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
      {onBack && <Btn kind="ghost" small onClick={onBack} style={{ fontSize: 16 }}>←</Btn>}
      <div style={{ flex: 1 }}>
        <div style={{ fontFamily: SANS, fontWeight: 900, fontSize: 22, color: C.ink, letterSpacing: "-0.02em" }}>{title}</div>
        {sub && <div style={{ fontFamily: SANS, fontSize: 13, color: C.sub, marginTop: 2 }}>{sub}</div>}
      </div>
      {right}
    </div>
  );
}

/* ---------- Team-Detail (Tabs: Kader / Spiele / Statistik) ---------- */
function TeamScreen({ data, update, go, teamId, tab }) {
  const team = data.teams.find((t) => t.id === teamId);
  if (!team) return <Empty>Team nicht gefunden.</Empty>;
  const games = data.games.filter((g) => g.teamId === teamId);
  const tabs = [["kader", "Kader"], ["spiele", "Spiele"], ["stats", "Statistik"]];
  return (
    <div>
      <PageTitle title={team.name} onBack={() => go({ name: "teams" })} />
      <div style={{ display: "flex", gap: 6, marginBottom: 16 }}>
        {tabs.map(([id, label]) => (
          <button key={id} onClick={() => go({ name: "team", teamId, tab: id })} style={{
            ...btnBase, flex: 1, padding: "11px 8px",
            background: tab === id ? C.navy : "#fff", color: tab === id ? "#fff" : C.sub,
            border: `2px solid ${tab === id ? C.navy : C.line}`,
          }}>{label}</button>
        ))}
      </div>
      {tab === "kader" && <RosterTab team={team} update={update} />}
      {tab === "spiele" && <GamesTab team={team} games={games} update={update} go={go} />}
      {tab === "stats" && <StatsTab team={team} games={games} go={go} teamId={teamId} />}
    </div>
  );
}

/* ---------- Kader ---------- */
function RosterTab({ team, update }) {
  const [name, setName] = useState("");
  const [number, setNumber] = useState("");
  const [pos, setPos] = useState("F");
  const add = () => {
    const n = name.trim();
    if (!n) return;
    update((d) => {
      const t = d.teams.find((x) => x.id === team.id);
      t.players.push({ id: uid(), name: n, number: number.trim() || "?", pos });
      t.players.sort((a, b) => (parseInt(a.number) || 99) - (parseInt(b.number) || 99));
    });
    setName(""); setNumber("");
  };
  const field = team.players.filter((p) => p.pos === "F");
  const keepers = team.players.filter((p) => p.pos === "TW");
  const PlayerRow = ({ p }) => (
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 0", borderBottom: `1px solid ${C.line}` }}>
      <span style={{
        fontFamily: MONO, fontWeight: 700, fontSize: 15, color: "#fff", background: p.pos === "TW" ? C.orange : C.blue,
        borderRadius: 8, minWidth: 34, textAlign: "center", padding: "5px 0",
      }}>{p.number}</span>
      <span style={{ flex: 1, fontFamily: SANS, fontSize: 15, fontWeight: 600, color: C.ink }}>{p.name}</span>
      <ConfirmBtn onConfirm={() => update((d) => {
        const t = d.teams.find((x) => x.id === team.id);
        t.players = t.players.filter((x) => x.id !== p.id);
      })} />
    </div>
  );
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <Card>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <input style={{ ...inputStyle, width: 70 }} placeholder="Nr." value={number}
            onChange={(e) => setNumber(e.target.value)} />
          <input style={{ ...inputStyle, flex: 1, minWidth: 160 }} placeholder="Name"
            value={name} onChange={(e) => setName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && add()} />
          <select style={{ ...inputStyle, width: 140 }} value={pos} onChange={(e) => setPos(e.target.value)}>
            <option value="F">Feldspieler</option>
            <option value="TW">Torhüter</option>
          </select>
          <Btn onClick={add}>+ Spieler</Btn>
        </div>
      </Card>
      <Card>
        <SectionH>Feldspieler ({field.length})</SectionH>
        {field.length === 0 ? <Empty>Noch keine Feldspieler.</Empty> : field.map((p) => <PlayerRow key={p.id} p={p} />)}
      </Card>
      <Card>
        <SectionH>Torhüter ({keepers.length})</SectionH>
        {keepers.length === 0 ? <Empty>Noch keine Torhüter.</Empty> : keepers.map((p) => <PlayerRow key={p.id} p={p} />)}
      </Card>
    </div>
  );
}
function SectionH({ children }) {
  return <div style={{
    fontFamily: SANS, fontSize: 13, fontWeight: 800, color: C.sub, textTransform: "uppercase",
    letterSpacing: "0.06em", marginBottom: 8,
  }}>{children}</div>;
}

/* ---------- Spiele-Liste ---------- */
function GamesTab({ team, games, update, go }) {
  const sorted = [...games].sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  return (
    <div style={{ display: "grid", gap: 10 }}>
      <Btn kind="accent" onClick={() => go({ name: "newGame", teamId: team.id })}
        style={{ padding: "16px", fontSize: 16 }}>+ Neues Spiel anlegen</Btn>
      {sorted.length === 0 && <Empty>Noch keine Spiele. Lege dein erstes Spiel an.</Empty>}
      {sorted.map((g) => {
        const sc = computeScore(g);
        return (
          <Card key={g.id}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              <div style={{ flex: 1, minWidth: 150 }}>
                <div style={{ fontFamily: SANS, fontWeight: 800, fontSize: 16, color: C.ink }}>
                  {g.home ? `${team.name} – ${g.opponent}` : `${g.opponent} – ${team.name}`}
                </div>
                <div style={{ fontFamily: SANS, fontSize: 13, color: C.sub, marginTop: 3, display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                  {fmtDate(g.date)} <StatusChip status={g.status} />
                  {g.test && <span style={{
                    fontFamily: SANS, fontSize: 12, fontWeight: 800, padding: "4px 10px",
                    borderRadius: 999, background: C.orangeSoft, color: C.orange,
                  }}>Testspiel</span>}
                </div>
              </div>
              {(g.status !== "open" || g.actions.length > 0) && (
                <div style={{ fontFamily: MONO, fontWeight: 800, fontSize: 20, color: C.ink }}>
                  {g.home ? `${sc.us}:${sc.them}` : `${sc.them}:${sc.us}`}
                </div>
              )}
              <div style={{ display: "flex", gap: 6 }}>
                {g.status === "open" && (
                  <Btn small kind="green" onClick={() => {
                    update((d) => { d.games.find((x) => x.id === g.id).status = "live"; });
                    go({ name: "live", teamId: team.id, gameId: g.id });
                  }}>▶ Starten</Btn>
                )}
                {g.status === "live" && (
                  <Btn small kind="accent" onClick={() => go({ name: "live", teamId: team.id, gameId: g.id })}>▶ Weiter</Btn>
                )}
                <Btn small kind="soft" onClick={() => go({ name: "review", teamId: team.id, gameId: g.id })}>Nachbereitung</Btn>
                <ConfirmBtn onConfirm={() => update((d) => { d.games = d.games.filter((x) => x.id !== g.id); })} />
              </div>
            </div>
          </Card>
        );
      })}
    </div>
  );
}

/* ---------- Spiel anlegen ---------- */
function NewGameScreen({ data, update, go, teamId }) {
  const team = data.teams.find((t) => t.id === teamId);
  const [opp, setOpp] = useState("");
  const [date, setDate] = useState(todayISO());
  const [home, setHome] = useState(true);
  const [test, setTest] = useState(false);
  const [sel, setSel] = useState(() => new Set(team.players.map((p) => p.id)));
  const [lineup, setLineup] = useState({});   // pos -> playerId
  const [pickPos, setPickPos] = useState(null);
  const [err, setErr] = useState("");
  /* Taktiktafel: Formationen + Abwehr-Vorbelegung. defSlots === null bedeutet
     „automatisch aus der Aufstellung" (LA, RL, RM, RR, RA, KL) – erst nach
     manueller Änderung wird die Zuordnung fest gespeichert. */
  const [attackF, setAttackF] = useState("5:1");
  const [defF, setDefF] = useState("6:0");
  const [defSlots, setDefSlots] = useState(null);
  const effDefSlots = defSlots || defaultDefenseSlots(lineup, defF);
  const changeDefFormation = (v) => {
    setDefSlots((s) => (s ? remapDefenseSlots(s, defF, v) : s));
    setDefF(v);
  };
  const ngDrop = (pid, fromKey, toKey) => {
    setDefSlots((s0) => {
      const s = { ...(s0 || defaultDefenseSlots(lineup, defF)) };
      if (!toKey) { if (fromKey) delete s[fromKey]; return s; }
      const prev = s[toKey] || null;
      for (const k of Object.keys(s)) if (s[k] === pid) delete s[k];
      if (prev && fromKey && prev !== pid) s[fromKey] = prev;
      s[toKey] = pid;
      return s;
    });
  };
  const ngOnField = new Set(Object.values(lineup));
  const ngUnassigned = DEF_PREFILL_ORDER.map((ps) => lineup[ps]).filter(Boolean)
    .filter((pid) => !Object.values(effDefSlots).includes(pid))
    .map((pid) => team.players.find((x) => x.id === pid)).filter(Boolean);
  const toggle = (id) => setSel((s) => {
    const n = new Set(s);
    if (n.has(id)) {
      n.delete(id);
      setLineup((lu) => {
        const out = { ...lu };
        for (const pos of POSITIONS) if (out[pos] === id) delete out[pos];
        return out;
      });
    } else n.add(id);
    return n;
  });
  const lineupComplete = POSITIONS.every((pos) => lineup[pos] && sel.has(lineup[pos]));
  const create = (start) => {
    if (!opp.trim()) { setErr("Bitte einen Gegner eintragen."); return; }
    if (sel.size === 0) { setErr("Bitte den Spieltagskader auswählen."); return; }
    if (!lineupComplete) { setErr("Bitte alle 7 Positionen der Startaufstellung besetzen."); return; }
    const id = uid();
    // Abwehr-Zuordnung: auf aufgestellte Spieler beschränken und fest speichern.
    const srcSlots = defSlots || defaultDefenseSlots(lineup, defF);
    const lineupIds = new Set(Object.values(lineup));
    const cleanSlots = {};
    for (const [k, v] of Object.entries(srcSlots)) if (lineupIds.has(v)) cleanSlots[k] = v;
    update((d) => {
      d.games.push({
        id, teamId, opponent: opp.trim(), date, home, test,
        status: start ? "live" : "open", roster: [...sel], actions: [],
        timerSec: 0, half: 1,
        startLineup: { ...lineup }, subs: [], penalties: [],
        activeKeeperId: lineup.TW || null,
        attackFormation: attackF, defenseFormation: defF, defenseSlots: cleanSlots,
      });
    });
    go(start ? { name: "live", teamId, gameId: id } : { name: "team", teamId, tab: "spiele" });
  };
  const chip = (p) => {
    const on = sel.has(p.id);
    return (
      <button key={p.id} onClick={() => toggle(p.id)} style={{
        ...btnBase, padding: "10px 12px", fontSize: 14,
        background: on ? (p.pos === "TW" ? C.orange : C.blue) : "#fff",
        color: on ? "#fff" : C.sub, border: `2px solid ${on ? "transparent" : C.line}`,
      }}>#{p.number} {p.name}</button>
    );
  };
  // Kandidaten für eine Position: gewählter Kader, noch nicht anderweitig aufgestellt
  const assigned = new Set(Object.entries(lineup).filter(([pos]) => pos !== pickPos).map(([, id]) => id));
  const candidates = team.players
    .filter((p) => sel.has(p.id) && !assigned.has(p.id))
    .sort((a, b) => {
      const twFirst = pickPos === "TW" ? -1 : 1;
      if (a.pos !== b.pos) return a.pos === "TW" ? twFirst : -twFirst;
      return (parseInt(a.number) || 99) - (parseInt(b.number) || 99);
    });
  return (
    <div>
      <PageTitle title="Spiel anlegen" sub={team.name} onBack={() => go({ name: "team", teamId, tab: "spiele" })} />
      <Card style={{ marginBottom: 14 }}>
        <Field label="Gegner"><input style={inputStyle} value={opp} onChange={(e) => setOpp(e.target.value)} placeholder="Gegnername" /></Field>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 160 }}>
            <Field label="Datum"><input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
          </div>
          <div style={{ flex: 1, minWidth: 160 }}>
            <Field label="Spielort">
              <div style={{ display: "flex", gap: 6 }}>
                {[["Heim", true], ["Auswärts", false]].map(([l, v]) => (
                  <button key={l} onClick={() => setHome(v)} style={{
                    ...btnBase, flex: 1, padding: "11px 8px",
                    background: home === v ? C.navy : "#fff", color: home === v ? "#fff" : C.sub,
                    border: `2px solid ${home === v ? C.navy : C.line}`,
                  }}>{l}</button>
                ))}
              </div>
            </Field>
          </div>
          <div style={{ flex: 1, minWidth: 160 }}>
            <Field label="Spielart">
              <div style={{ display: "flex", gap: 6 }}>
                {[["Pflichtspiel", false], ["Testspiel", true]].map(([l, v]) => (
                  <button key={l} onClick={() => setTest(v)} style={{
                    ...btnBase, flex: 1, padding: "11px 8px",
                    background: test === v ? (v ? C.orange : C.navy) : "#fff",
                    color: test === v ? "#fff" : C.sub,
                    border: `2px solid ${test === v ? "transparent" : C.line}`,
                  }}>{l}</button>
                ))}
              </div>
            </Field>
            {test && <div style={{ fontFamily: SANS, fontSize: 12, color: C.orange, marginTop: -8 }}>
              Zählt nicht in die Saison-Statistik.
            </div>}
          </div>
        </div>
      </Card>
      <Card style={{ marginBottom: 14 }}>
        <SectionH>Kader für dieses Spiel ({sel.size} gewählt)</SectionH>
        {team.players.length === 0 ? (
          <Empty>Der Kader ist leer – lege zuerst Spieler an.</Empty>
        ) : (
          <>
            <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, margin: "6px 0" }}>Torhüter</div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>{team.players.filter((p) => p.pos === "TW").map(chip)}</div>
            <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, margin: "12px 0 6px" }}>Feldspieler</div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>{team.players.filter((p) => p.pos === "F").map(chip)}</div>
          </>
        )}
      </Card>
      <Card style={{ marginBottom: 14 }}>
        <SectionH>Startaufstellung ({POSITIONS.filter((pos) => lineup[pos]).length}/7)</SectionH>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))", gap: 8 }}>
          {POSITIONS.map((pos) => {
            const pid = lineup[pos];
            const pl = pid ? team.players.find((x) => x.id === pid) : null;
            return (
              <button key={pos} onClick={() => setPickPos(pos)} style={{
                ...btnBase, padding: "12px 10px", textAlign: "left",
                background: pl ? (pos === "TW" ? C.orangeSoft : C.blueSoft) : "#fff",
                border: `2px solid ${pl ? "transparent" : C.line}`,
                display: "flex", flexDirection: "column", gap: 3, color: C.ink,
              }}>
                <span style={{ fontFamily: SANS, fontSize: 11, fontWeight: 800, color: C.sub, textTransform: "uppercase", letterSpacing: "0.05em" }}>
                  {pos} · {POS_LABEL[pos]}
                </span>
                <span style={{ fontFamily: SANS, fontSize: 15, fontWeight: 700, color: pl ? C.ink : C.sub }}>
                  {pl ? `#${pl.number} ${pl.name}` : "– wählen –"}
                </span>
              </button>
            );
          })}
        </div>
        {!lineupComplete && (
          <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, marginTop: 8 }}>
            Alle 7 Positionen müssen besetzt sein, bevor das Spiel angelegt werden kann.
          </div>
        )}
        <div style={{ display: "flex", gap: 16, alignItems: "center", flexWrap: "wrap", marginTop: 14 }}>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <span style={{ fontFamily: SANS, fontSize: 13, fontWeight: 800, color: C.sub }}>Angriffsformation</span>
            <FormationSelect value={attackF} options={ATTACK_FORMATIONS} onChange={setAttackF} />
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <span style={{ fontFamily: SANS, fontSize: 13, fontWeight: 800, color: C.sub }}>Abwehrformation</span>
            <FormationSelect value={defF} options={DEFENSE_FORMATIONS} onChange={changeDefFormation} />
          </div>
        </div>
        {ngOnField.size > 1 && (
          <div style={{ marginTop: 12 }}>
            <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, marginBottom: 6 }}>
              Abwehr-Vorbelegung (automatisch in der Reihenfolge LA, RL, RM, RR, RA, KL) – per Ziehen anpassbar:
            </div>
            <DefenseBoard team={team} formation={defF} slots={effDefSlots}
              onFieldIds={ngOnField} penByPid={{}} sec={0}
              unassigned={ngUnassigned} onDrop={ngDrop}
              keeper={lineup.TW ? team.players.find((x) => x.id === lineup.TW) : null}
              keeperPen={null} onKeeperSlot={() => setPickPos("TW")} />
          </div>
        )}
      </Card>
      {err && <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: C.red, marginBottom: 10 }}>{err}</div>}
      <div style={{ display: "flex", gap: 8 }}>
        <Btn kind="ghost" style={{ flex: 1 }} onClick={() => create(false)}>Speichern</Btn>
        <Btn kind="green" style={{ flex: 2, padding: "16px", fontSize: 16 }} onClick={() => create(true)}>▶ Anlegen & starten</Btn>
      </div>

      {pickPos && (
        <Modal title={`${POS_LABEL[pickPos]} besetzen`} onClose={() => setPickPos(null)}>
          {lineup[pickPos] && (
            <Btn kind="danger" small style={{ marginBottom: 10 }} onClick={() => {
              setLineup((lu) => { const n = { ...lu }; delete n[pickPos]; return n; });
              setPickPos(null);
            }}>Position freimachen</Btn>
          )}
          {candidates.length === 0 ? (
            <Empty>Keine verfügbaren Spieler – Kaderauswahl prüfen.</Empty>
          ) : (
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {candidates.map((p) => (
                <button key={p.id} onClick={() => { setLineup((lu) => ({ ...lu, [pickPos]: p.id })); setPickPos(null); }} style={{
                  ...btnBase, padding: "12px 14px", fontSize: 14,
                  background: p.pos === "TW" ? C.orangeSoft : C.blueSoft,
                  color: C.ink, border: `2px solid ${C.line}`,
                }}>#{p.number} {p.name}</button>
              ))}
            </div>
          )}
        </Modal>
      )}
    </div>
  );
}

/* ---------- Live-Tracking ---------- */
function LiveScreen({ data, update, go, teamId, gameId }) {
  const team = data.teams.find((t) => t.id === teamId);
  const game = data.games.find((g) => g.id === gameId);
  const [sec, setSec] = useState(game?.timerSec || 0);
  const [running, setRunning] = useState(false);
  const [pending, setPending] = useState(null); // {side, playerId?, assistId?, pickShooter?, zone?, target?}
  const [subModal, setSubModal] = useState(null); // {selPos?}
  const [showStats, setShowStats] = useState(false); // Live-Statistik-Overlay
  const [defOpen, setDefOpen] = useState(true); // Abwehrtafel ein-/ausklappen (reiner UI-Zustand)
  const secRef = useRef(sec);
  secRef.current = sec;

  useEffect(() => {
    if (!running) return;
    const iv = setInterval(() => setSec((s) => s + 1), 1000);
    return () => clearInterval(iv);
  }, [running]);
  // Timerstand beim Verlassen sichern
  useEffect(() => () => {
    update((d) => { const g = d.games.find((x) => x.id === gameId); if (g) g.timerSec = secRef.current; });
  }, []);

  // Abgelaufene 2-min-Strafen auflösen (geplanter Spieler kommt automatisch rein)
  useEffect(() => {
    if (!game || !game.startLineup) return;
    const due = (game.penalties || []).filter((p) => sec >= p.startSec + P2_SECONDS);
    if (due.length === 0) return;
    const dueIds = new Set(due.map((p) => p.id));
    update((d) => {
      const g = d.games.find((x) => x.id === gameId);
      if (!g) return;
      g.timerSec = secRef.current;
      for (const p of (g.penalties || []).filter((x) => dueIds.has(x.id))) {
        if (p.plannedInId) {
          g.subs.push({
            id: uid(), sec: p.startSec + P2_SECONDS, half: g.half || 1,
            pos: p.pos, outId: null, inId: p.plannedInId, reason: "p2in", srcAction: p.srcAction,
          });
        }
        // Der eintretende Spieler (bzw. der zurückkehrende Bestrafte) übernimmt
        // den Abwehr-Slot; kommt niemand, wird der Slot frei.
        if (g.defenseSlots) subDefenseSlots(g.defenseSlots, p.playerId, p.plannedInId || null);
      }
      g.penalties = (g.penalties || []).filter((x) => !dueIds.has(x.id));
    });
  }, [sec]);

  if (!team || !game) return <Empty>Spiel nicht gefunden.</Empty>;
  const num = (p) => parseInt(p.number) || 99;
  const roster = game.roster
    .map((id) => team.players.find((p) => p.id === id))
    .filter(Boolean)
    .sort((a, b) => (a.pos === b.pos ? num(a) - num(b) : a.pos === "TW" ? -1 : 1));
  const keepers = roster.filter((p) => p.pos === "TW");
  const sc = computeScore(game);
  const half = game.half || 1;

  const hasLineup = !!game.startLineup;
  const { lineup } = hasLineup ? lineupAndMinutes(game, sec) : { lineup: null };
  const penByPos = {};
  for (const p of game.penalties || []) penByPos[p.pos] = p;
  const onFieldIds = new Set(lineup ? POSITIONS.map((pos) => lineup[pos]).filter(Boolean) : []);
  const bench = hasLineup ? roster.filter((p) => !onFieldIds.has(p.id)) : [];
  const keeperNow = lineup ? lineup.TW : game.activeKeeperId;
  const activeKeeper = team.players.find((p) => p.id === keeperNow);

  /* Taktiktafel: Formationen + Abwehr-Slots (Legacy-Spiele ohne gespeicherte
     Slots folgen bis zur ersten manuellen Änderung der Angriffsaufstellung). */
  const attackFormation = game.attackFormation || "5:1";
  const defenseFormation = game.defenseFormation || "6:0";
  const defSlots = game.defenseSlots || defaultDefenseSlots(lineup, defenseFormation);
  const penByPid = {};
  for (const p of game.penalties || []) penByPid[p.playerId] = p;
  const defAssigned = new Set(Object.values(defSlots));
  const defUnassigned = hasLineup
    ? POSITIONS.filter((ps) => ps !== "TW").map((ps) => lineup[ps]).filter(Boolean)
        .filter((pid) => !defAssigned.has(pid))
        .map((pid) => team.players.find((x) => x.id === pid)).filter(Boolean)
    : [];

  const persist = (fn) => update((d) => {
    const g = d.games.find((x) => x.id === gameId);
    g.timerSec = secRef.current;
    fn(g, d);
  });
  const addAction = (a) => persist((g) => {
    g.actions.push({ id: uid(), sec: secRef.current, half: g.half || 1, ...a });
  });
  const undo = () => persist((g) => {
    const lastA = g.actions[g.actions.length - 1];
    if (lastA) removeActionCascade(g, lastA.id);
  });
  const setHalf = (h) => {
    if (h === 2 && sec < 1800) setSec(1800);
    persist((g) => { g.half = h; });
  };
  const endGame = () => {
    setRunning(false);
    persist((g) => { g.status = "finished"; });
    go({ name: "review", teamId, gameId });
  };

  /* Abwehr-Zuordnung materialisieren (Legacy-Spiele: Vorbelegung aus der
     aktuellen Aufstellung), damit ab jetzt kein Auto-Fill mehr passiert. */
  const ensureDefense = (g) => {
    g.defenseFormation ||= "6:0";
    if (!g.defenseSlots) {
      const { lineup: lu } = lineupAndMinutes(g, secRef.current);
      g.defenseSlots = defaultDefenseSlots(lu, g.defenseFormation);
    }
  };
  const setAttackFormation = (v) => persist((g) => { g.attackFormation = v; });
  const setDefenseFormation = (v) => persist((g) => {
    ensureDefense(g);
    g.defenseSlots = remapDefenseSlots(g.defenseSlots, g.defenseFormation, v);
    g.defenseFormation = v;
  });
  /* Drag & Drop in der Abwehr: Slot↔Slot tauscht, Chip→Slot belegt (verdrängt
     ggf. in „nicht zugeordnet"), Ablegen außerhalb entfernt die Zuordnung.
     Auch Spieler mit Strafzeit sind verschiebbar (Restabwehr anpassen). */
  const dropDefense = (pid, fromKey, toKey) => persist((g) => {
    ensureDefense(g);
    const s = g.defenseSlots;
    if (!toKey) { if (fromKey) delete s[fromKey]; return; }
    const prev = s[toKey] || null;
    for (const k of Object.keys(s)) if (s[k] === pid) delete s[k];
    if (prev && fromKey && prev !== pid) s[fromKey] = prev;
    s[toKey] = pid;
  });

  /* Strafe erfassen – bei 2min/Rot/Blau geht der Spieler automatisch runter */
  const addPenalty = (kind, playerId) => {
    const aid = uid();
    persist((g) => {
      g.actions.push({ id: aid, sec: secRef.current, half: g.half || 1, type: "penalty", kind, playerId });
      if (!g.startLineup || (kind !== "p2" && kind !== "red" && kind !== "blue")) return;
      const { lineup: lu } = lineupAndMinutes(g, secRef.current);
      const pos = POSITIONS.find((ps) => lu[ps] === playerId);
      if (!pos) return; // Spieler stand nicht auf dem Feld
      ensureDefense(g); // Abwehr-Slot bleibt erhalten, wird aber gesperrt angezeigt
      g.subs ||= []; g.penalties ||= [];
      g.subs.push({
        id: uid(), sec: secRef.current, half: g.half || 1,
        pos, outId: playerId, inId: null, reason: kind === "p2" ? "p2" : "card", srcAction: aid,
      });
      g.penalties.push({
        id: uid(), pos, playerId, startSec: secRef.current,
        plannedInId: kind === "p2" ? playerId : null, srcAction: aid,
      });
    });
    setPending(null);
  };

  /* Wurf-Flow abschließen */
  const finishThrow = (result) => {
    const p = pending;
    addAction({
      type: "throw", side: p.side, playerId: p.playerId || null,
      zone: p.zone, target: p.target, result,
      assistId: p.assistId || null,
      keeperId: p.side === "them" ? keeperNow : null,
    });
    setPending(null);
  };

  /* Wechsel-Modal-Logik */
  const clickFieldPos = (pos) => {
    if (!subModal) return;
    const selPos = subModal.selPos;
    if (!selPos) { setSubModal({ selPos: pos }); return; }
    if (selPos === pos) { setSubModal({}); return; }
    // Feld ↔ Feld
    if (penByPos[selPos] || penByPos[pos]) return; // gesperrte Slots nicht tauschen
    const aId = lineup[selPos], bId = lineup[pos];
    if (aId && bId) {
      persist((g) => {
        g.subs.push({ id: uid(), sec: secRef.current, half: g.half || 1, kind: "swap", posA: selPos, posB: pos, aId, bId });
      });
      setSubModal(null);
    } else if (aId && !bId) {
      // Positionswechsel eines Spielers auf freien Slot
      persist((g) => {
        g.subs.push({ id: uid(), sec: secRef.current, half: g.half || 1, pos: selPos, outId: aId, inId: null, reason: "sub" });
        g.subs.push({ id: uid(), sec: secRef.current, half: g.half || 1, pos, outId: null, inId: aId, reason: "sub" });
      });
      setSubModal(null);
    } else {
      setSubModal({ selPos: pos });
    }
  };
  const clickBench = (pid) => {
    if (!subModal || !subModal.selPos) return;
    const selPos = subModal.selPos;
    const pen = penByPos[selPos];
    if (pen) {
      // Slot gesperrt: Spieler kommt erst nach Ablauf der Strafe
      persist((g) => {
        const pp = (g.penalties || []).find((x) => x.id === pen.id);
        if (pp) pp.plannedInId = pid;
      });
      setSubModal(null);
      return;
    }
    persist((g) => {
      ensureDefense(g);
      const outId = lineup[selPos] || null;
      g.subs.push({
        id: uid(), sec: secRef.current, half: g.half || 1,
        pos: selPos, outId, inId: pid, reason: "sub",
      });
      // Der neue Spieler übernimmt direkt die Abwehrposition des ausgewechselten.
      subDefenseSlots(g.defenseSlots, outId, pid);
    });
    setSubModal(null);
  };
  /* Torhüterwechsel über die "Im Tor"-Chips (erzeugt TW-Wechselereignis) */
  const setKeeper = (pid) => {
    if (!hasLineup) { persist((g) => { g.activeKeeperId = pid; }); return; }
    if (penByPos.TW || lineup.TW === pid) return;
    persist((g) => {
      g.subs.push({
        id: uid(), sec: secRef.current, half: g.half || 1,
        pos: "TW", outId: lineup.TW || null, inId: pid, reason: "sub",
      });
      g.activeKeeperId = pid;
    });
  };

  /* Aktionen + Wechsel als gemeinsamer Verlauf */
  const feed = [
    ...game.actions.map((a) => ({ kind: "action", e: a })),
    ...(game.subs || []).map((s) => ({ kind: "sub", e: s })),
  ].sort((x, y) => (x.e.sec || 0) - (y.e.sec || 0));
  const last = feed.slice(-5).reverse();

  const posBadge = (pos, bg) => (
    <span style={{
      fontFamily: SANS, fontWeight: 800, fontSize: 10, color: "#fff",
      background: bg || C.navy, borderRadius: 6, padding: "3px 6px", letterSpacing: "0.04em",
    }}>{pos}</span>
  );

  return (
    <div>
      {/* Scoreboard */}
      <div style={{
        background: C.navy, borderRadius: 18, padding: "14px 16px", marginBottom: 14,
        display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap",
      }}>
        <div style={{ flex: 1, minWidth: 200 }}>
          <div style={{ fontFamily: SANS, fontSize: 12, fontWeight: 700, color: "#9FB4D6", letterSpacing: "0.05em" }}>
            {team.name} vs. {game.opponent}
          </div>
          <div style={{ fontFamily: MONO, fontWeight: 800, fontSize: 40, color: "#fff", lineHeight: 1.1 }}>
            {sc.us}<span style={{ color: "#5E76A0" }}>:</span>{sc.them}
          </div>
        </div>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontFamily: MONO, fontWeight: 800, fontSize: 28, color: C.yellow }}>{fmtClock(sec)}</div>
          <div style={{ display: "flex", gap: 4, marginTop: 4, justifyContent: "center" }}>
            {[1, 2].map((h) => (
              <button key={h} onClick={() => setHalf(h)} style={{
                ...btnBase, padding: "4px 10px", fontSize: 12, borderRadius: 8,
                background: half === h ? C.yellow : "rgba(255,255,255,0.12)",
                color: half === h ? C.navy : "#B9C7DF",
              }}>HZ{h}</button>
            ))}
          </div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          <Btn small kind={running ? "danger" : "green"} onClick={() => {
            setRunning(!running);
            persist(() => {});
          }} style={{ minWidth: 92 }}>{running ? "⏸ Pause" : "▶ Start"}</Btn>
          <Btn small kind="soft" onClick={undo} style={{ background: "rgba(255,255,255,0.12)", color: "#fff" }}
            disabled={game.actions.length === 0}>↩ Rückgängig</Btn>
          <Btn small kind="soft" onClick={() => setShowStats(true)}
            style={{ background: "rgba(255,255,255,0.12)", color: "#fff" }}>📊 Statistik</Btn>
          <ConfirmBtn label="Spiel beenden" confirmLabel="Wirklich beenden?" onConfirm={endGame} />
        </div>
      </div>

      {/* Torhüter aktiv */}
      <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 12, flexWrap: "wrap" }}>
        <span style={{ fontFamily: SANS, fontSize: 13, fontWeight: 800, color: C.sub }}>Im Tor:</span>
        {keepers.length === 0 && <span style={{ fontFamily: SANS, fontSize: 13, color: C.sub }}>kein Torhüter im Kader-Auszug</span>}
        {keepers.map((k) => (
          <button key={k.id} onClick={() => setKeeper(k.id)} style={{
            ...btnBase, padding: "8px 12px", fontSize: 13,
            background: keeperNow === k.id ? C.orange : "#fff",
            color: keeperNow === k.id ? "#fff" : C.sub,
            border: `2px solid ${keeperNow === k.id ? "transparent" : C.line}`,
            opacity: hasLineup && penByPos.TW ? 0.5 : 1,
          }}>#{k.number} {k.name}</button>
        ))}
        {hasLineup && penByPos.TW && (
          <span style={{ fontFamily: SANS, fontSize: 12, color: C.red }}>TW-Slot gesperrt (Strafzeit)</span>
        )}
      </div>

      {/* Spieler-Grid */}
      {hasLineup ? (
        <>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
            <SectionH>Auf dem Feld – Spieler antippen</SectionH>
            <Btn small kind="soft" onClick={() => setSubModal({})}>⇄ Wechsel</Btn>
          </div>
          <Card style={{ marginBottom: 10, padding: 10 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 4 }}>
              <SectionH>Angriff</SectionH>
              <FormationSelect value={attackFormation} options={ATTACK_FORMATIONS} onChange={setAttackFormation} />
            </div>
            <AttackBoard team={team} lineup={lineup} formation={attackFormation}
              penByPos={penByPos} sec={sec}
              onPlayer={(pid) => setPending({ side: "us", playerId: pid })}
              onSlot={(pos) => setSubModal({ selPos: pos })} />
          </Card>
          <Card style={{ marginBottom: 14, padding: 10 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 4 }}>
              <button onClick={() => setDefOpen((o) => !o)} style={{
                ...btnBase, padding: "4px 8px", background: "transparent", color: C.sub,
                display: "flex", alignItems: "center", gap: 6, fontSize: 13, textTransform: "uppercase", letterSpacing: "0.06em",
              }}>
                <span style={{ fontSize: 12 }}>{defOpen ? "▾" : "▸"}</span> Abwehr
              </button>
              <FormationSelect value={defenseFormation} options={DEFENSE_FORMATIONS} onChange={setDefenseFormation} />
            </div>
            {defOpen && (
              <DefenseBoard team={team} formation={defenseFormation} slots={defSlots}
                onFieldIds={onFieldIds} penByPid={penByPid} sec={sec}
                unassigned={defUnassigned} onDrop={dropDefense}
                onPlayer={(pid) => setPending({ side: "us", playerId: pid })}
                keeper={lineup.TW ? team.players.find((x) => x.id === lineup.TW) : null}
                keeperPen={penByPos.TW || null}
                onKeeperSlot={() => setSubModal({ selPos: "TW" })} />
            )}
          </Card>
          {bench.length > 0 && (
            <>
              <SectionH>Bank</SectionH>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 8, marginBottom: 14 }}>
                {bench.map((p) => (
                  <button key={p.id} onClick={() => setPending({ side: "us", playerId: p.id })} style={{
                    ...btnBase, padding: "12px 10px", textAlign: "left", background: "#fff",
                    border: `2px solid ${C.line}`, display: "flex", alignItems: "center", gap: 10, color: C.ink, opacity: 0.85,
                  }}>
                    <span style={{
                      fontFamily: MONO, fontWeight: 800, fontSize: 15, color: "#fff",
                      background: p.pos === "TW" ? C.orange : C.sub, borderRadius: 8, minWidth: 34, textAlign: "center", padding: "5px 0",
                    }}>{p.number}</span>
                    <span style={{ fontSize: 14 }}>{p.name}</span>
                  </button>
                ))}
              </div>
            </>
          )}
        </>
      ) : (
        <>
          <SectionH>Eigene Aktion – Spieler antippen</SectionH>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 8, marginBottom: 14 }}>
            {roster.map((p) => (
              <button key={p.id} onClick={() => setPending({ side: "us", playerId: p.id })} style={{
                ...btnBase, padding: "16px 10px", textAlign: "left", background: "#fff",
                border: `2px solid ${C.line}`, display: "flex", alignItems: "center", gap: 10, color: C.ink,
              }}>
                <span style={{
                  fontFamily: MONO, fontWeight: 800, fontSize: 16, color: "#fff",
                  background: p.pos === "TW" ? C.orange : C.blue, borderRadius: 8, minWidth: 36, textAlign: "center", padding: "6px 0",
                }}>{p.number}</span>
                <span style={{ fontSize: 15 }}>{p.name}</span>
              </button>
            ))}
          </div>
        </>
      )}

      {/* Gegner */}
      <SectionH>Gegner</SectionH>
      <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 8, marginBottom: 14 }}>
        <Btn kind="accent" style={{ padding: "16px", fontSize: 16 }}
          onClick={() => setPending({ side: "them", step: "zone" })}>
          🥅 Wurf des Gegners {activeKeeper ? `(im Tor: ${activeKeeper.name})` : ""}
        </Btn>
        <Btn style={{ padding: "16px", fontSize: 15, background: C.redSoft, color: C.red, border: `2px solid ${C.red}` }}
          onClick={() => addAction({ type: "oppPenalty", kind: "p2" })}>
          ⏱ 2 min Gegner
        </Btn>
      </div>

      {/* Letzte Aktionen */}
      <Card>
        <SectionH>Letzte Aktionen</SectionH>
        {last.length === 0 ? <Empty>Noch keine Aktionen erfasst.</Empty> : last.map(({ kind, e }) => (
          <div key={e.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderBottom: `1px solid ${C.line}` }}>
            <span style={{ flex: 1, fontFamily: SANS, fontSize: 14, color: kind === "sub" ? C.sub : C.ink, fontStyle: kind === "sub" ? "italic" : "normal" }}>
              {kind === "action" ? actionText(team, e) : subText(team, e)}
            </span>
            {kind === "action" && (
              <ConfirmBtn onConfirm={() => persist((g) => { removeActionCascade(g, e.id); })} />
            )}
            {kind === "sub" && !e.srcAction && (
              <ConfirmBtn onConfirm={() => persist((g) => { g.subs = g.subs.filter((x) => x.id !== e.id); })} />
            )}
            {kind === "sub" && e.srcAction && (
              <span style={{ fontFamily: SANS, fontSize: 11, color: C.sub }}>autom.</span>
            )}
          </div>
        ))}
      </Card>

      {/* -------- Wechsel-Modal -------- */}
      {subModal && hasLineup && (
        <Modal title={subModal.selPos ? `Wechsel – ${POS_LABEL[subModal.selPos]}` : "Wechsel – Position wählen"} onClose={() => setSubModal(null)}>
          <SectionH>Auf dem Feld</SectionH>
          <div style={{ display: "grid", gap: 6, marginBottom: 14 }}>
            {POSITIONS.map((pos) => {
              const pen = penByPos[pos];
              const pid = lineup[pos];
              const pl = pid ? team.players.find((x) => x.id === pid) : null;
              const selected = subModal.selPos === pos;
              return (
                <button key={pos} onClick={() => clickFieldPos(pos)} style={{
                  ...btnBase, padding: "10px 12px", textAlign: "left",
                  background: selected ? C.blue : pen ? C.redSoft : C.bg,
                  color: selected ? "#fff" : pen ? C.red : C.ink,
                  border: `2px solid ${selected ? C.blue : C.line}`,
                  display: "flex", alignItems: "center", gap: 10,
                }}>
                  {posBadge(pos, selected ? "rgba(255,255,255,0.25)" : pen ? C.red : C.navy)}
                  <span style={{ flex: 1, fontFamily: SANS, fontSize: 14, fontWeight: 700 }}>
                    {pen
                      ? `gesperrt (⏱ ${fmtClock(Math.max(0, pen.startSec + P2_SECONDS - sec))})${pen.plannedInId ? ` – danach: ${playerName(team, pen.plannedInId)}` : ""}`
                      : pl ? `#${pl.number} ${pl.name}` : "– frei –"}
                  </span>
                </button>
              );
            })}
          </div>
          <SectionH>Bank {subModal.selPos ? `– kommt für ${POS_LABEL[subModal.selPos]}` : ""}</SectionH>
          {!subModal.selPos && (
            <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, marginBottom: 8 }}>
              Zuerst oben die Position antippen, dann den Spieler, der kommt. Zwei Feld-Positionen nacheinander = Positionstausch.
            </div>
          )}
          {subModal.selPos && penByPos[subModal.selPos] && (
            <div style={{ fontFamily: SANS, fontSize: 12, color: C.red, marginBottom: 8 }}>
              Dieser Slot ist gesperrt – der gewählte Spieler kommt automatisch, sobald die Strafzeit abgelaufen ist.
            </div>
          )}
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {bench.length === 0 ? <Empty>Keine Spieler auf der Bank.</Empty> : bench.map((p) => (
              <button key={p.id} onClick={() => clickBench(p.id)} disabled={!subModal.selPos} style={{
                ...btnBase, padding: "10px 12px", fontSize: 14,
                background: subModal.selPos ? C.blueSoft : "#fff",
                color: subModal.selPos ? C.blueDark : C.sub,
                border: `2px solid ${C.line}`, opacity: subModal.selPos ? 1 : 0.6,
              }}>#{p.number} {p.name}</button>
            ))}
          </div>
        </Modal>
      )}

      {/* -------- Modale des Wurf-/Aktions-Flows -------- */}
      {pending && pending.side === "us" && !pending.pickShooter && !pending.zone && !pending.isThrow && (
        <Modal title={playerLabel(team, pending.playerId)} onClose={() => setPending(null)}>
          <ActionGrid onPick={(act) => {
            if (act === "throw") { setPending((p) => ({ ...p, isThrow: true })); return; }
            if (act === "assist") { setPending({ side: "us", assistId: pending.playerId, pickShooter: true }); return; }
            if (act.startsWith("pen:")) { addPenalty(act.slice(4), pending.playerId); return; }
            addAction({ type: act, playerId: pending.playerId }); setPending(null);
          }} />
        </Modal>
      )}
      {pending && pending.pickShooter && (
        <Modal title={`Assist: ${playerName(team, pending.assistId)} – wer wirft?`} onClose={() => setPending(null)}>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 8 }}>
            {roster.filter((p) => p.id !== pending.assistId && (!hasLineup || onFieldIds.has(p.id))).map((p) => (
              <button key={p.id}
                onClick={() => setPending({ side: "us", playerId: p.id, assistId: pending.assistId, isThrow: true })}
                style={{
                  ...btnBase, padding: "14px 10px", textAlign: "left", background: "#fff",
                  border: `2px solid ${C.line}`, display: "flex", alignItems: "center", gap: 10, color: C.ink,
                }}>
                <span style={{
                  fontFamily: MONO, fontWeight: 800, fontSize: 15, color: "#fff",
                  background: p.pos === "TW" ? C.orange : C.blue, borderRadius: 8, minWidth: 34, textAlign: "center", padding: "5px 0",
                }}>{p.number}</span>
                <span style={{ fontSize: 14 }}>{p.name}</span>
              </button>
            ))}
          </div>
        </Modal>
      )}
      {pending && ((pending.isThrow && !pending.zone) || (pending.side === "them" && !pending.zone)) && (
        <Modal wide title={pending.side === "us"
          ? `Abwurfzone – ${playerName(team, pending.playerId)}${pending.assistId ? ` (Assist: ${playerName(team, pending.assistId)})` : ""}`
          : "Abwurfzone – Gegner"} onClose={() => setPending(null)}>
          <CourtPicker onPick={(zone) => setPending((p) => ({ ...p, zone }))} />
        </Modal>
      )}
      {pending && pending.zone && !pending.target && (
        <Modal wide title={`Zielzone (${ZONE_LABEL[pending.zone]})`} onClose={() => setPending(null)}>
          <GoalPicker onPick={(target) => {
            if (target === "POST" || target === "WIDE") {
              const result = target === "POST" ? "post" : "wide";
              addAction({
                type: "throw", side: pending.side, playerId: pending.playerId || null,
                zone: pending.zone, target, result,
                assistId: pending.assistId || null,
                keeperId: pending.side === "them" ? keeperNow : null,
              });
              setPending(null);
            } else setPending((p) => ({ ...p, target }));
          }} />
        </Modal>
      )}
      {pending && pending.target && (
        <Modal title={`${TARGET_LABEL[pending.target]} – Ergebnis?`} onClose={() => setPending(null)}>
          <div style={{ display: "flex", gap: 10 }}>
            <Btn kind={pending.side === "us" ? "green" : "danger"} style={{ flex: 1, padding: "22px", fontSize: 18 }}
              onClick={() => finishThrow("goal")}>
              {pending.side === "us" ? "⚽ Tor" : "Gegentor"}
            </Btn>
            <Btn kind={pending.side === "us" ? "soft" : "green"} style={{ flex: 1, padding: "22px", fontSize: 18 }}
              onClick={() => finishThrow("saved")}>
              {pending.side === "us" ? "🧤 Gehalten" : "🧤 Parade"}
            </Btn>
          </div>
        </Modal>
      )}

      {/* -------- Live-Statistik (Overlay, Timer läuft weiter) -------- */}
      {showStats && (
        <Modal wide title={`Statistik – ${sc.us}:${sc.them} (${fmtClock(sec)})`} onClose={() => setShowStats(false)}>
          <div style={{ display: "grid", gap: 14 }}>
            <HeatmapSection team={team} games={[game]} />
            <StatsTables team={team} agg={aggregate(team, [game], { endSec: sec })} />
          </div>
        </Modal>
      )}
    </div>
  );
}

function ActionGrid({ onPick }) {
  const big = { padding: "18px 10px", fontSize: 16 };
  return (
    <div style={{ display: "grid", gap: 8 }}>
      <Btn kind="accent" style={big} onClick={() => onPick("throw")}>🤾 Wurf</Btn>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
        <Btn kind="soft" onClick={() => onPick("assist")}>Assist</Btn>
        <Btn kind="soft" onClick={() => onPick("steal")}>Steal</Btn>
        <Btn kind="soft" onClick={() => onPick("block")}>Block</Btn>
        <Btn kind="soft" onClick={() => onPick("tf")}>Techn. Fehler</Btn>
        <Btn kind="soft" onClick={() => onPick("m7won")}>7m geholt</Btn>
        <Btn kind="soft" onClick={() => onPick("m7caused")}>7m verursacht</Btn>
      </div>
      <SectionH>Strafen</SectionH>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 8 }}>
        <Btn style={{ background: C.redSoft, color: C.red }} onClick={() => onPick("pen:p2")}>2 min</Btn>
        <Btn style={{ background: C.yellowSoft, color: "#8a6a10" }} onClick={() => onPick("pen:yellow")}>Gelb</Btn>
        <Btn style={{ background: C.red, color: "#fff" }} onClick={() => onPick("pen:red")}>Rot</Btn>
        <Btn style={{ background: C.blueDark, color: "#fff" }} onClick={() => onPick("pen:blue")}>Blau</Btn>
      </div>
    </div>
  );
}

/* ---------- Nachbereitung ---------- */
function ReviewScreen({ data, update, go, teamId, gameId }) {
  const team = data.teams.find((t) => t.id === teamId);
  const game = data.games.find((g) => g.id === gameId);
  const [editId, setEditId] = useState(null);
  const [verlaufOpen, setVerlaufOpen] = useState(true); // Spielverlauf ein-/ausklappen (reiner UI-Zustand)
  if (!team || !game) return <Empty>Spiel nicht gefunden.</Empty>;
  const sc = computeScore(game);
  const scoreStr = game.home ? `${sc.us}:${sc.them}` : `${sc.them}:${sc.us}`;

  // Aktionen + Wechsel gemeinsam, laufender Spielstand pro Toraktion
  const feed = [
    ...game.actions.map((a) => ({ kind: "action", e: a })),
    ...(game.subs || []).map((s) => ({ kind: "sub", e: s })),
  ].sort((x, y) => (x.e.sec || 0) - (y.e.sec || 0));
  let u = 0, t = 0;
  const rows = feed.map((r) => {
    const a = r.e;
    let score = null;
    if (r.kind === "action" && a.type === "throw" && a.result === "goal") {
      a.side === "us" ? u++ : t++;
      score = `${u}:${t}`;
    }
    return { ...r, score };
  });

  const mut = (fn) => update((d) => { const g = d.games.find((x) => x.id === gameId); fn(g); });
  const gStats = aggregate(team, [game]);

  return (
    <div>
      <PageTitle
        title={game.home ? `${team.name} – ${game.opponent}` : `${game.opponent} – ${team.name}`}
        sub={`${fmtDate(game.date)} · Endstand ${scoreStr}${game.test ? " · Testspiel (zählt nicht zur Saison)" : ""}`}
        onBack={() => go({ name: "team", teamId, tab: "spiele" })}
        right={
          <div style={{ display: "flex", gap: 6 }}>
            <Btn small kind={game.test ? "accent" : "ghost"}
              onClick={() => mut((g) => { g.test = !g.test; })}>
              {game.test ? "Testspiel ✓" : "Als Testspiel markieren"}
            </Btn>
            {game.status === "live" && <Btn small kind="accent" onClick={() => go({ name: "live", teamId, gameId })}>▶ Live weiter</Btn>}
          </div>
        }
      />
      <Card style={{ marginBottom: 14 }}>
        <button onClick={() => setVerlaufOpen((o) => !o)} style={{
          ...btnBase, padding: "4px 8px", background: "transparent", color: C.ink,
          display: "flex", alignItems: "center", gap: 8, width: "100%",
          justifyContent: "flex-start", marginBottom: verlaufOpen ? 10 : 0,
        }}>
          <span style={{ fontSize: 12, color: C.sub }}>{verlaufOpen ? "▾" : "▸"}</span>
          <span style={{ fontFamily: SANS, fontSize: 15, fontWeight: 800, color: C.ink }}>
            Spielverlauf{!verlaufOpen ? ` (${rows.length})` : ""}
          </span>
        </button>
        {verlaufOpen && (
          <>
            {rows.length === 0 && <Empty>Keine Aktionen erfasst.</Empty>}
            {rows.map(({ kind, e: a, score }) => (
          <div key={a.id} style={{ borderBottom: `1px solid ${C.line}`, padding: "8px 0" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{
                fontFamily: MONO, fontSize: 11, fontWeight: 700, color: C.sub,
                background: C.bg, borderRadius: 6, padding: "3px 6px",
              }}>HZ{a.half || 1}</span>
              <span style={{
                flex: 1, fontFamily: SANS, fontSize: 14,
                color: kind === "sub" ? C.sub : C.ink, fontStyle: kind === "sub" ? "italic" : "normal",
              }}>
                {kind === "action" ? actionText(team, a) : subText(team, a)}
                {score && <b style={{ color: C.green, marginLeft: 6 }}>({score})</b>}
              </span>
              {(kind === "action" || !a.srcAction) && (
                <Btn small kind="ghost" onClick={() => setEditId(editId === a.id ? null : a.id)}>✎</Btn>
              )}
              {kind === "action" && (
                <ConfirmBtn onConfirm={() => mut((g) => { removeActionCascade(g, a.id); })} />
              )}
              {kind === "sub" && !a.srcAction && (
                <ConfirmBtn onConfirm={() => mut((g) => { g.subs = g.subs.filter((x) => x.id !== a.id); })} />
              )}
              {kind === "sub" && a.srcAction && (
                <span style={{ fontFamily: SANS, fontSize: 11, color: C.sub }}>autom. (über Strafe löschen)</span>
              )}
            </div>
            {editId === a.id && kind === "action" && (
              <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap", background: C.bg, borderRadius: 10, padding: 10 }}>
                {a.side !== "them" && a.playerId !== undefined && (
                  <select style={{ ...inputStyle, width: "auto", flex: 1, minWidth: 150 }} value={a.playerId || ""}
                    onChange={(e) => mut((g) => { g.actions.find((x) => x.id === a.id).playerId = e.target.value; })}>
                    {team.players.map((p) => <option key={p.id} value={p.id}>#{p.number} {p.name}</option>)}
                  </select>
                )}
                {a.type === "throw" && (
                  <>
                    <select style={{ ...inputStyle, width: "auto" }} value={a.result}
                      onChange={(e) => mut((g) => { g.actions.find((x) => x.id === a.id).result = e.target.value; })}>
                      {Object.entries(RESULT_LABEL).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
                    </select>
                    <select style={{ ...inputStyle, width: "auto" }} value={a.zone}
                      onChange={(e) => mut((g) => { g.actions.find((x) => x.id === a.id).zone = e.target.value; })}>
                      {Object.entries(ZONE_LABEL).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
                    </select>
                    {a.side === "us" && (
                      <select style={{ ...inputStyle, width: "auto" }} value={a.assistId || ""}
                        onChange={(e) => mut((g) => { g.actions.find((x) => x.id === a.id).assistId = e.target.value || null; })}>
                        <option value="">– kein Assist –</option>
                        {team.players.filter((p) => p.id !== a.playerId).map((p) => (
                          <option key={p.id} value={p.id}>Assist: #{p.number} {p.name}</option>
                        ))}
                      </select>
                    )}
                  </>
                )}
                {a.type === "penalty" && (
                  <select style={{ ...inputStyle, width: "auto" }} value={a.kind}
                    onChange={(e) => mut((g) => { g.actions.find((x) => x.id === a.id).kind = e.target.value; })}>
                    {Object.entries(PENALTY_LABEL).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
                  </select>
                )}
                {a.type === "oppPenalty" && (
                  <>
                    <span style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: C.sub, alignSelf: "center" }}>Spielminute:</span>
                    <input type="number" min="0" max="70" style={{ ...inputStyle, width: 90 }}
                      value={Math.floor((a.sec || 0) / 60)}
                      onChange={(e) => {
                        const min = Math.max(0, Math.min(70, parseInt(e.target.value) || 0));
                        mut((g) => { const x = g.actions.find((y) => y.id === a.id); if (x) x.sec = min * 60; });
                      }} />
                  </>
                )}
              </div>
            )}
            {editId === a.id && kind === "sub" && !a.srcAction && (
              <div style={{ display: "flex", gap: 8, marginTop: 8, alignItems: "center", flexWrap: "wrap", background: C.bg, borderRadius: 10, padding: 10 }}>
                <span style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: C.sub }}>Spielminute:</span>
                <input type="number" min="0" max="70" style={{ ...inputStyle, width: 90 }}
                  value={Math.floor((a.sec || 0) / 60)}
                  onChange={(e) => {
                    const min = Math.max(0, Math.min(70, parseInt(e.target.value) || 0));
                    mut((g) => { const s = g.subs.find((x) => x.id === a.id); if (s) s.sec = min * 60; });
                  }} />
              </div>
            )}
          </div>
        ))}
          </>
        )}
      </Card>
      <div style={{ marginBottom: 14 }}>
        <HeatmapSection team={team} games={[game]} />
      </div>
      <StatsTables team={team} agg={gStats}
        onPlayer={(pid) => go({
          name: "player", teamId, playerId: pid,
          init: { sel: gameId },
          back: { name: "review", teamId, gameId },
        })} />
    </div>
  );
}

/* ---------- Heatmap (Wurfleistung / Paradenquote) ---------- */
const heatFill = (q) => (q == null ? C.court : `hsl(${Math.round(q * 125)}, 62%, 82%)`);
const heatText = (q) => (q == null ? C.sub : `hsl(${Math.round(q * 125)}, 70%, 24%)`);

function HeatChip({ label, d, countOnly, onClick, selected, dimmed }) {
  const q = d && d.n ? d.k / d.n : null;
  const wrap = {
    flex: 1, textAlign: "center", padding: "10px 6px", borderRadius: 12,
    border: `${selected ? 3 : 1}px solid ${selected ? C.ink : C.line}`,
    opacity: dimmed ? 0.4 : 1, cursor: onClick ? "pointer" : "default",
    transition: "opacity 0.15s, border 0.15s",
  };
  if (countOnly) {
    return (
      <div onClick={onClick} style={{ ...wrap, background: C.bg }}>
        <div style={{ fontFamily: SANS, fontSize: 12, fontWeight: 800, color: C.sub }}>{label}</div>
        <div style={{ fontFamily: MONO, fontSize: 15, fontWeight: 800, color: C.ink }}>
          {d && d.n ? `${d.n}×` : "–"}
        </div>
      </div>
    );
  }
  return (
    <div onClick={onClick} style={{ ...wrap, background: heatFill(q) }}>
      <div style={{ fontFamily: SANS, fontSize: 12, fontWeight: 800, color: C.sub }}>{label}</div>
      <div style={{ fontFamily: MONO, fontSize: 15, fontWeight: 800, color: heatText(q) }}>
        {q == null ? "–" : `${d.k}/${d.n} · ${Math.round(q * 100)}%`}
      </div>
    </div>
  );
}

function CourtHeat({ data, selectedZone, onZoneClick }) {
  const clickable = !!onZoneClick;
  return (
    <div>
      <svg viewBox="-4 -8 408 306" style={{ width: "100%", display: "block", touchAction: "manipulation" }}>
        {Object.keys(ZONE_PATHS).map((z) => {
          const d = data[z];
          const q = d && d.n ? d.k / d.n : null;
          const isSel = selectedZone === z;
          const dimmed = !!selectedZone && !isSel;
          return (
            <path key={z} d={ZONE_PATHS[z]} fill={heatFill(q)}
              stroke={isSel ? C.ink : "#fff"} strokeWidth={isSel ? 4 : 2}
              opacity={dimmed ? 0.35 : 1} style={{ cursor: clickable ? "pointer" : "default" }}
              onClick={clickable ? () => onZoneClick(z) : undefined}>
              <title>{ZONE_LABEL[z]}</title>
            </path>
          );
        })}
        <path d={path(seg(120, 50, 350), false)} fill="none" stroke={C.blue} strokeWidth="2.5" pointerEvents="none" />
        <path d={path(seg(180, 0, 400), false)} fill="none" stroke={C.blue} strokeWidth="2"
          strokeDasharray="12 9" pointerEvents="none" />
        <line x1="0" y1="0" x2="400" y2="0" stroke={C.sub} strokeWidth="2" pointerEvents="none" />
        <line x1="170" y1="-3" x2="230" y2="-3" stroke={C.red} strokeWidth="7" pointerEvents="none" />
        {Object.entries(ZONE_LABEL_POS).map(([z, [x, y]]) => {
          const d = data[z];
          const q = d && d.n ? d.k / d.n : null;
          const dimmed = !!selectedZone && selectedZone !== z;
          return (
            <g key={z} pointerEvents="none" opacity={dimmed ? 0.35 : 1}>
              <text x={x} y={y - 8} textAnchor="middle"
                style={{ fontFamily: SANS, fontSize: 10, fontWeight: 700, fill: C.sub }}>{ZONE_SHORT[z]}</text>
              <text x={x} y={y + 6} textAnchor="middle"
                style={{ fontFamily: MONO, fontSize: 13, fontWeight: 800, fill: heatText(q) }}>
                {q == null ? "–" : `${d.k}/${d.n}`}
              </text>
              {q != null && (
                <text x={x} y={y + 19} textAnchor="middle"
                  style={{ fontFamily: MONO, fontSize: 11, fontWeight: 700, fill: heatText(q) }}>
                  {Math.round(q * 100)}%
                </text>
              )}
            </g>
          );
        })}
      </svg>
      <div style={{ fontFamily: SANS, fontSize: 12, fontWeight: 800, color: C.sub, margin: "12px 0 6px" }}>
        Standardsituationen (ohne Feldzone)
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        {["SIEBEN_M", "KONTER", "FREIWURF"].map((z) => (
          <HeatChip key={z} label={ZONE_LABEL[z]} d={data[z]}
            onClick={clickable ? () => onZoneClick(z) : undefined}
            selected={selectedZone === z}
            dimmed={!!selectedZone && selectedZone !== z} />
        ))}
      </div>
    </div>
  );
}

/* Metriken für die Zielzonen-Heatmap. Die Abwurfzonen-Heatmap bleibt bewusst
   immer bei "x/y · %" – umgeschaltet wird nur die Zielverteilung. */
const GOAL_METRIC_OPTIONS = [
  { id: "shots", label: () => "Würfe" },
  { id: "goals", label: (isKeeper) => (isKeeper ? "Paraden" : "Tore") },
  { id: "quote", label: () => "Quote" },
];

function GoalHeat({ data, metricMode = "quote" }) {
  const cells = [];
  for (let r = 0; r < 3; r++)
    for (let c = 0; c < 3; c++)
      cells.push({ id: `t${r * 3 + c + 1}`, x: 46 + c * 96, y: 46 + r * 64 });
  return (
    <div>
      <svg viewBox="0 0 380 268" style={{ width: "100%", display: "block" }}>
        <defs>
          <pattern id="rwH2" width="24" height="14" patternUnits="userSpaceOnUse">
            <rect width="24" height="14" fill="#fff" />
            <rect width="12" height="14" fill={C.red} />
          </pattern>
          <pattern id="rwV2" width="14" height="24" patternUnits="userSpaceOnUse">
            <rect width="14" height="24" fill="#fff" />
            <rect width="14" height="12" fill={C.red} />
          </pattern>
        </defs>
        <rect x="0" y="0" width="380" height="268" rx="14" fill="#F2ECE4" />
        {cells.map((cl) => {
          const d = data[cl.id];
          const q = d && d.n ? d.k / d.n : null;
          const hasN = !!(d && d.n);
          const isCount = metricMode !== "quote";
          const fill = isCount ? (hasN ? C.bg : "#fff") : heatFill(q);
          const textColor = isCount ? (hasN ? C.ink : C.sub) : heatText(q);
          const mainText = !hasN ? "–"
            : metricMode === "shots" ? `${d.n}×`
            : metricMode === "goals" ? `${d.k}×`
            : `${d.k}/${d.n}`;
          const subText = metricMode === "quote" && q != null ? `${Math.round(q * 100)}%` : null;
          return (
            <g key={cl.id}>
              <rect x={cl.x} y={cl.y} width="96" height="64" fill={fill} stroke="#fff" strokeWidth="2">
                <title>{TARGET_LABEL[cl.id]}</title>
              </rect>
              <text x={cl.x + 48} y={cl.y + 30} textAnchor="middle" pointerEvents="none"
                style={{ fontFamily: MONO, fontSize: 14, fontWeight: 800, fill: textColor }}>
                {mainText}
              </text>
              {subText != null && (
                <text x={cl.x + 48} y={cl.y + 46} textAnchor="middle" pointerEvents="none"
                  style={{ fontFamily: MONO, fontSize: 11, fontWeight: 700, fill: textColor }}>
                  {subText}
                </text>
              )}
            </g>
          );
        })}
        <rect x="32" y="32" width="316" height="14" fill="url(#rwH2)" stroke={C.sub} strokeWidth="1" />
        <rect x="32" y="46" width="14" height="192" fill="url(#rwV2)" stroke={C.sub} strokeWidth="1" />
        <rect x="334" y="46" width="14" height="192" fill="url(#rwV2)" stroke={C.sub} strokeWidth="1" />
        <line x1="20" y1="238" x2="360" y2="238" stroke={C.sub} strokeWidth="2.5" />
      </svg>
      {(data.POST || data.WIDE) && (
        <>
          <div style={{ fontFamily: SANS, fontSize: 12, fontWeight: 800, color: C.sub, margin: "12px 0 6px" }}>
            Würfe neben das Tor
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <HeatChip label={TARGET_LABEL.POST} d={data.POST} countOnly />
            <HeatChip label={TARGET_LABEL.WIDE} d={data.WIDE} countOnly />
          </div>
        </>
      )}
    </div>
  );
}

/* Wurf-/Paraden-Verteilung nach Abwurfzone und Zielzone.
   Gemeinsame Datenquelle für Heatmap-Ansicht UND Excel-Export (Regel 1).
   zoneTargets[zone][target] = {n,k} ist die Kreuzauswertung für den
   Abwurfzone→Zielzone-Drilldown; der Export nutzt sie nicht (rein additiv). */
export function computeHeat(games, selId, isKeeper) {
  const zones = {}, targets = {}, zoneTargets = {};
  let n = 0, k = 0;
  const bump = (map, key, hit) => {
    if (!key) return;
    const e = (map[key] ||= { n: 0, k: 0 });
    e.n++; if (hit) e.k++;
  };
  const bumpCross = (zoneKey, targetKey, hit) => {
    if (!zoneKey || !targetKey) return;
    const zmap = (zoneTargets[zoneKey] ||= {});
    const e = (zmap[targetKey] ||= { n: 0, k: 0 });
    e.n++; if (hit) e.k++;
  };
  for (const g of games) {
    for (const a of g.actions || []) {
      if (a.type !== "throw") continue;
      if (!isKeeper) {
        if (a.side !== "us") continue;
        if (selId !== "team" && a.playerId !== selId) continue;
        const hit = a.result === "goal";
        bump(zones, a.zone, hit); bump(targets, a.target, hit);
        bumpCross(a.zone, a.target, hit);
        n++; if (hit) k++;
      } else {
        if (a.side !== "them" || a.keeperId !== selId) continue;
        if (a.result !== "goal" && a.result !== "saved") continue; // Pfosten/vorbei zählt nicht für den TW
        const save = a.result === "saved";
        bump(zones, a.zone, save); bump(targets, a.target, save);
        bumpCross(a.zone, a.target, save);
        n++; if (save) k++;
      }
    }
  }
  return { zones, targets, zoneTargets, total: { n, k } };
}

/* Abwurfzone→Zielzone-Drilldown: Klick auf eine Abwurfzone (Feldzonen und
   Standardsituationen) hebt sie hervor und filtert die Zielzonen-Heatmap auf
   die Würfe aus genau dieser Zone. Gemeinsam genutzt von HeatmapSection und
   der kompakten Spielerdetail-Variante. */
function ThrowZoneDrilldown({ zones, targets, zoneTargets, isKeeper }) {
  const [selZone, setSelZone] = useState(null);
  const [metricMode, setMetricMode] = useState("quote");
  const shownTargets = selZone ? ((zoneTargets || {})[selZone] || {}) : targets;
  const selCount = selZone ? (zones[selZone]?.n || 0) : null;
  return (
    <div style={{ display: "grid", gap: 8 }}>
      <div>
        <div style={{ fontFamily: SANS, fontSize: 14, fontWeight: 800, color: C.ink, marginBottom: 6 }}>
          {isKeeper ? "Nach Abwurfzone des Gegners" : "Nach Abwurfzone"}
          <span style={{ fontWeight: 400, color: C.sub, fontSize: 12, marginLeft: 6 }}>
            (Zone antippen für die Zielverteilung)
          </span>
        </div>
        <CourtHeat data={zones} selectedZone={selZone}
          onZoneClick={(z) => setSelZone((cur) => (cur === z ? null : z))} />
      </div>
      <div style={{ borderTop: `2px solid ${C.line}`, paddingTop: 14, marginTop: 10 }}>
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 10, flexWrap: "wrap", marginBottom: 6 }}>
          <div>
            <div style={{ fontFamily: SANS, fontSize: 14, fontWeight: 800, color: C.ink }}>
              {selZone ? `Zielverteilung aus ${ZONE_LABEL[selZone]}` : "Nach Zielzone im Tor"}
            </div>
            {selZone && (
              <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, marginTop: 4, display: "flex", alignItems: "center", gap: 8 }}>
                <span>{selCount} {selCount === 1 ? "Wurf" : "Würfe"} aus dieser Zone</span>
                <Btn kind="ghost" small onClick={() => setSelZone(null)} style={{ padding: "3px 9px", fontSize: 11 }}>
                  ✕ Filter aufheben
                </Btn>
              </div>
            )}
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            {GOAL_METRIC_OPTIONS.map((opt) => (
              <button key={opt.id} onClick={() => setMetricMode(opt.id)} style={{
                ...btnBase, padding: "6px 10px", fontSize: 12,
                background: metricMode === opt.id ? C.navy : "#fff",
                color: metricMode === opt.id ? "#fff" : C.sub,
                border: `2px solid ${metricMode === opt.id ? C.navy : C.line}`,
              }}>{opt.label(isKeeper)}</button>
            ))}
          </div>
        </div>
        <GoalHeat data={shownTargets} metricMode={metricMode} />
      </div>
    </div>
  );
}

/* Kompakte Wurfanalyse für die Spielerdetailansicht: gleicher Drilldown, aber
   ohne Dropdown – fest auf diesen Spieler bezogen, nutzt den dort vorhandenen
   Zeitraum-/Spiele-Filter. */
function PlayerThrowZoneCard({ games, player }) {
  const isKeeper = player.pos === "TW";
  const { zones, targets, zoneTargets, total } = useMemo(
    () => computeHeat(games, player.id, isKeeper),
    [games, player.id, isKeeper]
  );
  if (total.n === 0) return null; // Feld-/Torhüterstatistik zeigt bereits den Hinweis
  return (
    <Card>
      <SectionH>Wurfanalyse – {isKeeper ? "Paradenquote" : "Torquote"}</SectionH>
      <ThrowZoneDrilldown zones={zones} targets={targets} zoneTargets={zoneTargets} isKeeper={isKeeper} />
    </Card>
  );
}

function HeatmapSection({ team, games }) {
  const [selId, setSelId] = useState("team");
  const player = team.players.find((p) => p.id === selId);
  const isKeeper = player?.pos === "TW";
  const keepers = team.players.filter((p) => p.pos === "TW");
  const field = team.players.filter((p) => p.pos === "F");

  const { zones, targets, zoneTargets, total } = useMemo(
    () => computeHeat(games, selId, isKeeper),
    [team, games, selId, isKeeper]
  );

  const metric = isKeeper ? "Paradenquote" : "Torquote";
  const totalQ = total.n ? Math.round((total.k / total.n) * 100) : null;

  return (
    <Card>
      <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginBottom: 12 }}>
        <div style={{ flex: 1, minWidth: 220 }}>
          <SectionH>Heatmap – {metric}</SectionH>
          <select style={inputStyle} value={selId} onChange={(e) => setSelId(e.target.value)}>
            <option value="team">Gesamtes Team (Würfe)</option>
            {field.length > 0 && (
              <optgroup label="Feldspieler">
                {field.map((p) => <option key={p.id} value={p.id}>#{p.number} {p.name}</option>)}
              </optgroup>
            )}
            {keepers.length > 0 && (
              <optgroup label="Torhüter (Paradenquote)">
                {keepers.map((p) => <option key={p.id} value={p.id}>#{p.number} {p.name}</option>)}
              </optgroup>
            )}
          </select>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontFamily: SANS, fontSize: 12, fontWeight: 800, color: C.sub }}>Gesamt</div>
          <div style={{ fontFamily: MONO, fontSize: 22, fontWeight: 800, color: heatText(total.n ? total.k / total.n : null) }}>
            {totalQ == null ? "–" : `${total.k}/${total.n} · ${totalQ}%`}
          </div>
        </div>
      </div>
      {total.n === 0 ? (
        <Empty>{isKeeper ? "Für diesen Torhüter sind noch keine gegnerischen Würfe erfasst." : "Noch keine Würfe im gewählten Zeitraum."}</Empty>
      ) : (
        <div style={{ display: "grid", gap: 8 }}>
          <ThrowZoneDrilldown zones={zones} targets={targets} zoneTargets={zoneTargets} isKeeper={isKeeper} />
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontFamily: SANS, fontSize: 12, color: C.sub, fontWeight: 700 }}>{metric} (Abwurfzone):</span>
            {[0, 0.25, 0.5, 0.75, 1].map((q) => (
              <span key={q} style={{
                fontFamily: MONO, fontSize: 11, fontWeight: 700, color: heatText(q),
                background: heatFill(q), borderRadius: 6, padding: "3px 8px",
              }}>{Math.round(q * 100)}%</span>
            ))}
          </div>
        </div>
      )}
    </Card>
  );
}

/* ---------- Statistik (Saison / Spiel) ---------- */
/* ---------- Spieler-Detailansicht ---------- */
/* Zeigt alle Stats eines Spielers mit eigenem Auswertungs-Filter.
   Neu: Spielzeit nach Position (Minuten + Anteil an seiner Gesamtspielzeit).
   Die Gesamtspielzeit selbst bleibt bewusst nur in der Statistik-Tabelle sichtbar. */
function PlayerScreen({ data, go, teamId, playerId, init = {}, back }) {
  const team = data.teams.find((t) => t.id === teamId);
  const games = data.games.filter((g) => g.teamId === teamId);
  const player = team?.players.find((p) => p.id === playerId);
  const validInit = init.sel === "all" || init.sel === "tests" || games.some((g) => g.id === init.sel);
  const [sel, setSel] = useState(validInit ? init.sel : "all");
  const [fromMin, setFromMin] = useState(init.fromMin || 0);
  const [toMin, setToMin] = useState(init.toMin == null ? 60 : init.toMin);
  const fs = statsFilterState(games, sel, fromMin, toMin);
  const { source, winOpts } = fs;
  const agg = useMemo(
    () => (team ? aggregate(team, source, winOpts) : { P: {}, K: {}, M: {}, MP: {}, R: {}, minutesTracked: 0, gamesCount: 0 }),
    [team, games, sel, fromMin, toMin]
  );
  if (!team || !player) return <Empty>Spieler nicht gefunden.</Empty>;

  const s = agg.P[playerId];
  const k = agg.K[playerId];
  const mp = agg.MP[playerId] || {};
  const totalSec = Object.values(mp).reduce((a, b) => a + b, 0);
  const rosterGames = agg.R[playerId] || 0;
  const seasonScope = sel === "all" || sel === "tests"; // Ø nur bei Mehrspiel-Auswertung
  const avgMin = rosterGames > 0 ? Math.round(((agg.M[playerId] || 0) / 60 / rosterGames) * 10) / 10 : null;

  const posRows = POSITIONS
    .filter((pos) => mp[pos] > 0)
    .map((pos) => ({
      pos,
      min: Math.round(mp[pos] / 60),
      pct: totalSec > 0 ? Math.round((mp[pos] / totalSec) * 100) : 0,
    }))
    .sort((a, b) => b.min - a.min);

  const tile = (label, value, color) => (
    <div key={label} style={{
      flex: "1 1 90px", minWidth: 90, textAlign: "center", padding: "10px 6px",
      borderRadius: 12, background: C.bg, border: `1px solid ${C.line}`,
    }}>
      <div style={{ fontFamily: SANS, fontSize: 12, fontWeight: 800, color: C.sub }}>{label}</div>
      <div style={{ fontFamily: MONO, fontSize: 16, fontWeight: 800, color: color || C.ink }}>{value}</div>
    </div>
  );
  const quote = s && s.shots ? Math.round((s.goals / s.shots) * 100) + "%" : "–";
  const kQuote = k && (k.saves + k.conceded) ? Math.round((k.saves / (k.saves + k.conceded)) * 100) + "%" : "–";
  const hasAny = s || k || totalSec > 0;

  return (
    <div style={{ display: "grid", gap: 14 }}>
      <PageTitle
        title={`#${player.number} ${player.name}`}
        sub={`${team.name}${player.pos ? ` · ${POS_LABEL[player.pos] || player.pos}` : ""}`}
        onBack={() => go(back || { name: "team", teamId, tab: "stats" })}
      />
      <StatsFilterCard games={games} sel={sel} setSel={setSel}
        fromMin={fromMin} setFromMin={setFromMin} toMin={toMin} setToMin={setToMin} fs={fs} />
      {source.length === 0 ? <Empty>Keine Spiele im gewählten Zeitraum.</Empty> : !hasAny ? (
        <Empty>Keine Daten für {player.name} im gewählten Zeitraum.</Empty>
      ) : (
        <>
          <Card>
            <SectionH>Spielzeit nach Position</SectionH>
            {posRows.length === 0 ? (
              <div style={{ fontFamily: SANS, fontSize: 13, color: C.sub }}>
                Keine Spielzeit-Daten im gewählten Zeitraum (erfasste Startaufstellung erforderlich).
              </div>
            ) : posRows.map((r) => (
              <div key={r.pos} style={{ display: "flex", alignItems: "center", gap: 10, padding: "7px 0", borderBottom: `1px solid ${C.line}` }}>
                <span style={{
                  fontFamily: MONO, fontSize: 12, fontWeight: 800, color: C.blueDark,
                  background: C.blueSoft, borderRadius: 6, padding: "3px 7px", width: 34, textAlign: "center",
                }}>{r.pos}</span>
                <span style={{ flex: 1, fontFamily: SANS, fontSize: 14, fontWeight: 700, color: C.ink }}>
                  {POS_LABEL[r.pos]}
                </span>
                <span style={{ fontFamily: MONO, fontSize: 14, color: C.ink }}>
                  {r.min}&thinsp;Min. <span style={{ color: C.sub }}>({r.pct}%)</span>
                </span>
              </div>
            ))}
            {seasonScope && posRows.length > 0 && (
              <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 0 2px" }}>
                <span style={{ flex: 1, fontFamily: SANS, fontSize: 13, fontWeight: 800, color: C.sub }}>
                  Ø Min. pro Spiel <span style={{ fontWeight: 400 }}>(über {rosterGames} Spiel{rosterGames === 1 ? "" : "e"} im Kader)</span>
                </span>
                <span style={{ fontFamily: MONO, fontSize: 15, fontWeight: 800, color: C.ink }}>
                  {avgMin == null ? "–" : avgMin}
                </span>
              </div>
            )}
            {agg.minutesTracked < agg.gamesCount && (
              <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, marginTop: 8 }}>
                Spielzeit wird nur aus Spielen mit erfasster Startaufstellung berechnet
                ({agg.minutesTracked} von {agg.gamesCount} Spielen). Kader-Spiele ohne Startaufstellung
                zählen im Ø mit 0 Minuten.
              </div>
            )}
          </Card>
          {s && (
            <Card>
              <SectionH>Feldstatistik</SectionH>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {tile("Tore", s.goals, C.green)}
                {tile("Würfe", s.shots)}
                {tile("Quote", quote)}
                {tile("Assists", s.assist)}
                {tile("Techn. Fehler", s.tf)}
                {tile("Steals", s.steal)}
                {tile("Blocks", s.block)}
                {tile("2 Minuten", s.p2, s.p2 ? C.red : undefined)}
                {tile("7m +", s.m7won)}
                {tile("7m −", s.m7caused)}
                {(s.yellow > 0 || s.red > 0 || s.blue > 0) && tile("Gelb/Rot/Blau", `${s.yellow}/${s.red}/${s.blue}`, s.red || s.blue ? C.red : undefined)}
              </div>
            </Card>
          )}
          {k && (
            <Card>
              <SectionH>Torhüter</SectionH>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {tile("Paraden", k.saves, C.green)}
                {tile("Gegentore", k.conceded)}
                {tile("Quote", kQuote)}
              </div>
            </Card>
          )}
          <PlayerThrowZoneCard games={source} player={player} />
        </>
      )}
    </div>
  );
}

/* ---------- Statistik-Filter (gemeinsam für Statistik-Tab und Spieleransicht) ---------- */
function statsFilterState(games, sel, fromMin, toMin) {
  const finished = games.filter((g) => g.status === "finished" && !g.test);
  const finishedTests = games.filter((g) => g.status === "finished" && g.test);
  const source =
    sel === "all" ? finished :
    sel === "tests" ? finishedTests :
    games.filter((g) => g.id === sel);
  const rangeActive = fromMin > 0 || toMin < 60;
  const rangeValid = fromMin < toMin;
  const fromSec = fromMin * 60;
  const toSec = toMin >= 60 ? Infinity : toMin * 60;
  const winOpts = rangeActive && rangeValid ? { fromSec, toSec } : {};
  return { finished, finishedTests, source, rangeActive, rangeValid, fromSec, toSec, winOpts };
}
function StatsFilterCard({ games, sel, setSel, fromMin, setFromMin, toMin, setToMin, fs }) {
  const { finished, finishedTests, source, rangeActive, rangeValid } = fs;
  const minInput = { ...inputStyle, width: 90, textAlign: "center", fontFamily: MONO };
  const clampMin = (v) => Math.max(0, Math.min(60, parseInt(v) || 0));
  return (
    <Card>
      <SectionH>Auswertung über</SectionH>
      <select style={inputStyle} value={sel} onChange={(e) => setSel(e.target.value)}>
        <option value="all">Saison – alle beendeten Pflichtspiele ({finished.length})</option>
        {finishedTests.length > 0 && (
          <option value="tests">Nur Testspiele ({finishedTests.length})</option>
        )}
        {games.map((g) => (
          <option key={g.id} value={g.id}>{fmtDate(g.date)} – {g.opponent}{g.test ? " (Test)" : ""}</option>
        ))}
      </select>
      {sel === "all" && finishedTests.length > 0 && (
        <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, marginTop: 8 }}>
          Testspiele sind hier ausgenommen und einzeln bzw. gesammelt auswählbar.
        </div>
      )}
      <div style={{ display: "flex", gap: 10, alignItems: "center", marginTop: 12, flexWrap: "wrap" }}>
        <span style={{ fontFamily: SANS, fontSize: 13, fontWeight: 800, color: C.sub }}>Zeitfenster:</span>
        <label style={{ display: "flex", alignItems: "center", gap: 6, fontFamily: SANS, fontSize: 13, color: C.sub }}>
          von Minute
          <input type="number" min="0" max="59" style={minInput} value={fromMin}
            onChange={(e) => setFromMin(clampMin(e.target.value))} />
        </label>
        <label style={{ display: "flex", alignItems: "center", gap: 6, fontFamily: SANS, fontSize: 13, color: C.sub }}>
          bis Minute
          <input type="number" min="1" max="60" style={minInput} value={toMin}
            onChange={(e) => setToMin(clampMin(e.target.value))} />
        </label>
        {rangeActive && rangeValid && (
          <Btn small kind="ghost" onClick={() => { setFromMin(0); setToMin(60); }}>Zurücksetzen</Btn>
        )}
      </div>
      {!rangeValid && (
        <div style={{ fontFamily: SANS, fontSize: 12, color: C.red, marginTop: 8 }}>
          „von" muss kleiner als „bis" sein – es wird das ganze Spiel ausgewertet.
        </div>
      )}
      {rangeActive && rangeValid && (
        <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, marginTop: 8 }}>
          Ausgewertet wird Minute {fromMin + 1}–{toMin}{toMin >= 60 ? " (bis Spielende)" : ""}
          {source.length > 1 ? " – pro Spiel angewendet und aufsummiert." : "."}
        </div>
      )}
    </Card>
  );
}

/* ============================================================
   Excel-Export (nur Statistik-Tab)
   Regeln: siehe export-regeln.md –
   1. keine eigene Statistik-Logik (nutzt aggregate/buildStatRows/computeHeat),
   2. Spalten kommen aus FIELD_STAT_COLS / KEEPER_STAT_COLS,
   3. jedes Blatt ist eine eigene build*Sheet-Funktion.
   ============================================================ */
const XL_INK = "16202E", XL_SUB = "5C6B7E", XL_BLUE = "2E5EAA", XL_NAVY = "1F3F76";
const XL = {
  title: { font: { bold: true, sz: 14, color: { rgb: XL_INK } } },
  sub: { font: { sz: 10, color: { rgb: XL_SUB } } },
  section: { font: { bold: true, sz: 12, color: { rgb: XL_NAVY } } },
  th: {
    font: { bold: true, sz: 11, color: { rgb: "FFFFFF" } },
    fill: { fgColor: { rgb: XL_BLUE } },
    alignment: { horizontal: "right" },
  },
  thLeft: {
    font: { bold: true, sz: 11, color: { rgb: "FFFFFF" } },
    fill: { fgColor: { rgb: XL_BLUE } },
    alignment: { horizontal: "left" },
  },
  note: { font: { italic: true, sz: 10, color: { rgb: XL_SUB } } },
  bold: { font: { bold: true, color: { rgb: XL_INK } } },
  /* Gesamtzeile: fett + dickere Trennlinie oben (wie in der App-Tabelle). */
  total: {
    font: { bold: true, color: { rgb: XL_INK } },
    border: { top: { style: "medium", color: { rgb: XL_INK } } },
    alignment: { horizontal: "right" },
  },
  totalLeft: {
    font: { bold: true, color: { rgb: XL_INK } },
    border: { top: { style: "medium", color: { rgb: XL_INK } } },
    alignment: { horizontal: "left" },
  },
};
const xlText = (v, s) => ({ v: v == null ? "" : String(v), t: "s", ...(s ? { s } : {}) });
const xlNum = (v, s) => ({ v: v == null ? 0 : v, t: "n", ...(s ? { s } : {}) });
/* Quote als echter Excel-Prozentwert (Bruch + Format "0%"), "–" wenn kein Nenner. */
const xlPct = (num, den, s) => (den ? { v: num / den, t: "n", z: "0%", ...(s ? { s } : {}) } : xlText("–", s));
/* Minuten wie in der App: "–" ohne erfasste Startaufstellung, sonst gerundete Minuten. */
const xlMin = (sec) => (sec == null || sec < 0 ? xlText("–") : xlNum(Math.round(sec / 60)));

const sanitizeFileName = (s) => s.replace(/[\\/:*?"<>|]/g, "-").trim();
const sanitizeSheetName = (s, used) => {
  let base = s.replace(/[\\/:*?\[\]]/g, "-").slice(0, 31).trim() || "Blatt";
  let name = base, i = 2;
  while (used.has(name)) name = `${base.slice(0, 28)} ${i++}`;
  used.add(name);
  return name;
};

export function exportFileName(games, sel) {
  const today = fmtDate(todayISO());
  if (sel === "all") return sanitizeFileName(`Saison-Pflichtspiele – Export ${today}`);
  if (sel === "tests") return sanitizeFileName(`Testspiele – Export ${today}`);
  const g = games.find((x) => x.id === sel);
  if (!g) return sanitizeFileName(`Statistik – Export ${today}`);
  return sanitizeFileName(`${g.opponent} – ${fmtDate(g.date)}${g.test ? " (Test)" : ""}`);
}

/* Überschrift + Untertitel je nach Filterzustand (Kontext auf jedem Blatt). */
function exportHeading(team, games, sel, gamesCount, rangeActive, fromMin, toMin) {
  const g = games.find((x) => x.id === sel);
  const title =
    sel === "all" ? "Saison – alle beendeten Pflichtspiele" :
    sel === "tests" ? "Testspiele" :
    g ? `${g.opponent} – ${fmtDate(g.date)}${g.test ? " (Testspiel)" : ""}` : "Statistik";
  const parts = [team.name, `${gamesCount} Spiel${gamesCount === 1 ? "" : "e"}`];
  if (rangeActive) parts.push(`Minute ${fromMin + 1}–${toMin}${toMin >= 60 ? " (bis Spielende)" : ""}`);
  parts.push(`Exportiert am ${fmtDate(todayISO())}`);
  return { title, subtitle: parts.join(" · ") };
}

/* Zonen-/Zielzonen-Tabelle als AoA-Zeilen (gemeinsam für Team- und Spieler-Blätter). */
function heatRowsAoA(map, labelMap, order, countOnlyKeys = []) {
  const out = [];
  for (const key of order) {
    const d = map[key];
    if (!d || !d.n) continue;
    const countOnly = countOnlyKeys.includes(key);
    out.push([
      xlText(labelMap[key] || key),
      xlNum(d.n),
      countOnly ? xlText("–") : xlNum(d.k),
      countOnly ? xlText("–") : xlPct(d.k, d.n),
    ]);
  }
  return out;
}

function buildTeamSheet({ team, agg, rows, kRows, zRows, heading }) {
  const aoa = [];
  aoa.push([xlText(heading.title, XL.title)]);
  aoa.push([xlText(heading.subtitle, XL.sub)]);
  aoa.push([]);
  aoa.push([xlText("Feldspieler", XL.section)]);
  aoa.push([
    xlText("Nr.", XL.thLeft), xlText("Spieler", XL.thLeft),
    ...FIELD_STAT_COLS.map((c) => xlText(c.label, XL.th)),
  ]);
  for (const r of rows) {
    aoa.push([
      xlText(r.num === 999 ? "–" : `#${r.num}`), xlText(r.name, XL.bold),
      ...FIELD_STAT_COLS.map((c) =>
        c.percent ? xlPct(r.goals, r.shots) : c.isMin ? xlMin(r.min) : xlNum(r[c.key])
      ),
    ]);
  }
  if (rows.length > 0) {
    const t = statTotals(rows, FIELD_STAT_COLS);
    aoa.push([
      xlText("–", XL.totalLeft), xlText("Gesamt", XL.totalLeft),
      ...FIELD_STAT_COLS.map((c) =>
        c.percent ? xlPct(t.goals, t.shots, XL.total)
          : c.isMin ? xlText("–", XL.total) : xlNum(t[c.key], XL.total)
      ),
    ]);
  }
  if (agg.minutesTracked < agg.gamesCount) {
    aoa.push([xlText(
      `Spielzeit („Min.") wird nur aus Spielen mit erfasster Startaufstellung berechnet (${agg.minutesTracked} von ${agg.gamesCount} Spielen).`,
      XL.note
    )]);
  }
  if (agg.oppP2 > 0) {
    aoa.push([xlText(`Gegner: ${agg.oppP2} × 2-Minuten-Strafe (Überzahl-Situationen) im gewählten Zeitraum.`, XL.note)]);
  }
  aoa.push([]);
  aoa.push([xlText("Torhüter", XL.section)]);
  aoa.push([
    xlText("Nr.", XL.thLeft), xlText("Torhüter", XL.thLeft),
    ...KEEPER_STAT_COLS.map((c) => xlText(c.label, XL.th)),
  ]);
  for (const r of kRows) {
    aoa.push([
      xlText(r.num === 999 ? "–" : `#${r.num}`), xlText(r.name, XL.bold),
      ...KEEPER_STAT_COLS.map((c) =>
        c.percent ? xlPct(r.saves, r.saves + r.conceded) : c.isMin ? xlMin(r.min) : xlNum(r[c.key])
      ),
    ]);
  }
  if (kRows.length > 0) {
    const t = statTotals(kRows, KEEPER_STAT_COLS);
    aoa.push([
      xlText("–", XL.totalLeft), xlText("Gesamt", XL.totalLeft),
      ...KEEPER_STAT_COLS.map((c) =>
        c.percent ? xlPct(t.saves, t.saves + t.conceded, XL.total)
          : c.isMin ? xlText("–", XL.total) : xlNum(t[c.key], XL.total)
      ),
    ]);
  }
  aoa.push([]);
  aoa.push([xlText("Wurfzonen (eigene Würfe)", XL.section)]);
  aoa.push([xlText("Zone", XL.thLeft), xlText("Würfe", XL.th), xlText("Tore", XL.th), xlText("Quote", XL.th)]);
  for (const r of zRows) {
    aoa.push([xlText(ZONE_LABEL[r.z]), xlNum(r.shots), xlNum(r.goals), xlPct(r.goals, r.shots)]);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 6 }, { wch: 22 }, ...FIELD_STAT_COLS.map(() => ({ wch: 13 }))];
  return ws;
}

function buildPlayerSheet({ player, agg, heat, heading, seasonScope }) {
  const isKeeper = player.pos === "TW";
  const s = agg.P[player.id];
  const k = agg.K[player.id];
  const mp = agg.MP[player.id] || {};
  const totalSec = Object.values(mp).reduce((a, b) => a + b, 0);
  const aoa = [];
  aoa.push([xlText(`#${player.number} ${player.name}`, XL.title)]);
  aoa.push([xlText(
    `${POS_LABEL[player.pos] || (isKeeper ? "Torwart" : "Feldspieler")} · ${heading.title} · ${heading.subtitle}`,
    XL.sub
  )]);
  aoa.push([]);
  if (s) {
    aoa.push([xlText("Feldstatistik", XL.section)]);
    aoa.push([xlText("Kennzahl", XL.thLeft), xlText("Wert", XL.th)]);
    for (const c of FIELD_STAT_COLS) {
      if (c.isMin) continue; // Gesamtspielzeit bewusst nur in der Team-Tabelle (Trennungsregel)
      aoa.push([xlText(c.label), c.percent ? xlPct(s.goals, s.shots) : xlNum(s[c.key])]);
    }
    for (const c of PLAYER_CARD_ROWS) aoa.push([xlText(c.label), xlNum(s[c.key])]);
    aoa.push([]);
  }
  if (k) {
    aoa.push([xlText("Torhüter", XL.section)]);
    aoa.push([xlText("Kennzahl", XL.thLeft), xlText("Wert", XL.th)]);
    for (const c of KEEPER_STAT_COLS) {
      if (c.isMin) continue;
      aoa.push([xlText(c.label), c.percent ? xlPct(k.saves, k.saves + k.conceded) : xlNum(k[c.key])]);
    }
    aoa.push([]);
  }
  const hitLabel = isKeeper ? "Paraden" : "Tore";
  if (heat.total.n > 0) {
    aoa.push([xlText(isKeeper ? "Paraden nach gegnerischer Abwurfzone" : "Würfe nach Abwurfzone", XL.section)]);
    aoa.push([xlText("Zone", XL.thLeft), xlText("Würfe", XL.th), xlText(hitLabel, XL.th), xlText("Quote", XL.th)]);
    aoa.push(...heatRowsAoA(heat.zones, ZONE_LABEL, Object.keys(ZONE_LABEL)));
    aoa.push([]);
    aoa.push([xlText(isKeeper ? "Paraden nach Zielzone im Tor" : "Würfe nach Zielzone im Tor", XL.section)]);
    aoa.push([xlText("Zielzone", XL.thLeft), xlText("Würfe", XL.th), xlText(hitLabel, XL.th), xlText("Quote", XL.th)]);
    aoa.push(...heatRowsAoA(heat.targets, TARGET_LABEL, Object.keys(TARGET_LABEL), ["POST", "WIDE"]));
    aoa.push([]);
  }
  aoa.push([xlText("Spielzeit nach Position", XL.section)]);
  const posRows = POSITIONS.filter((pos) => mp[pos] > 0);
  if (posRows.length === 0) {
    aoa.push([xlText("Keine Spielzeit-Daten im gewählten Zeitraum (erfasste Startaufstellung erforderlich).", XL.note)]);
  } else {
    aoa.push([xlText("Position", XL.thLeft), xlText("Minuten", XL.th), xlText("Anteil", XL.th)]);
    for (const pos of posRows) {
      aoa.push([xlText(`${pos} – ${POS_LABEL[pos]}`), xlNum(Math.round(mp[pos] / 60)), xlPct(mp[pos], totalSec)]);
    }
    if (seasonScope) {
      const rosterGames = agg.R[player.id] || 0;
      const avgMin = rosterGames > 0 ? Math.round(((agg.M[player.id] || 0) / 60 / rosterGames) * 10) / 10 : null;
      aoa.push([
        xlText(`Ø Min. pro Spiel (über ${rosterGames} Spiel${rosterGames === 1 ? "" : "e"} im Kader)`, XL.bold),
        avgMin == null ? xlText("–") : xlNum(avgMin),
      ]);
    }
    if (agg.minutesTracked < agg.gamesCount) {
      aoa.push([xlText(
        `Spielzeit nur aus Spielen mit erfasster Startaufstellung (${agg.minutesTracked} von ${agg.gamesCount} Spielen); Kader-Spiele ohne Startaufstellung zählen im Ø mit 0 Minuten.`,
        XL.note
      )]);
    }
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 40 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];
  return ws;
}

export function buildExportWorkbook({ team, games, sel, fromMin, toMin, rangeActive, agg, rows, kRows, zRows, heatGames }) {
  const heading = exportHeading(team, games, sel, agg.gamesCount, rangeActive, fromMin, toMin);
  const wb = XLSX.utils.book_new();
  const used = new Set();
  XLSX.utils.book_append_sheet(wb, buildTeamSheet({ team, agg, rows, kRows, zRows, heading }), sanitizeSheetName("Team", used));
  // Spieler-Blätter: alle Spieler mit Daten im Filter, nach Trikotnummer sortiert.
  const ids = new Set([...Object.keys(agg.P), ...Object.keys(agg.K)]);
  const players = team.players.filter((p) => ids.has(p.id));
  const sorted = [...players].sort((a, b) => (parseInt(a.number) || 999) - (parseInt(b.number) || 999));
  const seasonScope = sel === "all" || sel === "tests";
  for (const player of sorted) {
    const heat = computeHeat(heatGames, player.id, player.pos === "TW");
    const ws = buildPlayerSheet({ player, agg, heat, heading, seasonScope });
    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(`#${player.number} ${player.name}`, used));
  }
  return { wb, fileName: exportFileName(games, sel) };
}

function exportStatsToExcel(payload) {
  const { wb, fileName } = buildExportWorkbook(payload);
  XLSX.writeFile(wb, `${fileName}.xlsx`);
}

export function StatsTab({ team, games, go, teamId }) {
  const [sel, setSel] = useState("all");
  const [fromMin, setFromMin] = useState(0);  // Zeitfenster: von Minute (0–59)
  const [toMin, setToMin] = useState(60);     // bis Minute (60 = bis Spielende)
  const fs = statsFilterState(games, sel, fromMin, toMin);
  const { source, rangeActive, rangeValid, fromSec, toSec, winOpts } = fs;

  const agg = useMemo(() => aggregate(team, source, winOpts), [team, games, sel, fromMin, toMin]);
  // Für die Heatmap: Aktionen auf das Zeitfenster begrenzen (flache Kopie, Originaldaten bleiben unberührt)
  const heatGames = useMemo(() => {
    if (!rangeActive || !rangeValid) return source;
    return source.map((g) => ({
      ...g,
      actions: (g.actions || []).filter((a) => (a.sec || 0) >= fromSec && (a.sec || 0) < toSec),
    }));
  }, [team, games, sel, fromMin, toMin]);

  const onPlayer = go
    ? (pid) => go({
        name: "player", teamId, playerId: pid,
        init: { sel, fromMin, toMin },
        back: { name: "team", teamId, tab: "stats" },
      })
    : undefined;

  return (
    <div style={{ display: "grid", gap: 14 }}>
      <StatsFilterCard games={games} sel={sel} setSel={setSel}
        fromMin={fromMin} setFromMin={setFromMin} toMin={toMin} setToMin={setToMin} fs={fs} />
      {source.length === 0 ? <Empty>Noch keine Daten. Beende zuerst ein Spiel.</Empty> : (
        <>
          <HeatmapSection team={team} games={heatGames} />
          <StatsTables team={team} agg={agg} onPlayer={onPlayer}
            exportCtx={{ games, sel, fromMin, toMin, rangeActive: rangeActive && rangeValid, heatGames }} />
        </>
      )}
    </div>
  );
}

/* Klickbarer Spaltenkopf: Klick sortiert, erneuter Klick dreht die Richtung um. */
function SortTh({ label, k, sort, setSort, defaultDir = -1, left }) {
  const active = sort.key === k;
  return (
    <th
      onClick={() => setSort(active ? { key: k, dir: -sort.dir } : { key: k, dir: defaultDir })}
      style={{
        fontFamily: SANS, fontSize: 12, fontWeight: 800,
        color: active ? C.blueDark : C.sub,
        textAlign: left ? "left" : "right", padding: "6px 6px",
        whiteSpace: "nowrap", cursor: "pointer", userSelect: "none",
      }}
      title="Zum Sortieren antippen"
    >
      {label}{active ? (sort.dir === 1 ? " ▲" : " ▼") : ""}
    </th>
  );
}
/* Vergleich nach Sortier-Schlüssel, bei Gleichstand Standardreihenfolge (tie). */
const mkCmp = (sort, tie) => (a, b) => {
  const va = a[sort.key], vb = b[sort.key];
  const d = typeof va === "string" || typeof vb === "string"
    ? String(va ?? "").localeCompare(String(vb ?? ""), "de")
    : (va ?? 0) - (vb ?? 0);
  return d !== 0 ? sort.dir * d : tie(a, b);
};

/* Zeilen-Aufbau für Feldspieler-/Torhüter-/Zonen-Tabelle (unsortiert).
   Gemeinsame Datenquelle für App-Tabelle UND Excel-Export (Regel 1). */
export function buildStatRows(team, agg) {
  const { P, K, Z, M = {} } = agg;
  const meta = (id) => {
    const p = team.players.find((x) => x.id === id);
    return { num: parseInt(p?.number) || 999, name: p?.name || "Unbekannt" };
  };
  const rows = Object.entries(P).map(([id, s]) => ({
    id, ...s, ...meta(id),
    quote: s.shots ? Math.round((s.goals / s.shots) * 100) : 0,
    min: M[id] == null ? -1 : M[id],
  }));
  const kRows = Object.entries(K).map(([id, s]) => ({
    id, ...s, ...meta(id),
    quote: s.saves + s.conceded ? Math.round((s.saves / (s.saves + s.conceded)) * 100) : 0,
    min: M[id] == null ? -1 : M[id],
  }));
  const zRows = Object.keys(ZONE_LABEL).map((z) => ({ z, ...(Z[z] || { shots: 0, goals: 0 }) })).filter((r) => r.shots > 0);
  return { rows, kRows, zRows };
}
const cmpFieldDefault = (a, b) => b.goals - a.goals || b.shots - a.shots || a.num - b.num;
const cmpKeeperDefault = (a, b) => b.saves - a.saves || a.num - b.num;

export function StatsTables({ team, agg, onPlayer, exportCtx }) {
  const { M = {}, minutesTracked = 0, gamesCount = 0, oppP2 = 0 } = agg;
  const rowProps = (id) => onPlayer
    ? { onClick: () => onPlayer(id), style: { cursor: "pointer" }, title: "Antippen für Spieler-Details" }
    : {};
  const [sortP, setSortP] = useState({ key: "goals", dir: -1 });
  const [sortK, setSortK] = useState({ key: "saves", dir: -1 });
  const fmtMin = (id) => (M[id] == null ? "–" : String(Math.round(M[id] / 60)));
  const base = buildStatRows(team, agg);
  const rows = [...base.rows].sort(mkCmp(sortP, cmpFieldDefault));
  const kRows = [...base.kRows].sort(mkCmp(sortK, cmpKeeperDefault));
  const zRows = base.zRows;
  const td = { fontFamily: MONO, fontSize: 14, color: C.ink, textAlign: "right", padding: "7px 6px", borderTop: `1px solid ${C.line}` };
  const tdName = { ...td, fontFamily: SANS, fontWeight: 700, textAlign: "left" };
  const tdNum = { ...td, color: C.sub, textAlign: "left", width: 40 };
  /* Gesamtzeile: fett + dickere Trennlinie, fix unten (sortier-unabhängig), nicht klickbar. */
  const tdTotal = { ...td, fontWeight: 800, borderTop: `2px solid ${C.ink}` };
  const tdTotalName = { ...tdTotal, fontFamily: SANS, textAlign: "left" };
  const tdTotalNum = { ...tdTotal, color: C.sub, textAlign: "left", width: 40 };
  const fTot = statTotals(rows, FIELD_STAT_COLS);
  const kTot = statTotals(kRows, KEEPER_STAT_COLS);
  const fTotQuote = fTot.shots ? Math.round((fTot.goals / fTot.shots) * 100) + "%" : "–";
  const kTotQuote = (kTot.saves + kTot.conceded)
    ? Math.round((kTot.saves / (kTot.saves + kTot.conceded)) * 100) + "%" : "–";
  const minNote = minutesTracked < gamesCount && (
    <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, marginTop: 8 }}>
      Spielzeit („Min.") wird nur aus Spielen mit erfasster Startaufstellung berechnet
      ({minutesTracked} von {gamesCount} Spielen).
    </div>
  );
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <Card style={{ overflowX: "auto" }}>
        <SectionH>Torschützen & Feldstatistik</SectionH>
        {rows.length === 0 ? <Empty>Keine Feldaktionen.</Empty> : (
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>
              <SortTh label="Nr." k="num" sort={sortP} setSort={setSortP} defaultDir={1} left />
              <SortTh label="Spieler" k="name" sort={sortP} setSort={setSortP} defaultDir={1} left />
              {FIELD_STAT_COLS.map((c) => (
                <SortTh key={c.key} label={c.short} k={c.key} sort={sortP} setSort={setSortP} />
              ))}
            </tr></thead>
            <tbody>{rows.map((r) => (
              <tr key={r.id} {...rowProps(r.id)}>
                <td style={tdNum}>#{r.num === 999 ? "–" : r.num}</td>
                <td style={{ ...tdName, color: onPlayer ? C.blueDark : C.ink }}>{r.name}</td>
                {FIELD_STAT_COLS.map((c) => {
                  const content = c.percent ? (r.shots ? r.quote + "%" : "–")
                    : c.isMin ? fmtMin(r.id) : r[c.key];
                  const style = c.key === "goals" ? { ...td, fontWeight: 800, color: C.green }
                    : c.key === "p2" ? { ...td, color: r.p2 ? C.red : C.ink } : td;
                  return <td key={c.key} style={style}>{content}</td>;
                })}
              </tr>
            ))}
            <tr>
              <td style={tdTotalNum}>–</td>
              <td style={tdTotalName}>Gesamt</td>
              {FIELD_STAT_COLS.map((c) => (
                <td key={c.key} style={tdTotal}>
                  {c.percent ? fTotQuote : c.isMin ? "–" : fTot[c.key]}
                </td>
              ))}
            </tr>
            </tbody>
          </table>
        )}
        {onPlayer && rows.length > 0 && (
          <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, marginTop: 8 }}>
            Spielerzeile antippen für die Detailansicht (inkl. Spielzeit nach Position).
          </div>
        )}
        {minNote}
        {oppP2 > 0 && (
          <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, marginTop: 8 }}>
            Gegner: <b style={{ color: C.red }}>{oppP2} × 2-Minuten-Strafe</b> (Überzahl-Situationen) im gewählten Zeitraum.
          </div>
        )}
      </Card>
      <Card style={{ overflowX: "auto" }}>
        <SectionH>Torhüter</SectionH>
        {kRows.length === 0 ? <Empty>Keine Torhüter-Daten (werden aus gegnerischen Würfen abgeleitet).</Empty> : (
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>
              <SortTh label="Nr." k="num" sort={sortK} setSort={setSortK} defaultDir={1} left />
              <SortTh label="Torhüter" k="name" sort={sortK} setSort={setSortK} defaultDir={1} left />
              {KEEPER_STAT_COLS.map((c) => (
                <SortTh key={c.key} label={c.short} k={c.key} sort={sortK} setSort={setSortK} />
              ))}
            </tr></thead>
            <tbody>{kRows.map((r) => (
              <tr key={r.id} {...rowProps(r.id)}>
                <td style={tdNum}>#{r.num === 999 ? "–" : r.num}</td>
                <td style={{ ...tdName, color: onPlayer ? C.blueDark : C.ink }}>{r.name}</td>
                {KEEPER_STAT_COLS.map((c) => {
                  const content = c.percent ? r.quote + "%" : c.isMin ? fmtMin(r.id) : r[c.key];
                  const style = c.key === "saves" ? { ...td, fontWeight: 800, color: C.green } : td;
                  return <td key={c.key} style={style}>{content}</td>;
                })}
              </tr>
            ))}
            <tr>
              <td style={tdTotalNum}>–</td>
              <td style={tdTotalName}>Gesamt</td>
              {KEEPER_STAT_COLS.map((c) => (
                <td key={c.key} style={tdTotal}>
                  {c.percent ? kTotQuote : c.isMin ? "–" : kTot[c.key]}
                </td>
              ))}
            </tr>
            </tbody>
          </table>
        )}
      </Card>
      <Card>
        <SectionH>Wurfzonen (eigene Würfe)</SectionH>
        {zRows.length === 0 ? <Empty>Keine Würfe erfasst.</Empty> : zRows.map((r) => {
          const q = Math.round((r.goals / r.shots) * 100);
          return (
            <div key={r.z} style={{ display: "flex", alignItems: "center", gap: 10, padding: "6px 0" }}>
              <span style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: C.ink, width: 150 }}>{ZONE_LABEL[r.z]}</span>
              <div style={{ flex: 1, height: 14, background: C.bg, borderRadius: 7, overflow: "hidden" }}>
                <div style={{ width: `${q}%`, height: "100%", background: q >= 60 ? C.green : q >= 40 ? C.yellow : C.red }} />
              </div>
              <span style={{ fontFamily: MONO, fontSize: 13, color: C.sub, width: 90, textAlign: "right" }}>
                {r.goals}/{r.shots} · {q}%
              </span>
            </div>
          );
        })}
      </Card>
      {exportCtx && (
        <Card>
          <SectionH>Export</SectionH>
          <div style={{ fontFamily: SANS, fontSize: 13, color: C.sub, marginBottom: 10 }}>
            Exportiert die aktuell gewählte Auswertung (inkl. Zeitfenster und Tabellen-Sortierung)
            als Excel-Datei: ein Blatt fürs Team, je ein Blatt pro Spieler.
          </div>
          <Btn kind="soft" onClick={() => exportStatsToExcel({
            team, agg, rows, kRows, zRows, ...exportCtx,
          })}>
            Als Excel exportieren ↓
          </Btn>
        </Card>
      )}
    </div>
  );
}

/* ---------- App-Wurzel ---------- */
/* ---------- Backup: Daten als JSON-Datei sichern / laden ---------- */
function downloadBackup(data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `handball-backup-${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
function readBackup(file, onOk, onErr) {
  const r = new FileReader();
  r.onload = () => {
    try {
      const d = JSON.parse(r.result);
      if (!d || !Array.isArray(d.teams) || !Array.isArray(d.games)) throw new Error("Ungültiges Format");
      onOk(d);
    } catch { onErr(); }
  };
  r.onerror = onErr;
  r.readAsText(file);
}

/* ============================================================
   Cloud-Sync
   Ein Datensatz pro Team im Cloudflare-KV, erreichbar nur über den Worker.
   Live-Tracking bleibt vollständig offline – Sync passiert manuell nach dem
   Spiel (Hochladen) bzw. automatisch beim App-Start (Herunterladen).
   ============================================================ */
const CLOUD_KEY = "handball:cloud-v1";
const EMPTY_CLOUD = { apiBase: "", links: [] };
/* link: { code, slug, label, role, teamId, baseUpdatedAt, lastSyncAt } */

async function loadCloud() {
  try {
    const r = await window.storage.get(CLOUD_KEY);
    if (r && r.value) {
      const c = JSON.parse(r.value);
      return { apiBase: c.apiBase || "", links: Array.isArray(c.links) ? c.links : [] };
    }
  } catch {}
  return clone(EMPTY_CLOUD);
}
async function saveCloud(c) {
  try { await window.storage.set(CLOUD_KEY, JSON.stringify(c)); } catch {}
}

async function cloudFetch(base, path, code, init = {}) {
  if (!base) throw new Error("Es ist noch keine Worker-Adresse hinterlegt.");
  let res;
  try {
    res = await fetch(String(base).replace(/\/+$/, "") + path, {
      ...init,
      headers: { "X-Access-Code": code, ...(init.body ? { "Content-Type": "application/json" } : {}) },
    });
  } catch {
    throw new Error("Keine Verbindung zur Cloud. Bitte Internetverbindung prüfen.");
  }
  let body = null;
  try { body = await res.json(); } catch {}
  if (!res.ok) {
    const e = new Error((body && (body.message || body.error)) || `Cloud-Fehler ${res.status}`);
    e.status = res.status;
    e.updatedAt = body && body.updatedAt;
    throw e;
  }
  return body;
}
const cloudDownload = (base, code) => cloudFetch(base, "/api/team", code);
const cloudUpload = (base, code, payload, baseUpdatedAt, force) =>
  cloudFetch(base, "/api/team/sync", code, {
    method: "POST",
    body: JSON.stringify({ data: payload, baseUpdatedAt: baseUpdatedAt ?? null, force: !!force }),
  });

/* Ein Team samt seiner Spiele aus dem lokalen Datensatz herauslösen */
function teamPayload(data, teamId) {
  const team = data.teams.find((t) => t.id === teamId);
  if (!team) return null;
  return { team: clone(team), games: clone(data.games.filter((g) => g.teamId === teamId)) };
}
/* Cloud-Payload einspielen: ersetzt genau dieses Team und dessen Spiele */
function applyPayload(data, payload) {
  if (!payload || !payload.team || !payload.team.id) return data;
  const n = clone(data);
  const tid = payload.team.id;
  const i = n.teams.findIndex((t) => t.id === tid);
  if (i >= 0) n.teams[i] = clone(payload.team); else n.teams.push(clone(payload.team));
  // Reihenfolge der Spieleliste bleibt erhalten: bekannte Spiele werden an Ort und
  // Stelle ersetzt, in der Cloud gelöschte entfernt, neue hinten angehängt.
  const incoming = new Map((payload.games || []).map((g) => [g.id, clone(g)]));
  const games = [];
  for (const g of n.games) {
    if (g.teamId !== tid) { games.push(g); continue; }
    if (incoming.has(g.id)) { games.push(incoming.get(g.id)); incoming.delete(g.id); }
  }
  for (const g of incoming.values()) games.push(g);
  n.games = games;
  return n;
}
const fmtSync = (ms) => {
  if (!ms) return "noch nie";
  const d = new Date(ms);
  const p = (x) => String(x).padStart(2, "0");
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()}, ${p(d.getHours())}:${p(d.getMinutes())} Uhr`;
};
const ROLE_LABEL = { tracker: "Tracker (lesen & schreiben)", reader: "Leser (nur lesen)" };

/* ---------- Cloud-Bildschirm ---------- */
function CloudScreen({ data, setData, cloud, setCloud, go }) {
  const [base, setBase] = useState(cloud.apiBase);
  const [busy, setBusy] = useState(null);
  const [msg, setMsg] = useState(null); // { kind: "ok"|"err", text }
  const [addOpen, setAddOpen] = useState(false);
  const [codeIn, setCodeIn] = useState("");
  const [pending, setPending] = useState(null); // { code, slug, label, role } – Cloud noch leer
  const [pickTeam, setPickTeam] = useState("");

  const saveBase = () => {
    const v = base.trim();
    setBase(v);
    setCloud({ ...cloud, apiBase: v });
    setMsg({ kind: "ok", text: "Worker-Adresse gespeichert." });
  };
  const patchLink = (code, patch) =>
    setCloud({ ...cloud, links: cloud.links.map((l) => (l.code === code ? { ...l, ...patch } : l)) });

  const doUpload = async (link, force = false) => {
    const payload = teamPayload(data, link.teamId);
    if (!payload) { setMsg({ kind: "err", text: "Das verknüpfte Team existiert lokal nicht mehr." }); return; }
    setBusy(link.code); setMsg(null);
    try {
      const res = await cloudUpload(cloud.apiBase, link.code, payload, link.baseUpdatedAt, force);
      patchLink(link.code, { baseUpdatedAt: res.updatedAt, lastSyncAt: Date.now() });
      setMsg({ kind: "ok", text: `„${link.label}“ wurde in die Cloud hochgeladen.` });
    } catch (e) {
      if (e.status === 409) {
        const ok = window.confirm(
          "Der Cloud-Stand wurde seit dem letzten Laden von einem anderen Gerät geändert.\n\n" +
          "OK  = Cloud mit dem Stand dieses Geräts überschreiben\n" +
          "Abbrechen = nichts tun (dann besser erst herunterladen)"
        );
        setBusy(null);
        if (ok) return doUpload(link, true);
        setMsg({ kind: "err", text: "Hochladen abgebrochen. Bitte erst den Cloud-Stand herunterladen." });
        return;
      }
      setMsg({ kind: "err", text: e.message });
    } finally { setBusy(null); }
  };

  const doDownload = async (link) => {
    setBusy(link.code); setMsg(null);
    try {
      const res = await cloudDownload(cloud.apiBase, link.code);
      if (!res || !res.data || !res.data.team) {
        setMsg({ kind: "err", text: "In der Cloud liegen für dieses Team noch keine Daten." });
        return;
      }
      if (!window.confirm(`„${link.label}“ wird durch den Cloud-Stand ersetzt. Nicht hochgeladene Änderungen an diesem Team gehen dabei verloren. Fortfahren?`)) {
        setMsg({ kind: "err", text: "Herunterladen abgebrochen." });
        return;
      }
      setData((d) => applyPayload(d, res.data));
      patchLink(link.code, {
        teamId: res.data.team.id, baseUpdatedAt: res.updatedAt,
        label: res.label || link.label, role: res.role || link.role, lastSyncAt: Date.now(),
      });
      setMsg({ kind: "ok", text: `„${res.label || link.label}“ wurde aus der Cloud geladen.` });
    } catch (e) {
      setMsg({ kind: "err", text: e.message });
    } finally { setBusy(null); }
  };

  const addCode = async () => {
    const code = codeIn.trim();
    if (!/^\d{6}$/.test(code)) { setMsg({ kind: "err", text: "Bitte einen 6-stelligen Zugangscode eingeben." }); return; }
    if (cloud.links.some((l) => l.code === code)) { setMsg({ kind: "err", text: "Dieser Zugangscode ist bereits verknüpft." }); return; }
    setBusy("add"); setMsg(null);
    try {
      const res = await cloudDownload(cloud.apiBase, code);
      if (res.data && res.data.team) {
        setData((d) => applyPayload(d, res.data));
        setCloud({
          ...cloud,
          links: [...cloud.links, {
            code, slug: res.slug, label: res.label, role: res.role,
            teamId: res.data.team.id, baseUpdatedAt: res.updatedAt, lastSyncAt: Date.now(),
          }],
        });
        setAddOpen(false); setCodeIn("");
        setMsg({ kind: "ok", text: `„${res.label}“ verknüpft und geladen.` });
      } else if (res.role === "tracker") {
        setPending({ code, slug: res.slug, label: res.label, role: res.role });
        setPickTeam(data.teams[0] ? data.teams[0].id : "");
      } else {
        setMsg({ kind: "err", text: "In der Cloud liegen für dieses Team noch keine Daten. Der Tracker muss zuerst hochladen." });
      }
    } catch (e) {
      setMsg({ kind: "err", text: e.message });
    } finally { setBusy(null); }
  };

  const bindAndUpload = async () => {
    if (!pickTeam) { setMsg({ kind: "err", text: "Bitte ein Team auswählen." }); return; }
    const payload = teamPayload(data, pickTeam);
    if (!payload) return;
    setBusy("add"); setMsg(null);
    try {
      const res = await cloudUpload(cloud.apiBase, pending.code, payload, null, true);
      setCloud({
        ...cloud,
        links: [...cloud.links, {
          code: pending.code, slug: pending.slug, label: pending.label, role: pending.role,
          teamId: pickTeam, baseUpdatedAt: res.updatedAt, lastSyncAt: Date.now(),
        }],
      });
      setPending(null); setAddOpen(false); setCodeIn("");
      setMsg({ kind: "ok", text: `„${payload.team.name}“ wurde als Cloud-Team „${pending.label}“ angelegt.` });
    } catch (e) {
      setMsg({ kind: "err", text: e.message });
    } finally { setBusy(null); }
  };

  return (
    <div>
      <PageTitle title="Cloud-Sync" sub="Daten teamweise sichern und auf mehreren Geräten nutzen"
        onBack={() => go({ name: "teams" })} />

      {msg && (
        <div style={{
          marginBottom: 14, padding: "10px 14px", borderRadius: 12, fontFamily: SANS, fontSize: 14, fontWeight: 700,
          background: msg.kind === "ok" ? C.greenSoft : C.redSoft, color: msg.kind === "ok" ? C.green : C.red,
        }}>{msg.text}</div>
      )}

      <Card style={{ marginBottom: 14 }}>
        <SectionH>Worker-Adresse</SectionH>
        <Field label="Adresse des Cloudflare-Workers">
          <input value={base} onChange={(e) => setBase(e.target.value)} placeholder="https://handball-sync.dein-name.workers.dev"
            style={inputStyle} autoCapitalize="off" autoCorrect="off" spellCheck={false} />
        </Field>
        <Btn kind="soft" small onClick={saveBase}>Adresse speichern</Btn>
      </Card>

      <Card>
        <SectionH>Verknüpfte Teams</SectionH>
        {cloud.links.length === 0 && <Empty>Noch kein Zugangscode hinterlegt.</Empty>}
        {cloud.links.map((l) => {
          const local = data.teams.find((t) => t.id === l.teamId);
          const on = busy === l.code;
          return (
            <div key={l.code} style={{ border: `1px solid ${C.line}`, borderRadius: 14, padding: 14, marginBottom: 10 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
                <div style={{ flex: 1, fontFamily: SANS, fontWeight: 800, fontSize: 16, color: C.ink }}>{l.label}</div>
                <span style={{
                  fontFamily: SANS, fontSize: 12, fontWeight: 800, padding: "4px 10px", borderRadius: 999,
                  background: l.role === "tracker" ? C.blueSoft : C.yellowSoft, color: l.role === "tracker" ? C.blueDark : C.yellow,
                }}>{l.role === "tracker" ? "Tracker" : "Leser"}</span>
              </div>
              <div style={{ fontFamily: SANS, fontSize: 13, color: C.sub, marginBottom: 10 }}>
                Code {l.code} · lokal: {local ? local.name : "— Team fehlt —"} · letzter Sync: {fmtSync(l.lastSyncAt)}
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                {l.role === "tracker" && (
                  <Btn kind="accent" small disabled={on} onClick={() => doUpload(l)}>
                    {on ? "…" : "In Cloud hochladen ↑"}
                  </Btn>
                )}
                <Btn kind="soft" small disabled={on} onClick={() => doDownload(l)}>
                  {on ? "…" : "Aus Cloud laden ↓"}
                </Btn>
                <ConfirmBtn label="Trennen" confirmLabel="Wirklich trennen?"
                  onConfirm={() => setCloud({ ...cloud, links: cloud.links.filter((x) => x.code !== l.code) })} />
              </div>
            </div>
          );
        })}
        <Btn kind="primary" small onClick={() => { setAddOpen(true); setMsg(null); }} style={{ marginTop: 4 }}>
          + Zugangscode hinzufügen
        </Btn>
        <div style={{ fontFamily: SANS, fontSize: 12, color: C.sub, marginTop: 12, lineHeight: 1.5 }}>
          Beim Start der App wird der Cloud-Stand automatisch geladen, sofern eine Internetverbindung
          besteht. Das Hochladen nach dem Spiel startest du hier von Hand. „Trennen“ entfernt nur die
          Verknüpfung – die Daten auf diesem Gerät und in der Cloud bleiben erhalten.
        </div>
      </Card>

      {addOpen && (
        <Modal title={pending ? "Team zuordnen" : "Zugangscode hinzufügen"}
          onClose={() => { setAddOpen(false); setPending(null); setCodeIn(""); }}>
          {!pending && (
            <>
              <Field label="6-stelliger Zugangscode">
                <input value={codeIn} onChange={(e) => setCodeIn(e.target.value.replace(/\D/g, "").slice(0, 6))}
                  inputMode="numeric" placeholder="123456"
                  style={{ ...inputStyle, fontFamily: MONO, fontSize: 24, letterSpacing: "0.3em", textAlign: "center" }} />
              </Field>
              <Btn kind="primary" disabled={busy === "add"} onClick={addCode} style={{ width: "100%" }}>
                {busy === "add" ? "Prüfe …" : "Verknüpfen"}
              </Btn>
            </>
          )}
          {pending && (
            <>
              <div style={{ fontFamily: SANS, fontSize: 14, color: C.sub, marginBottom: 12, lineHeight: 1.5 }}>
                Für „{pending.label}“ liegen in der Cloud noch keine Daten. Wähle das lokale Team,
                das ab jetzt mit diesem Cloud-Team synchronisiert werden soll.
              </div>
              <Field label="Lokales Team">
                <select value={pickTeam} onChange={(e) => setPickTeam(e.target.value)} style={inputStyle}>
                  {data.teams.length === 0 && <option value="">— keine Teams vorhanden —</option>}
                  {data.teams.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                </select>
              </Field>
              <Btn kind="accent" disabled={busy === "add" || !pickTeam} onClick={bindAndUpload} style={{ width: "100%" }}>
                {busy === "add" ? "Lade hoch …" : "Verknüpfen und hochladen"}
              </Btn>
            </>
          )}
        </Modal>
      )}
    </div>
  );
}

export default function App() {
  const [data, setData] = useState(null);
  const [cloud, setCloudState] = useState(null);
  const [boot, setBoot] = useState("Lade Daten …");
  const [toast, setToast] = useState(null);
  const [route, setRoute] = useState({ name: "teams" });
  const loaded = data !== null && cloud !== null;
  const saveT = useRef(null);
  const skip = useRef(true);
  const fileRef = useRef(null);

  useEffect(() => {
    (async () => {
      let d;
      try {
        const r = await window.storage.get(KEY);
        d = r && r.value ? JSON.parse(r.value) : clone(EMPTY);
      } catch { d = clone(EMPTY); }
      const c = await loadCloud();
      // Cloud-Stand beim Start holen – nur online; Fehler werden still übergangen,
      // damit die App in der Halle ohne Netz sofort startet.
      if (c.apiBase && c.links.length && navigator.onLine !== false) {
        setBoot("Cloud-Stand wird geladen …");
        let changed = false;
        for (const link of c.links) {
          try {
            const res = await cloudDownload(c.apiBase, link.code);
            if (res && res.data && res.data.team) {
              d = applyPayload(d, res.data);
              link.teamId = res.data.team.id;
              link.baseUpdatedAt = res.updatedAt;
              link.label = res.label || link.label;
              link.role = res.role || link.role;
              link.lastSyncAt = Date.now();
              changed = true;
            }
          } catch {}
        }
        if (changed) {
          await saveCloud(c);
          try { await window.storage.set(KEY, JSON.stringify(d)); } catch {}
        }
      }
      setCloudState(c);
      setData(d);
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    if (skip.current) { skip.current = false; return; }
    clearTimeout(saveT.current);
    saveT.current = setTimeout(async () => {
      try { await window.storage.set(KEY, JSON.stringify(data)); } catch {}
    }, 400);
    return () => clearTimeout(saveT.current);
  }, [data, loaded]);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 2600);
    return () => clearTimeout(t);
  }, [toast]);

  // Nur-Lese-Modus: greift, sobald das Gerät ausschließlich Leser-Codes hinterlegt hat.
  const readOnly = !!cloud && cloud.links.length > 0 && cloud.links.every((l) => l.role === "reader");
  const setCloud = (next) => { setCloudState(next); saveCloud(next); };
  const update = (fn) => {
    if (readOnly) { setToast("Nur-Lese-Zugang – Änderungen sind nicht möglich."); return; }
    setData((d) => { const n = clone(d); fn(n); return n; });
  };
  const go = (r) => setRoute(r);

  if (!loaded) {
    return (
      <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: SANS, color: C.sub }}>
        {boot}
      </div>
    );
  }
  return (
    <div style={{ minHeight: "100vh", background: C.bg }}>
      <div style={{
        background: C.navy, padding: "10px 16px", display: "flex", alignItems: "center", gap: 10,
        position: "sticky", top: 0, zIndex: 40,
      }}>
        <span onClick={() => go({ name: "teams" })} style={{ cursor: "pointer", fontFamily: SANS, fontWeight: 900, fontSize: 16, color: "#fff", letterSpacing: "-0.01em" }}>
          🤾 Handball-Tracker
        </span>
        <span style={{ flex: 1, fontFamily: SANS, fontSize: 12, color: readOnly ? C.yellow : "#7E93B8" }}>
          {readOnly ? "Nur-Lese-Zugang" : "Saison-Statistik live vom Spielfeldrand"}
        </span>
        <button onClick={() => go({ name: "cloud" })} title="Cloud-Sync"
          style={{ ...btnBase, padding: "6px 10px", fontSize: 12, background: "rgba(255,255,255,0.12)", color: "#fff" }}>
          ☁ Cloud
        </button>
        <button onClick={() => downloadBackup(data)} title="Daten als JSON-Datei sichern"
          style={{ ...btnBase, padding: "6px 10px", fontSize: 12, background: "rgba(255,255,255,0.12)", color: "#fff" }}>
          Backup ↓
        </button>
        <button onClick={() => fileRef.current && fileRef.current.click()} title="Daten aus JSON-Datei laden"
          style={{ ...btnBase, padding: "6px 10px", fontSize: 12, background: "rgba(255,255,255,0.12)", color: "#fff" }}>
          Import ↑
        </button>
        <input ref={fileRef} type="file" accept="application/json,.json" style={{ display: "none" }}
          onChange={(e) => {
            const f = e.target.files && e.target.files[0];
            e.target.value = "";
            if (!f) return;
            if (readOnly) { setToast("Nur-Lese-Zugang – Import ist nicht möglich."); return; }
            if (!window.confirm("Import ersetzt alle aktuellen Daten auf diesem Gerät. Fortfahren?")) return;
            readBackup(f,
              (d) => { setData(d); go({ name: "teams" }); },
              () => window.alert("Die Datei konnte nicht gelesen werden. Bitte ein Backup dieser App auswählen."));
          }} />
      </div>
      <div style={{ maxWidth: 860, margin: "0 auto", padding: 16 }}>
        {route.name === "teams" && <TeamsScreen data={data} update={update} go={go} />}
        {route.name === "team" && <TeamScreen data={data} update={update} go={go} teamId={route.teamId} tab={route.tab || "kader"} />}
        {route.name === "newGame" && <NewGameScreen data={data} update={update} go={go} teamId={route.teamId} />}
        {route.name === "live" && <LiveScreen data={data} update={update} go={go} teamId={route.teamId} gameId={route.gameId} />}
        {route.name === "review" && <ReviewScreen data={data} update={update} go={go} teamId={route.teamId} gameId={route.gameId} />}
        {route.name === "player" && <PlayerScreen data={data} go={go} teamId={route.teamId} playerId={route.playerId} init={route.init} back={route.back} />}
        {route.name === "cloud" && <CloudScreen data={data} setData={setData} cloud={cloud} setCloud={setCloud} go={go} />}
      </div>
      {toast && (
        <div style={{
          position: "fixed", left: "50%", transform: "translateX(-50%)", bottom: 20, zIndex: 60,
          background: C.navy, color: "#fff", fontFamily: SANS, fontSize: 14, fontWeight: 700,
          padding: "10px 18px", borderRadius: 999, boxShadow: "0 8px 24px rgba(0,0,0,0.28)",
        }}>{toast}</div>
      )}
    </div>
  );
}
