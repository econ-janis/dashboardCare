import React, { useMemo, useState } from "react";
import Papa from "papaparse";
import {
  LineChart,
  Line,
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
  Legend,
  ResponsiveContainer,
} from "recharts";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";

/**
 * Janis Commerce - Care Executive Dashboard (React)
 *
 * Reglas importantes:
 * - Fechas: 19/ene/26 12:47 PM (meses en español)
 * - SLA Response: Cumplido si valor >= 0 o vacío. Incumplido solo si valor < 0.
 *   (Incluye 0, 0:00, 00:00 como Cumplido.)
 * - Columna SLA a considerar: "Campo personalizado (Time to first response)" (con o sin punto final).
 * - Organizaciones: basarse en "Campo personalizado (Organizations)"
 * - Dotación: 5 personas (Jun-2024 a Jun-2025), 3 personas (Jul-2025+)
 */

// --- UI (estilo similar al screenshot) ---
const UI = {
  pageBg: "bg-[#eef2f7]",
  card: "bg-white border border-slate-200 rounded-xl shadow-none",
  title: "text-sm font-semibold text-slate-700",
  subtle: "text-xs text-slate-500",
  primary: "#2563eb", // azul principal
  primaryLight: "#60a5fa", // azul claro
  warning: "#f59e0b", // naranjo (incumplido)
  danger: "#ef4444", // rojo (warning)
  ok: "#22c55e", // verde (óptimo)
  grid: "#e5e7eb",
};

const PIE_COLORS = [
  "#2563eb",
  "#60a5fa",
  "#1d4ed8",
  "#93c5fd",
  "#3b82f6",
  "#1e40af",
  "#7dd3fc",
  "#38bdf8",
  "#0ea5e9",
  "#0284c7",
  "#94a3b8", // Otros
];

function coalesce(a: any, b: any) {
  return a === null || a === undefined ? b : a;
}

function hexToRgb(hex: string) {
  const h = String(hex || "").replace("#", "").trim();
  if (h.length !== 6) return null;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  if (![r, g, b].every((x) => Number.isFinite(x))) return null;
  return { r, g, b };
}

// Blanco -> Azul más oscuro (según repetición)
function heatBg(count: number, max: number) {
  if (!count || !max) return { backgroundColor: "#ffffff", color: "#0f172a" };
  const rgb = hexToRgb(UI.primary) || { r: 37, g: 99, b: 235 };
  const ratio = Math.max(0, Math.min(1, count / max));
  const alpha = 0.06 + ratio * 0.82;
  const bg = `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${alpha})`;
  const text = alpha > 0.55 ? "#ffffff" : "#0f172a";
  return { backgroundColor: bg, color: text };
}

const MONTH_MAP: Record<string, string> = {
  ene: "Jan",
  feb: "Feb",
  mar: "Mar",
  abr: "Apr",
  may: "May",
  jun: "Jun",
  jul: "Jul",
  ago: "Aug",
  sep: "Sep",
  oct: "Oct",
  nov: "Nov",
  dic: "Dec",
};

function normalizeSpanishMonth(dateStr: string) {
  if (!dateStr) return "";
  return String(dateStr).replace(
    /\/(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)\//gi,
    (m) => {
      const key = m.split("/").join("").toLowerCase();
      const repl = MONTH_MAP[key] || key;
      return `/${repl}/`;
    }
  );
}

function parseCreated(dateStr: string) {
  if (!dateStr) return null;
  const en = normalizeSpanishMonth(dateStr).trim();

  // Example: 19/Jan/26 12:47 PM
  const match = en.match(
    /^(\d{1,2})\/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\/(\d{2})\s+(\d{1,2}):(\d{2})\s*(AM|PM)$/i
  );
  if (!match) return null;

  const dd = Number(match[1]);
  const mon = match[2];
  const yy = Number(match[3]);
  const hh = Number(match[4]);
  const mm = Number(match[5]);
  const ampm = String(match[6]).toUpperCase();

  const year = yy >= 70 ? 1900 + yy : 2000 + yy;
  const months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  const monthIndex = months.indexOf(mon[0].toUpperCase() + mon.slice(1).toLowerCase());
  if (monthIndex < 0) return null;

  let hour = hh;
  if (ampm === "PM" && hour !== 12) hour += 12;
  if (ampm === "AM" && hour === 12) hour = 0;

  const d = new Date(year, monthIndex, dd, hour, mm, 0, 0);
  return Number.isNaN(d.getTime()) ? null : d;
}

function ym(date: Date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

function teamSizeForMonth(monthStr: string) {
  if (!monthStr || !monthStr.includes("-")) return null;
  const parts = monthStr.split("-");
  if (parts.length !== 2) return null;
  const y = Number(parts[0]);
  const m = Number(parts[1]);
  if (!Number.isFinite(y) || !Number.isFinite(m)) return null;

  // Jun 2024 to Jun 2025 inclusive => 5
  const inFive =
    (y > 2024 || (y === 2024 && m >= 6)) &&
    (y < 2025 || (y === 2025 && m <= 6));
  if (inFive) return 5;

  // Jul 2025 onward => 3
  const inThree = y > 2025 || (y === 2025 && m >= 7);
  if (inThree) return 3;

  return null;
}

function parseSlaHours(s: any) {
  if (s == null) return null;
  const str = String(s).trim();
  if (!str) return null;

  // Numeric (e.g., 1.25, -2, -2,5)
  const numRe = new RegExp("^[+-]?\\d+(?:[\\.,]\\d+)?$");
  if (numRe.test(str)) {
    const num = Number(str.replace(",", "."));
    return Number.isFinite(num) ? num : null;
  }

  // HH:MM with optional sign (e.g., -0:30)
  const hmRe = new RegExp("^([+-])?(\\d+)\\s*:\\s*(\\d{1,2})$");
  const match = str.match(hmRe);
  if (!match) return null;

  const signChar = match[1] || "+";
  const hoursAbs = Number(match[2]);
  const minutesAbs = Number(match[3]);
  if (!Number.isFinite(hoursAbs) || !Number.isFinite(minutesAbs)) return null;

  const val = hoursAbs + minutesAbs / 60;
  return signChar === "-" ? -val : val;
}

function pct(n: number, d: number) {
  if (!d) return 0;
  return (n / d) * 100;
}

function formatInt(n: any) {
  return new Intl.NumberFormat("es-CL").format(Number(n) || 0);
}

function formatPct(n: any) {
  const val = Number(n) || 0;
  return `${val.toFixed(2)}%`;
}

function formatDateCLShort(d: Date) {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}`;
}

function escapeHtml(s: any) {
  const str = String(s ?? "");
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function monthLabel(m: string) {
  // m = YYYY-MM
  if (!m || !m.includes("-")) return m;
  const [y, mm] = m.split("-");
  const names = [
    "Ene",
    "Feb",
    "Mar",
    "Abr",
    "May",
    "Jun",
    "Jul",
    "Ago",
    "Sep",
    "Oct",
    "Nov",
    "Dic",
  ];
  const idx = Number(mm) - 1;
  const n = idx >= 0 && idx < 12 ? names[idx] : mm;
  return `${n} ${y}`;
}

function buildExecutiveReportHtml(args: {
  title: string;
  generatedAt: Date;
  filters: { fromMonth: string; toMonth: string; org: string; assignee: string; status: string };
  autoRange: { minMonth: string | null; maxMonth: string | null };
  kpis: {
    total: number;
    latestMonth: string | null;
    monthCount: number;
    respInc: number;
    respOkPct: number;
    csatAvg: number | null;
    csatCoverage: number;
    tpp6m: number | null;
    tppHealth: { label: string };
  };
  series: {
    ticketsByMonth: Array<{ month: string; tickets: number }>;
    ticketsByYear: Array<{ year: string; tickets: number }>;
    slaByYear: Array<{ year: string; Cumplido: number; Incumplido: number; CumplidoPct: number; IncumplidoPct: number }>;
    csatByYear: Array<{ year: string; csatAvg: number | null; responses: number }>;
    topAssignees: Array<{ name: string; tickets: number }>;
    topOrgsPie: Array<{ name: string; tickets: number }>;
    heatMap: { states: string[]; rows: any[]; range: string };
    hourHeatMap: { data: Array<{ hour: number; tickets: number }>; max: number };
    weekHeatMap: { days: string[]; matrix: any[]; max: number };
  };
}) {
  const { title, generatedAt, filters, autoRange, kpis, series } = args;

  const f = (v: any) => escapeHtml(v);
  const fmtInt = (n: any) => new Intl.NumberFormat("es-CL").format(Number(n) || 0);
  const fmtPct = (n: any) => `${(Number(n) || 0).toFixed(2)}%`;

  const gen = generatedAt;
  const genStr = `${gen.getFullYear()}-${String(gen.getMonth() + 1).padStart(2, "0")}-${String(
    gen.getDate()
  ).padStart(2, "0")} ${String(gen.getHours()).padStart(2, "0")}:${String(gen.getMinutes()).padStart(
    2,
    "0"
  )}`;

  const filterLine = [
    `Archivo: ${f(autoRange.minMonth || "—")} → ${f(autoRange.maxMonth || "—")}`,
    `Vista: ${f(filters.fromMonth)} → ${f(filters.toMonth)}`,
    `Org: ${f(filters.org)}`,
    `Asignado: ${f(filters.assignee)}`,
    `Estado: ${f(filters.status)}`,
  ].join(" • ");

  const kpiRows = [
    { label: "Tickets (vista)", value: fmtInt(kpis.total) },
    {
      label: "Tickets último mes (vista)",
      value: `${fmtInt(kpis.monthCount)}${kpis.latestMonth ? ` (${f(kpis.latestMonth)})` : ""}`,
    },
    { label: "Cumplimiento SLA Response", value: fmtPct(kpis.respOkPct) },
    {
      label: "Incumplidos SLA Response",
      value: `${fmtInt(kpis.respInc)} (${fmtPct(kpis.total ? (kpis.respInc / kpis.total) * 100 : 0)})`,
    },
    {
      label: "CSAT promedio",
      value: kpis.csatAvg == null ? "—" : String(kpis.csatAvg.toFixed(2)),
    },
    { label: "Cobertura CSAT", value: fmtPct(kpis.csatCoverage) },
    {
      label: "Tickets / Persona (prom. 6 meses)",
      value: kpis.tpp6m == null ? "—" : `${kpis.tpp6m.toFixed(1)} (${f(kpis.tppHealth.label)})`,
    },
  ];

  const table = (headers: string[], rows: Array<Array<any>>) => {
    const th = headers.map((h) => `<th>${f(h)}</th>`).join("");
    const tr = rows
      .map((r) => `<tr>${r.map((c) => `<td>${f(c)}</td>`).join("")}</tr>`)
      .join("");
    return `<table><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table>`;
  };

  const ticketsByMonthTable = table(
    ["Mes", "Tickets"],
    series.ticketsByMonth.map((x) => [monthLabel(x.month), fmtInt(x.tickets)])
  );
  const ticketsByYearTable = table(
    ["Año", "Tickets"],
    series.ticketsByYear.map((x) => [x.year, fmtInt(x.tickets)])
  );

  const slaByYearTable = table(
    ["Año", "Cumplido", "Incumplido", "Cumplido %", "Incumplido %"],
    series.slaByYear.map((x) => [
      x.year,
      fmtInt(x.Cumplido),
      fmtInt(x.Incumplido),
      fmtPct(x.CumplidoPct),
      fmtPct(x.IncumplidoPct),
    ])
  );

  const csatByYearTable = table(
    ["Año", "CSAT prom.", "Respuestas"],
    series.csatByYear.map((x) => [x.year, x.csatAvg == null ? "—" : x.csatAvg.toFixed(2), fmtInt(x.responses)])
  );

  const topAssigneesTable = table(
    ["Asignado", "Tickets"],
    series.topAssignees.map((x) => [x.name, fmtInt(x.tickets)])
  );

  const topOrgsTable = table(
    ["Organización", "Tickets", "%"],
    (() => {
      const total = series.topOrgsPie.reduce((s, x) => s + (Number(x.tickets) || 0), 0);
      return series.topOrgsPie.map((x) => [x.name, fmtInt(x.tickets), fmtPct(total ? (x.tickets / total) * 100 : 0)]);
    })()
  );

  // Heatmap mensual (últimos 6 meses)
  const heatHeaders = ["Mes", ...series.heatMap.states];
  const heatRows = series.heatMap.rows.map((r) => [
    monthLabel(r.month),
    ...series.heatMap.states.map((s) => fmtInt(r[s] || 0)),
  ]);
  const heatMonthTable = table(heatHeaders, heatRows);

  const hourTable = table(
    ["Hora", "Tickets"],
    series.hourHeatMap.data.map((x) => [`${String(x.hour).padStart(2, "0")}:00`, fmtInt(x.tickets)])
  );

  const weekHeaders = ["Hora", ...series.weekHeatMap.days];
  const weekRows = series.weekHeatMap.matrix.map((row) => [
    `${String(row.hour).padStart(2, "0")}:00`,
    ...series.weekHeatMap.days.map((d) => fmtInt(row[d] || 0)),
  ]);
  const weekTable = table(weekHeaders, weekRows);

  const css = `
  @page { size: A4; margin: 16mm; }
  body { font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial; color:#0f172a; }
  h1 { font-size: 18px; margin: 0 0 6px; }
  .meta { font-size: 11px; color:#475569; margin-bottom: 14px; }
  .section { margin: 14px 0 18px; }
  .section h2 { font-size: 13px; margin: 0 0 8px; color:#1f2937; }
  .kpis { display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; }
  .kpi { border: 1px solid #e2e8f0; border-radius: 10px; padding: 10px; }
  .kpi .label { font-size: 11px; color:#64748b; }
  .kpi .value { font-size: 16px; font-weight: 700; margin-top: 2px; }
  table { width: 100%; border-collapse: collapse; font-size: 10px; }
  th, td { border: 1px solid #e2e8f0; padding: 6px 8px; }
  th { background: #f8fafc; text-align: left; }
  .twoCol { display:grid; grid-template-columns: 1fr 1fr; gap: 12px; }
  .note { font-size: 10px; color:#64748b; margin-top: 6px; }
  .badge { display:inline-block; padding:2px 8px; border-radius:999px; background:#eff6ff; color:#1d4ed8; font-weight:600; font-size:10px; }
  `;

  return `<!doctype html>
  <html>
    <head>
      <meta charset="utf-8" />
      <title>${f(title)}</title>
      <style>${css}</style>
    </head>
    <body>
      <h1>${f(title)}</h1>
      <div class="meta">
        <div><span class="badge">Informe Ejecutivo</span> • Generado: ${f(genStr)}</div>
        <div style="margin-top:6px;">${filterLine}</div>
      </div>

      <div class="section">
        <h2>KPIs</h2>
        <div class="kpis">
          ${kpiRows
            .map(
              (k) => `<div class="kpi"><div class="label">${f(k.label)}</div><div class="value">${f(
                k.value
              )}</div></div>`
            )
            .join("")}
        </div>
        <div class="note">Regla SLA Response (Time to first response): Cumplido si valor &gt;= 0 o vacío; Incumplido solo si valor &lt; 0.</div>
      </div>

      <div class="section twoCol">
        <div>
          <h2>Tickets por Mes</h2>
          ${ticketsByMonthTable}
        </div>
        <div>
          <h2>Tickets por Año</h2>
          ${ticketsByYearTable}
        </div>
      </div>

      <div class="section twoCol">
        <div>
          <h2>SLA Response por Año</h2>
          ${slaByYearTable}
        </div>
        <div>
          <h2>CSAT promedio por Año</h2>
          ${csatByYearTable}
        </div>
      </div>

      <div class="section twoCol">
        <div>
          <h2>Top 10 Asignados</h2>
          ${topAssigneesTable}
        </div>
        <div>
          <h2>Top 10 Organizaciones + Otros</h2>
          ${topOrgsTable}
        </div>
      </div>

      <div class="section">
        <h2>Heatmap Mes vs Estado (últimos 6 meses)</h2>
        ${heatMonthTable}
      </div>

      <div class="section twoCol">
        <div>
          <h2>Heatmap Horario (por hora)</h2>
          ${hourTable}
        </div>
        <div>
          <h2>Heatmap Semana (día vs hora)</h2>
          ${weekTable}
        </div>
      </div>

      <div class="note">Sugerencia: aplica enfoque Pareto 80/20 sobre Top Organizaciones/Asignados para reducir demanda recurrente.</div>
    </body>
  </html>`;
}

function kpiCard(
  title: string,
  value: any,
  subtitle?: string,
  right?: string,
  badge?: React.ReactNode
) {
  return (
    <Card className={UI.card}>
      <CardHeader className="pb-2">
        <CardTitle className={UI.title}>{title}</CardTitle>
      </CardHeader>
      <CardContent>
        <div className="flex items-start justify-between gap-3">
          <div>
            <div className="text-2xl font-semibold tracking-tight text-slate-900">{value}</div>
            {badge ? <div className="mt-2">{badge}</div> : null}
            {subtitle ? <div className={"mt-1 " + UI.subtle}>{subtitle}</div> : null}
          </div>
          {right ? <div className={"text-right " + UI.subtle}>{right}</div> : null}
        </div>
      </CardContent>
    </Card>
  );
}

function getField(row: Record<string, any>, candidates: string[]) {
  for (const c of candidates) {
    const v = row[c];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  for (const c of candidates) {
    if (Object.prototype.hasOwnProperty.call(row, c)) return row[c];
  }
  return undefined;
}

function pieTooltipFormatterFactory(data: Array<{ name: string; tickets: number }>) {
  return (value: any, _name: any, props: any) => {
    const v = Number(value) || 0;
    const total = (data || []).reduce((s, x) => s + (Number(x.tickets) || 0), 0);
    const p = total ? (v / total) * 100 : 0;
    const label = props && props.payload && props.payload.name ? props.payload.name : "";
    return [`${formatInt(v)} (${p.toFixed(2)}%)`, label];
  };
}

// Mini-tests (dev) para mantener consistente el parsing
let __testsRan = false;
function runParserTestsOnce() {
  if (__testsRan) return;
  __testsRan = true;

  console.assert(
    parseCreated("19/ene/26 12:47 PM") instanceof Date,
    "parseCreated should parse Spanish months"
  );
  console.assert(ym(new Date(2026, 0, 19)) === "2026-01", "ym should format YYYY-MM");

  console.assert(parseSlaHours("") === null, "blank SLA should be null");
  console.assert(parseSlaHours(null) === null, "null SLA should be null");
  console.assert(parseSlaHours("0:00") === 0, "0:00 SLA should be 0");
  console.assert(parseSlaHours("00:00") === 0, "00:00 SLA should be 0");
  console.assert(parseSlaHours("1:30") > 0, "positive SLA should be > 0");
  console.assert(parseSlaHours("-1:30") < 0, "negative SLA should be < 0");

  console.assert(
    getField({ a: "", b: "x" }, ["a", "b"]) === "x",
    "getField should return first non-empty"
  );

  console.assert(teamSizeForMonth("2024-06") === 5, "Jun 2024 => 5");
  console.assert(teamSizeForMonth("2025-06") === 5, "Jun 2025 => 5");
  console.assert(teamSizeForMonth("2025-07") === 3, "Jul 2025 => 3");
}

type Row = {
  key: string;
  organization: string;
  estado: string;
  asignado: string;
  creada: Date;
  year: number;
  month: string;
  slaResponseHours: number | null;
  slaResponseStatus: "Cumplido" | "Incumplido";
  satisfaction: number | null;
};

export default function JiraExecutiveDashboard() {
  if (typeof window !== "undefined") runParserTestsOnce();

  const [rows, setRows] = useState<Row[]>([]);
  const [error, setError] = useState<string | null>(null);

  // Filters: rango de fechas por mes (YYYY-MM)
  // Se inicializan según la completitud del CSV (minMonth → maxMonth)
  const [fromMonth, setFromMonth] = useState<string>("all");
  const [toMonth, setToMonth] = useState<string>("all");

  const [autoRange, setAutoRange] = useState<{ minMonth: string | null; maxMonth: string | null }>({
    minMonth: null,
    maxMonth: null,
  });

  const [orgFilter, setOrgFilter] = useState("all");
  const [assigneeFilter, setAssigneeFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");

  const onFile = (file: File) => {
    setError(null);

    Papa.parse(file, {
      transformHeader: (h) => String(h || "").trim().toLowerCase(),
      header: true,
      skipEmptyLines: true,
      complete: (res: any) => {
        try {
          const data = (res.data || []).filter(Boolean);
          const parsed: Row[] = [];
          let badDate = 0;

          for (const r of data) {
            const creadaRaw = String(coalesce(r["creada"], "")).trim();
            const creada = parseCreated(creadaRaw);
            if (!creada) {
              badDate += 1;
              continue;
            }

            const slaRespRaw = getField(r, [
              "campo personalizado (time to first response)",
              "campo personalizado (time to first response).",
              "custom field (time to first response)",
              "custom field (time to first response).",
              "time to first response",
              "time to first response (hrs)",
              "sla response",
              "sla de response",
            ]);
            const slaResp = parseSlaHours(slaRespRaw);
            const respStatus: Row["slaResponseStatus"] =
              slaResp != null && slaResp < 0 ? "Incumplido" : "Cumplido";

            const satRaw = getField(r, [
              "calificación de satisfacción",
              "calificacion de satisfaccion",
              "satisfaction",
            ]);
            const satStr = satRaw == null ? "" : String(satRaw).trim();
            const satVal = satStr === "" ? null : Number(satStr);
            const sat = Number.isFinite(satVal as any) ? (satVal as number) : null;

            const org = String(
              coalesce(
                getField(r, [
                  "campo personalizado (organizations)",
                  "organizations",
                  "organization",
                  "organisation",
                ]),
                ""
              )
            ).trim();

            const estado = String(coalesce(r["estado"], "")).trim();
            // Excluir Block/Hold (no entran a KPIs ni gráficos)
            if (/\b(block|hold)\b/i.test(estado)) continue;

            parsed.push({
              key: String(coalesce(r["clave de incidencia"], coalesce(r["key"], ""))).trim(),
              organization: org,
              estado,
              asignado: String(coalesce(r["persona asignada"], "")).trim(),
              creada,
              year: creada.getFullYear(),
              month: ym(creada),
              slaResponseHours: slaResp,
              slaResponseStatus: respStatus,
              satisfaction: sat,
            });
          }

          if (!parsed.length) {
            setRows([]);
            setError(
              "No pude parsear filas con fecha 'creada'. Revisa que el CSV tenga columna 'Creada' y formato tipo 19/ene/26 12:47 PM."
            );
            return;
          }

          parsed.sort((a, b) => a.creada.getTime() - b.creada.getTime());
          setRows(parsed);

          // Auto ajustar rango de fechas al mínimo y máximo del archivo
          const minMonth = parsed[0].month;
          const maxMonth = parsed[parsed.length - 1].month;
          setAutoRange({ minMonth, maxMonth });
          setFromMonth(minMonth);
          setToMonth(maxMonth);

          if (badDate > 0) {
            setError(`Aviso: ${badDate} filas fueron omitidas porque la fecha 'creada' no era interpretable.`);
          }
        } catch (e: any) {
          setError((e && e.message) || "Error procesando el CSV");
          setRows([]);
        }
      },
      error: (err: any) => {
        setError(err.message);
        setRows([]);
      },
    });
  };

  const filterOptions = useMemo(() => {
    const orgs = Array.from(new Set(rows.map((r) => r.organization).filter(Boolean))).sort();
    const assignees = Array.from(new Set(rows.map((r) => r.asignado).filter(Boolean))).sort();
    const estados = Array.from(new Set(rows.map((r) => r.estado).filter(Boolean))).sort();
    const months = Array.from(new Set(rows.map((r) => r.month))).sort();
    return { orgs, assignees, estados, months };
  }, [rows]);

  const filtered = useMemo(() => {
    return rows.filter((r) => {
      if (fromMonth !== "all" && r.month < fromMonth) return false;
      if (toMonth !== "all" && r.month > toMonth) return false;
      if (orgFilter !== "all" && r.organization !== orgFilter) return false;
      if (assigneeFilter !== "all" && r.asignado !== assigneeFilter) return false;
      if (statusFilter !== "all" && r.estado !== statusFilter) return false;
      return true;
    });
  }, [rows, fromMonth, toMonth, orgFilter, assigneeFilter, statusFilter]);

  const kpis = useMemo(() => {
    const total = filtered.length;
    const respInc = filtered.filter((r) => r.slaResponseStatus === "Incumplido").length;

    const rated = filtered.filter((r) => r.satisfaction != null);
    const csatAvg =
      rated.length > 0
        ? rated.reduce((s, r) => s + (r.satisfaction == null ? 0 : r.satisfaction), 0) /
          rated.length
        : null;

    const latestMonth = total > 0 ? filtered[total - 1].month : null;
    const monthCount = latestMonth ? filtered.filter((r) => r.month === latestMonth).length : 0;

    // Tickets/Persona: Promedio últimos 6 meses (sin considerar mes actual si no está cerrado)
    const monthsSorted = Array.from(new Set(filtered.map((r) => r.month))).sort();
    const maxCreated = filtered.length ? filtered[filtered.length - 1].creada : null;
    const currentMonth = maxCreated ? ym(maxCreated) : null;

    const isClosedMonth = (d: Date | null) => {
      if (!d) return true;
      const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
      return d.getDate() === lastDay;
    };

    const monthsForAvg = (() => {
      if (!monthsSorted.length) return [];
      if (!maxCreated || !currentMonth) return monthsSorted;
      if (!isClosedMonth(maxCreated)) return monthsSorted.filter((m) => m !== currentMonth);
      return monthsSorted;
    })();

    const last6 = monthsForAvg.slice(-6);
    const tppByMonth = last6
      .map((m) => {
        const ts = teamSizeForMonth(m);
        if (!ts) return null;
        const tickets = filtered.filter((r) => r.month === m).length;
        return { month: m, tickets, team: ts, tpp: tickets / ts };
      })
      .filter(Boolean) as Array<{ month: string; tickets: number; team: number; tpp: number }>;

    const tpp6m =
      tppByMonth.length > 0 ? tppByMonth.reduce((s, x) => s + x.tpp, 0) / tppByMonth.length : null;

    const tppHealth = (() => {
      if (tpp6m == null) return { label: "Sin dato", color: "#94a3b8" };
      if (tpp6m < 40) return { label: "Con Capacidad", color: UI.primary };
      if (tpp6m >= 40 && tpp6m <= 70) return { label: "Óptimo", color: UI.ok };
      if (tpp6m > 70 && tpp6m <= 95) return { label: "Al Límite", color: UI.warning };
      return { label: "Warning", color: UI.danger };
    })();

    return {
      total,
      latestMonth,
      monthCount,
      respInc,
      respOkPct: 100 - pct(respInc, total),
      csatAvg,
      csatCoverage: pct(rated.length, total),
      tpp6m,
      tppHealth,
    };
  }, [filtered]);

  const series = useMemo(() => {
    // Tickets por mes
    const byMonth = new Map<string, { month: string; tickets: number }>();
    for (const r of filtered) {
      const cur = byMonth.get(r.month) || { month: r.month, tickets: 0 };
      cur.tickets += 1;
      byMonth.set(r.month, cur);
    }
    const ticketsByMonth = Array.from(byMonth.values()).sort((a, b) => a.month.localeCompare(b.month));

    // Tickets por Año
    const byYear = new Map<string, { year: string; tickets: number }>();
    for (const r of filtered) {
      const key = String(r.year);
      const cur = byYear.get(key) || { year: key, tickets: 0 };
      cur.tickets += 1;
      byYear.set(key, cur);
    }
    const ticketsByYear = Array.from(byYear.values()).sort((a, b) => Number(a.year) - Number(b.year));

    // Estado por Año
    const yearStatus = new Map<string, any>();
    for (const r of filtered) {
      const y = String(r.year);
      const obj = yearStatus.get(y) || { year: y };
      const s = r.estado || "(Sin estado)";
      obj[s] = (obj[s] || 0) + 1;
      yearStatus.set(y, obj);
    }
    const estadoByYear = Array.from(yearStatus.values()).sort((a, b) => Number(a.year) - Number(b.year));

    // SLA Response por Año (conteo + porcentajes)
    const slaYear = new Map<string, any>();
    for (const r of filtered) {
      const y = String(r.year);
      const obj = slaYear.get(y) || {
        year: y,
        Total: 0,
        Cumplido: 0,
        Incumplido: 0,
        CumplidoPct: 0,
        IncumplidoPct: 0,
      };
      obj.Total += 1;
      if (r.slaResponseStatus === "Incumplido") obj.Incumplido += 1;
      else obj.Cumplido += 1;
      slaYear.set(y, obj);
    }

    const slaByYear = Array.from(slaYear.values())
      .map((x) => ({
        ...x,
        CumplidoPct: x.Total ? (x.Cumplido / x.Total) * 100 : 0,
        IncumplidoPct: x.Total ? (x.Incumplido / x.Total) * 100 : 0,
      }))
      .sort((a, b) => Number(a.year) - Number(b.year));

    // Conteos genéricos (Top)
    const count = (keyFn: (r: Row) => string) => {
      const m = new Map<string, number>();
      for (const r of filtered) {
        const k = (keyFn(r) || "(Vacío)").trim();
        m.set(k, (m.get(k) || 0) + 1);
      }
      return Array.from(m.entries())
        .map(([k, v]) => ({ name: k, tickets: v }))
        .sort((a, b) => b.tickets - a.tickets);
    };

    const totalTickets = filtered.length;
    const topAssignees = count((r) => r.asignado).slice(0, 10);

    // Pie Top 10 orgs + Otros
    const allOrgs = count((r) => r.organization);
    const top10 = allOrgs.slice(0, 10);
    const top10Sum = top10.reduce((s, x) => s + x.tickets, 0);
    const others = Math.max(0, totalTickets - top10Sum);
    const topOrgsPie = [...top10];
    if (others > 0) topOrgsPie.push({ name: "Otros", tickets: others });

    // CSAT promedio por Año
    const csatByYearMap = new Map<string, { year: string; sum: number; cnt: number }>();
    for (const r of filtered) {
      if (r.satisfaction == null) continue;
      const y = String(r.year);
      const cur = csatByYearMap.get(y) || { year: y, sum: 0, cnt: 0 };
      cur.sum += Number(r.satisfaction) || 0;
      cur.cnt += 1;
      csatByYearMap.set(y, cur);
    }
    const csatByYear = Array.from(csatByYearMap.values())
      .map((x) => ({
        year: x.year,
        csatAvg: x.cnt ? x.sum / x.cnt : null,
        responses: x.cnt,
      }))
      .sort((a, b) => Number(a.year) - Number(b.year));

    // Tickets/persona vs % incumplido RESPONSE por mes
    const eff = ticketsByMonth.map((m) => {
      const monthRows = filtered.filter((r) => r.month === m.month);
      const incResp = monthRows.filter((r) => r.slaResponseStatus === "Incumplido").length;
      const ts = teamSizeForMonth(m.month);
      const tpp = ts ? m.tickets / ts : null;
      return {
        month: m.month,
        ticketsPerPerson: tpp,
        respIncPct: pct(incResp, monthRows.length),
      };
    });

    // Heatmap mes vs estado (solo últimos 6 meses)
    const heatMap = (() => {
      const states = Array.from(new Set(filtered.map((r) => r.estado || "(Sin estado)"))).sort();
      const byM = new Map<string, any>();
      for (const r of filtered) {
        const key = r.month;
        const obj = byM.get(key) || { month: key };
        const s = r.estado || "(Sin estado)";
        obj[s] = (obj[s] || 0) + 1;
        byM.set(key, obj);
      }
      const allRows = Array.from(byM.values()).sort((a, b) => a.month.localeCompare(b.month));
      const rows = allRows.slice(-6);
      const range = rows.length ? `${rows[0].month} → ${rows[rows.length - 1].month}` : "—";
      return { states, rows, range };
    })();

    // Heatmap horario (global): tickets por hora 0-23, ajustado a filtros
    const hourHeatMap = (() => {
      const hours = Array.from({ length: 24 }, (_, i) => i);
      const counts = new Map<number, number>();
      for (const r of filtered) {
        const h = r.creada instanceof Date ? r.creada.getHours() : null;
        if (h == null) continue;
        counts.set(h, (counts.get(h) || 0) + 1);
      }
      const data = hours.map((h) => ({ hour: h, tickets: counts.get(h) || 0 }));
      const max = data.reduce((m, x) => Math.max(m, x.tickets || 0), 0);
      return { hours, data, max };
    })();

    // Heatmap semanal: filas = horas (0-23), columnas = dias (Lun..Dom)
    const weekHeatMap = (() => {
      const days = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]; // orden ISO
      const hours = Array.from({ length: 24 }, (_, i) => i);
      const matrix = hours.map((h) => ({
        hour: h,
        ...Object.fromEntries(days.map((d) => [d, 0])),
      })) as Array<Record<string, any>>;

      const isoDayIndex = (jsDay: number) => (jsDay + 6) % 7;

      for (const r of filtered) {
        if (!(r.creada instanceof Date)) continue;
        const h = r.creada.getHours();
        const di = isoDayIndex(r.creada.getDay());
        const dLabel = days[di];
        matrix[h][dLabel] = (matrix[h][dLabel] || 0) + 1;
      }

      let max = 0;
      for (const row of matrix) {
        for (const d of days) max = Math.max(max, row[d] || 0);
      }

      return { days, hours, matrix, max };
    })();

    return {
      ticketsByMonth,
      ticketsByYear,
      estadoByYear,
      slaByYear,
      csatByYear,
      topAssignees,
      topOrgsPie,
      eff,
      heatMap,
      hourHeatMap,
      weekHeatMap,
    };
  }, [filtered]);

  const estadoKeys = useMemo(() => {
    const keys = new Set<string>();
    for (const obj of series.estadoByYear) {
      Object.keys(obj).forEach((k) => {
        if (k !== "year") keys.add(k);
      });
    }
    return Array.from(keys).sort();
  }, [series.estadoByYear]);

  const pieTooltipFormatter = useMemo(
    () => pieTooltipFormatterFactory(series.topOrgsPie),
    [series.topOrgsPie]
  );

  const ticketsByYearBars = useMemo(() => {
    const items = series.ticketsByYear || [];
    const maxTickets = items.reduce((m, x) => Math.max(m, Number(x.tickets) || 0), 0);

    const maxCreated = filtered.length ? filtered[filtered.length - 1].creada : null;
    const maxYear = maxCreated ? maxCreated.getFullYear() : null;
    const isPartialYear =
      !!maxCreated && !(maxCreated.getMonth() === 11 && maxCreated.getDate() === 31);

    return {
      maxTickets,
      rows: items.map((x) => {
        const y = Number(x.year);
        const partial = maxYear != null && y === maxYear && isPartialYear;
        return {
          year: String(x.year),
          tickets: Number(x.tickets) || 0,
          partialLabel: partial && maxCreated ? ` (parcial al ${formatDateCLShort(maxCreated)})` : "",
        };
      }),
    };
  }, [series.ticketsByYear, filtered]);

  return (
    <div className={`min-h-screen ${UI.pageBg} p-4 md:p-8`}>
      <div className="mx-auto max-w-7xl">
        <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
          <div>
            <h1 className="text-2xl md:text-3xl font-semibold tracking-tight">
              Janis Commerce -  Care Executive Dashboard
            </h1>
            <p className="text-sm text-slate-500 mt-1">
              Sube tu CSV y verás KPIs y gráficos. Regla SLA Response (Time to first response):
              &gt;= 0 (o vacío) = cumplido; &lt; 0 = incumplido. Dotación: 5 (Jun-2024 a Jun-2025), 3
              (Jul-2025+).
            </p>
          </div>

          <div className="flex flex-col sm:flex-row gap-2">
            <Input
              type="file"
              accept=".csv,text/csv"
              onChange={(e) => {
                const f = e.target.files && e.target.files[0];
                if (f) onFile(f);
              }}
            />

            <Button
              className="text-white"
              style={{ backgroundColor: UI.primary }}
              onClick={() => {
                if (!filtered.length) {
                  setError("No hay datos filtrados para exportar.");
                  return;
                }

                const html = buildExecutiveReportHtml({
                  title: "Janis Commerce -  Care Executive Dashboard",
                  generatedAt: new Date(),
                  filters: {
                    fromMonth: fromMonth === "all" ? (autoRange.minMonth || "all") : fromMonth,
                    toMonth: toMonth === "all" ? (autoRange.maxMonth || "all") : toMonth,
                    org: orgFilter === "all" ? "Todas" : orgFilter,
                    assignee: assigneeFilter === "all" ? "Todos" : assigneeFilter,
                    status: statusFilter === "all" ? "Todos" : statusFilter,
                  },
                  autoRange,
                  kpis,
                  series,
                });

                const w = window.open("", "_blank");
                if (!w) {
                  setError("Tu navegador bloqueó la ventana emergente. Habilita pop-ups para exportar.");
                  return;
                }
                w.document.open();
                w.document.write(html);
                w.document.close();

                // Espera a que el contenido se renderice y abre el diálogo de impresión (guardar como PDF)
                w.onload = () => {
                  w.focus();
                  w.print();
                };
              }}
            >
              Exportar Informe
            </Button>

            <Button
              variant="secondary"
              onClick={() => {
                setRows([]);
                setError(null);
                setFromMonth("all");
                setToMonth("all");
                setAutoRange({ minMonth: null, maxMonth: null });
                setOrgFilter("all");
                setAssigneeFilter("all");
                setStatusFilter("all");
              }}
            >
              Limpiar
            </Button>
          </div>
        </div>

        {error ? (
          <div className="mt-4 rounded-xl bg-white border border-slate-200 p-3 text-sm text-slate-700">
            {error}
          </div>
        ) : null}

        {/* Filters */}
        <div className="mt-6 grid grid-cols-1 md:grid-cols-5 gap-3">
          <Card className={UI.card + " md:col-span-5"}>
            <CardContent className="p-4">
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-6 gap-3">
                <div className="flex flex-col gap-1">
                  <div className="text-xs text-slate-500">Desde</div>
                  <Input
                    type="month"
                    value={fromMonth === "all" ? "" : fromMonth}
                    onChange={(e) => setFromMonth(e.target.value || "all")}
                    className="rounded-xl bg-white border-slate-200"
                    min={filterOptions.months[0] || undefined}
                    max={filterOptions.months[filterOptions.months.length - 1] || undefined}
                  />
                </div>

                <div className="flex flex-col gap-1">
                  <div className="text-xs text-slate-500">Hasta</div>
                  <Input
                    type="month"
                    value={toMonth === "all" ? "" : toMonth}
                    onChange={(e) => setToMonth(e.target.value || "all")}
                    className="rounded-xl bg-white border-slate-200"
                    min={filterOptions.months[0] || undefined}
                    max={filterOptions.months[filterOptions.months.length - 1] || undefined}
                  />
                </div>

                <Select value={orgFilter} onValueChange={setOrgFilter}>
                  <SelectTrigger className="rounded-xl bg-white border-slate-200">
                    <SelectValue placeholder="Organización" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="all">Todas las orgs</SelectItem>
                    {filterOptions.orgs.map((o) => (
                      <SelectItem key={o} value={o}>
                        {o}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>

                <Select value={assigneeFilter} onValueChange={setAssigneeFilter}>
                  <SelectTrigger className="rounded-xl bg-white border-slate-200">
                    <SelectValue placeholder="Asignado" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="all">Todos los asignados</SelectItem>
                    {filterOptions.assignees.map((a) => (
                      <SelectItem key={a} value={a}>
                        {a}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>

                <Select value={statusFilter} onValueChange={setStatusFilter}>
                  <SelectTrigger className="rounded-xl bg-white border-slate-200">
                    <SelectValue placeholder="Estado" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="all">Todos los estados</SelectItem>
                    {filterOptions.estados.map((s) => (
                      <SelectItem key={s} value={s}>
                        {s}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="mt-3 text-xs text-slate-500">
                Rango archivo: <span className="font-medium">{autoRange.minMonth || "—"}</span> →{" "}
                <span className="font-medium">{autoRange.maxMonth || "—"}</span>
                <span className="mx-2 text-slate-300">|</span>
                Filas en vista: <span className="font-medium">{formatInt(filtered.length)}</span>
                {kpis.latestMonth ? (
                  <>
                    {" — "}Último mes en vista: <span className="font-medium">{kpis.latestMonth}</span>
                  </>
                ) : null}
              </div>
            </CardContent>
          </Card>
        </div>

        {/* KPI Cards */}
        <div className="mt-6 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-3">
          {kpiCard("Tickets (vista)", formatInt(kpis.total), "Total según filtros")}
          {kpiCard(
            "Tickets en último mes",
            formatInt(kpis.monthCount),
            kpis.latestMonth ? `Mes: ${kpis.latestMonth}` : "—"
          )}
          {kpiCard(
            "% Cumplimiento SLA Response",
            formatPct(kpis.respOkPct),
            `${formatInt(kpis.total - kpis.respInc)} cumplidos`
          )}
          {kpiCard(
            "CSAT promedio",
            kpis.csatAvg == null ? "—" : kpis.csatAvg.toFixed(2),
            `Cobertura: ${formatPct(kpis.csatCoverage)}`
          )}
          {kpiCard(
            "Tickets / Persona (prom. 6 meses)",
            kpis.tpp6m == null ? "—" : kpis.tpp6m.toFixed(1),
            "Promedio últimos 6 meses",
            "5 (Jun24-Jun25) | 3 (Jul25+)",
            <span
              className="inline-flex items-center gap-2 rounded-full px-3 py-1 text-xs font-medium"
              style={{ background: "rgba(15, 23, 42, 0.04)", color: kpis.tppHealth.color }}
            >
              ● {kpis.tppHealth.label}
            </span>
          )}
        </div>

        {/* Main grid */}
        <div className="mt-6 grid grid-cols-1 lg:grid-cols-12 gap-3">
          <Card className={UI.card + " lg:col-span-8"}>
            <CardHeader>
              <CardTitle className={UI.title}>Tickets por mes</CardTitle>
            </CardHeader>
            <CardContent className="h-[320px]">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={series.ticketsByMonth} margin={{ left: 8, right: 8 }}>
                  <CartesianGrid stroke={UI.grid} strokeDasharray="3 3" />
                  <XAxis dataKey="month" tick={{ fontSize: 12 }} />
                  <YAxis tick={{ fontSize: 12 }} />
                  <Tooltip />
                  <Legend />
                  <Line
                    type="monotone"
                    dataKey="tickets"
                    name="Tickets"
                    dot={false}
                    stroke={UI.primary}
                    strokeWidth={2}
                  />
                </LineChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>

          <Card className={UI.card + " lg:col-span-4"}>
            <CardHeader>
              <CardTitle className={UI.title}>Tickets por Año</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="flex flex-col gap-3">
                {ticketsByYearBars.rows.map((r) => {
                  const pctW = ticketsByYearBars.maxTickets
                    ? (r.tickets / ticketsByYearBars.maxTickets) * 100
                    : 0;
                  return (
                    <div key={r.year} className="grid grid-cols-[48px_1fr_auto] items-center gap-3">
                      <div className="text-xs font-medium text-slate-600 tabular-nums">{r.year}</div>
                      <div className="h-3 rounded-full bg-slate-200 overflow-hidden">
                        <div
                          className="h-full rounded-full"
                          style={{ width: `${pctW}%`, backgroundColor: UI.primary }}
                        />
                      </div>
                      <div className="text-xs font-semibold text-slate-700 tabular-nums whitespace-nowrap">
                        {formatInt(r.tickets)}
                        <span className="ml-2 text-[10px] text-slate-400 font-medium">
                          {r.partialLabel}
                        </span>
                      </div>
                    </div>
                  );
                })}
              </div>
            </CardContent>
          </Card>

          <Card className={UI.card + " lg:col-span-7"}>
            <CardHeader>
              <CardTitle className={UI.title}>SLA Response por año (Cumplido vs Incumplido %)</CardTitle>
            </CardHeader>
            <CardContent className="h-[340px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={series.slaByYear} margin={{ left: 8, right: 8 }}>
                  <CartesianGrid stroke={UI.grid} strokeDasharray="3 3" />
                  <XAxis dataKey="year" tick={{ fontSize: 12 }} />
                  <YAxis tick={{ fontSize: 12 }} domain={[0, 100]} tickFormatter={(v) => `${v}%`} />
                  <Tooltip formatter={(v: any) => [`${Number(v).toFixed(2)}%`, ""]} />
                  <Legend />
                  <Bar dataKey="CumplidoPct" name="Cumplido %" stackId="respPct" fill={UI.primary} />
                  <Bar dataKey="IncumplidoPct" name="Incumplido %" stackId="respPct" fill={UI.warning} />
                </BarChart>
              </ResponsiveContainer>
              <div className="mt-2 text-xs text-slate-500">
                Nota: SLA Response (Time to first response) cumplido si valor &gt;= 0 o vacío. Incumplido solo si valor
                &lt; 0.
              </div>
            </CardContent>
          </Card>

          <Card className={UI.card + " lg:col-span-5"}>
            <CardHeader>
              <CardTitle className={UI.title}>Carga vs calidad (mensual) — SLA Response</CardTitle>
            </CardHeader>
            <CardContent className="h-[340px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={series.eff} margin={{ left: 8, right: 8 }}>
                  <CartesianGrid stroke={UI.grid} strokeDasharray="3 3" />
                  <XAxis dataKey="month" tick={{ fontSize: 12 }} />
                  <YAxis tick={{ fontSize: 12 }} />
                  <Tooltip />
                  <Legend />
                  <Bar dataKey="ticketsPerPerson" name="Tickets/Persona" fill={UI.primary} />
                  <Bar dataKey="respIncPct" name="% Incumplido Response" fill={UI.primaryLight} />
                </BarChart>
              </ResponsiveContainer>
              <div className="mt-2 text-xs text-slate-500">
                Tickets/Persona solo se calcula cuando existe dotación (Jun-2024+).
              </div>
            </CardContent>
          </Card>

          <Card className={UI.card + " lg:col-span-7"}>
            <CardHeader>
              <CardTitle className={UI.title}>Tickets por Estado (por Año)</CardTitle>
            </CardHeader>
            <CardContent className="h-[340px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={series.estadoByYear} margin={{ left: 8, right: 8 }}>
                  <CartesianGrid stroke={UI.grid} strokeDasharray="3 3" />
                  <XAxis dataKey="year" tick={{ fontSize: 12 }} />
                  <YAxis tick={{ fontSize: 12 }} />
                  <Tooltip />
                  <Legend />
                  {estadoKeys.map((k) => (
                    <Bar key={k} dataKey={k} name={k} stackId="estado" fill={UI.primary} />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>

          <Card className={UI.card + " lg:col-span-5"}>
            <CardHeader>
              <CardTitle className={UI.title}>Distribución de Tickets por Organización (Top 10)</CardTitle>
            </CardHeader>
            <CardContent className="h-[340px]">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Tooltip formatter={pieTooltipFormatter} />
                  <Legend />
                  <Pie
                    data={series.topOrgsPie}
                    dataKey="tickets"
                    nameKey="name"
                    innerRadius={55}
                    outerRadius={95}
                    paddingAngle={2}
                  >
                    {series.topOrgsPie.map((_, idx) => (
                      <Cell
                        key={`cell-${idx}`}
                        fill={PIE_COLORS[idx % PIE_COLORS.length]}
                        stroke="#ffffff"
                        strokeWidth={1}
                      />
                    ))}
                  </Pie>
                </PieChart>
              </ResponsiveContainer>
              <div className="mt-2 text-xs text-slate-500">
                Muestra el Top 10 en colores. Todo lo demás se agrupa como “Otros”.
              </div>
            </CardContent>
          </Card>

          <Card className={UI.card + " lg:col-span-6"}>
            <CardHeader>
              <CardTitle className={UI.title}>Satisfacción (CSAT) promedio por Año</CardTitle>
            </CardHeader>
            <CardContent className="h-[320px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={series.csatByYear} margin={{ left: 8, right: 8 }}>
                  <CartesianGrid stroke={UI.grid} strokeDasharray="3 3" />
                  <XAxis dataKey="year" tick={{ fontSize: 12 }} />
                  <YAxis tick={{ fontSize: 12 }} domain={[0, 5]} />
                  <Tooltip
                    formatter={(v: any, name: any, props: any) => {
                      const p = props && props.payload ? props.payload : {};
                      const avg = v == null ? null : Number(v);
                      const n = Number(p.responses) || 0;
                      return [avg == null ? "—" : avg.toFixed(2), `${name} (n=${n})`];
                    }}
                  />
                  <Legend />
                  <Bar dataKey="csatAvg" name="CSAT promedio" fill={UI.primary} />
                </BarChart>
              </ResponsiveContainer>
              <div className="mt-2 text-xs text-slate-500">
                Promedio calculado solo con tickets que tienen calificación de satisfacción.
              </div>
            </CardContent>
          </Card>

          <Card className={UI.card + " lg:col-span-6"}>
            <CardHeader>
              <CardTitle className={UI.title}>Top 10 Asignados</CardTitle>
            </CardHeader>
            <CardContent className="h-[320px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={series.topAssignees} layout="vertical" margin={{ left: 110, right: 8 }}>
                  <CartesianGrid stroke={UI.grid} strokeDasharray="3 3" />
                  <XAxis type="number" tick={{ fontSize: 12 }} />
                  <YAxis type="category" dataKey="name" tick={{ fontSize: 12 }} width={100} />
                  <Tooltip />
                  <Bar dataKey="tickets" name="Tickets" fill={UI.primary} />
                </BarChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>

          <Card className={UI.card + " lg:col-span-12"}>
            <CardHeader>
              <CardTitle className={UI.title}>Heatmap Mes vs Estado (tabla)</CardTitle>
              <div className="mt-1 text-xs text-slate-500">Mostrando últimos 6 meses: {series.heatMap.range}</div>
            </CardHeader>
            <CardContent>
              <div className="overflow-auto rounded-xl border border-slate-200 bg-white">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 bg-white">
                    <tr>
                      <th className="text-left p-2 border-b">Mes</th>
                      {series.heatMap.states.map((s: string) => (
                        <th key={s} className="text-right p-2 border-b whitespace-nowrap">
                          {s}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {series.heatMap.rows.map((r: any) => (
                      <tr key={r.month} className="hover:bg-slate-50">
                        <td className="p-2 border-b whitespace-nowrap">{r.month}</td>
                        {series.heatMap.states.map((s: string) => (
                          <td key={s} className="p-2 border-b text-right">
                            {formatInt(r[s] || 0)}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="mt-2 text-xs text-slate-500">
                Sugerencia: filtra por organización o asignado para ver patrones de backlog.
              </div>
            </CardContent>
          </Card>

          <Card className={UI.card + " lg:col-span-12"}>
            <CardHeader>
              <CardTitle className={UI.title}>Heatmap Horario (tickets por hora)</CardTitle>
              <div className="mt-1 text-xs text-slate-500">
                Global (según filtros) • Agrupa por hora usando "Creada" (ej: 15:09, 15:45 → 15 hrs)
              </div>
            </CardHeader>
            <CardContent>
              <div className="overflow-auto rounded-xl border border-slate-200 bg-white">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 bg-white">
                    <tr>
                      <th className="text-left p-2 border-b whitespace-nowrap">Hora</th>
                      {series.hourHeatMap.hours.map((h: number) => (
                        <th key={h} className="text-center p-2 border-b whitespace-nowrap">
                          {String(h).padStart(2, "0")}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td className="p-2 border-b whitespace-nowrap text-slate-700">Tickets</td>
                      {series.hourHeatMap.data.map((d: any) => {
                        const style = heatBg(d.tickets, series.hourHeatMap.max);
                        return (
                          <td
                            key={d.hour}
                            className="p-2 border-b text-center font-medium"
                            style={style}
                            title={`${String(d.hour).padStart(2, "0")}:00 → ${formatInt(d.tickets)} tickets`}
                          >
                            {d.tickets ? formatInt(d.tickets) : ""}
                          </td>
                        );
                      })}
                    </tr>
                  </tbody>
                </table>
              </div>
              <div className="mt-2 text-xs text-slate-500">
                Escala: blanco = 0/bajo • azul más oscuro = mayor concentración.
              </div>
            </CardContent>
          </Card>

          <Card className={UI.card + " lg:col-span-12"}>
            <CardHeader>
              <CardTitle className={UI.title}>Heatmap Semana (día vs hora)</CardTitle>
              <div className="mt-1 text-xs text-slate-500">
                Global (según filtros) • Columnas = días • Filas = horas (0–23) • Basado en "Creada"
              </div>
            </CardHeader>
            <CardContent>
              <div className="overflow-auto rounded-xl border border-slate-200 bg-white">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 bg-white">
                    <tr>
                      <th className="text-left p-2 border-b whitespace-nowrap">Hora</th>
                      {series.weekHeatMap.days.map((d: string) => (
                        <th key={d} className="text-center p-2 border-b whitespace-nowrap">
                          {d}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {series.weekHeatMap.matrix.map((row: any) => (
                      <tr key={row.hour} className="hover:bg-slate-50">
                        <td className="p-2 border-b whitespace-nowrap text-slate-700">
                          {String(row.hour).padStart(2, "0")}:00
                        </td>
                        {series.weekHeatMap.days.map((d: string) => {
                          const v = row[d] || 0;
                          const style = heatBg(v, series.weekHeatMap.max);
                          return (
                            <td
                              key={d}
                              className="p-2 border-b text-center font-medium"
                              style={style}
                              title={`${d} ${String(row.hour).padStart(2, "0")}:00 → ${formatInt(v)} tickets`}
                            >
                              {v ? formatInt(v) : ""}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="mt-2 text-xs text-slate-500">
                Escala: blanco = 0/bajo • azul más oscuro = mayor concentración. Útil para detectar horas pico por día.
              </div>
            </CardContent>
          </Card>
        </div>

        <div className="mt-8 text-xs text-slate-500">
          Consejo: Si el CSV cambia nombres de columnas, ajusta los keys usados en el parser ("creada", "estado",
          "persona asignada", y los campos personalizados).
        </div>
      </div>
    </div>
  );
}
