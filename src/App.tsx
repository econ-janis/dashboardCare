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
 * - Excluir estados Block/Hold del conteo
 * - Dotación: 5 personas (Jun-2024 a Jun-2025), 3 personas (Jul-2025+)
 */

// --- UI (estilo similar al screenshot) ---
const UI = {
  pageBg: "bg-[#eef2f7]",
  card: "bg-white border border-slate-200 rounded-xl shadow-none",
  title: "text-sm font-semibold text-slate-700",
  subtle: "text-xs text-slate-500",
  primary: "#2563eb", // azul principal (no-SLA y SLA cumplido)
  primaryLight: "#60a5fa",
  warning: "#f59e0b", // naranjo (SLA incumplido)
  danger: "#ef4444",
  ok: "#22c55e",
  grid: "#e5e7eb",
};

const PIE_COLORS = [
  "#2563eb",
  "#60a5fa",
  "#1d4ed8",
  "#93c5fd",
  "#3b82f6",
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
  const monthIndex = months.indexOf(
    mon[0].toUpperCase() + mon.slice(1).toLowerCase()
  );
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

function monthDeltaPct(current: number, previous: number) {
  if (!Number.isFinite(previous) || previous === 0) return null;
  return ((current - previous) / previous) * 100;
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
    .replace(/\"/g, "&quot;")
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

function toTitleCaseWords(s: string) {
  return String(s || "")
    .toLocaleLowerCase("es")
    .replace(/\p{L}[\p{L}\p{N}'-]*/gu, (w) => w.charAt(0).toLocaleUpperCase("es") + w.slice(1));
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

function pieTooltipFormatterFactory(
  data: Array<{ name: string; tickets: number }>
) {
  return (value: any, _name: any, props: any) => {
    const v = Number(value) || 0;
    const total = (data || []).reduce(
      (s, x) => s + (Number(x.tickets) || 0),
      0
    );
    const p = total ? (v / total) * 100 : 0;
    const label =
      props && props.payload && props.payload.name ? props.payload.name : "";
    return [`${formatInt(v)} (${p.toFixed(2)}%)`, label];
  };
}

async function exportExecutivePdfDirect(args: {
  html: string;
  filename: string;
}) {
  let iframe: HTMLIFrameElement | null = null;

  try {
    const [{ default: html2canvas }, { jsPDF }] = await Promise.all([
      import("html2canvas"),
      import("jspdf"),
    ]);

    iframe = document.createElement("iframe");
    iframe.style.position = "fixed";
    iframe.style.left = "-10000px";
    iframe.style.top = "0";
    iframe.style.width = "794px";
    iframe.style.height = "1123px";
    iframe.style.border = "0";
    iframe.style.background = "white";
    document.body.appendChild(iframe);

    const doc = iframe.contentDocument;
    if (!doc) throw new Error("No se pudo inicializar el documento para exportar.");

    doc.open();
    doc.write(args.html);
    doc.close();

    await new Promise<void>((resolve) => setTimeout(resolve, 300));

    try {
      // @ts-ignore
      if (doc.fonts && doc.fonts.ready) {
        // @ts-ignore
        await doc.fonts.ready;
      }
    } catch {
      // ignore
    }

    const target = doc.documentElement;

    const canvas = await html2canvas(target, {
      scale: 2,
      useCORS: true,
      backgroundColor: "#ffffff",
      windowWidth: 794,
    });

    const pdf = new jsPDF({ orientation: "p", unit: "mm", format: "a4" });

    const pageWidth = 210;
    const pageHeight = 297;
    const margin = 8;
    const contentW = pageWidth - margin * 2;
    const contentH = pageHeight - margin * 2;

    const pxPerMm = canvas.width / contentW;
    const slicePx = Math.floor(contentH * pxPerMm);

    let sy = 0;
    let page = 0;
    while (sy < canvas.height) {
      if (page > 0) pdf.addPage();

      const sh = Math.min(slicePx, canvas.height - sy);
      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = canvas.width;
      pageCanvas.height = sh;
      const ctx = pageCanvas.getContext("2d");
      if (!ctx) break;

      ctx.drawImage(canvas, 0, sy, canvas.width, sh, 0, 0, canvas.width, sh);

      const pageImg = pageCanvas.toDataURL("image/jpeg", 0.92);
      const imgH = (sh * contentW) / canvas.width;

      pdf.addImage(pageImg, "JPEG", margin, margin, contentW, imgH);

      sy += sh;
      page += 1;
    }

    const blob: Blob = pdf.output("blob");
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = args.filename;
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(() => URL.revokeObjectURL(url), 5000);
  } catch (e: any) {
    console.error("Export PDF failed", e);
    const msg = (e && (e.message || e.toString())) || "Error exportando PDF";
    throw new Error(msg);
  } finally {
    if (iframe && iframe.parentNode) iframe.parentNode.removeChild(iframe);
  }
}

function buildExecutiveReportHtml(args: {
  title: string;
  generatedAt: Date;
  filters: {
    fromMonth: string;
    toMonth: string;
    org: string;
    assignee: string;
    status: string;
  };
  autoRange: { minMonth: string | null; maxMonth: string | null };
  executive: {
    monthLabel: string;
    prevMonthLabel: string;
    insights: string[];
    metrics: Array<{
      label: string;
      value: string;
      mom: number | null;
      status: "good" | "warn" | "bad" | "neutral";
    }>;
  };
}) {
  const { title, generatedAt, filters, autoRange, executive } = args;
  const f = (v: any) => escapeHtml(v);

  const gen = generatedAt;
  const genStr = `${gen.getFullYear()}-${String(gen.getMonth() + 1).padStart(2, "0")}-${String(
    gen.getDate()
  ).padStart(2, "0")} ${String(gen.getHours()).padStart(2, "0")}:${String(
    gen.getMinutes()
  ).padStart(2, "0")}`;

  const filterLine = [
    `Archivo: ${f(autoRange.minMonth || "—")} → ${f(autoRange.maxMonth || "—")}`,
    `Vista: ${f(filters.fromMonth)} → ${f(filters.toMonth)}`,
    `Org: ${f(filters.org)}`,
    `Asignado: ${f(filters.assignee)}`,
    `Estado: ${f(filters.status)}`,
  ].join(" • ");

  const momText = (v: number | null) => {
    if (v == null || !Number.isFinite(v)) return "Sin comparativo";
    const sign = v > 0 ? "+" : "";
    return `${sign}${v.toFixed(1)}% vs mes anterior`;
  };

  const statusClass = (s: "good" | "warn" | "bad" | "neutral") => `dot ${s}`;

  const css = `
  @page { size: A4; margin: 12mm; }
  body { 
    font-family: 'Inter', -apple-system, sans-serif; 
    color: #1e293b; 
    background-color: white; 
  }
  h1 { font-size: 20px; color: #0f172a; margin-bottom: 4px; }
  h2 { font-size: 14px; color: #be185d; margin-bottom: 8px; margin-top: 0; }
  .meta { font-size: 10px; color: #64748b; margin-bottom: 14px; border-bottom: 1px solid #e2e8f0; padding-bottom: 10px; }
  .badge { padding: 2px 6px; border-radius: 4px; background: #fce7f3; color: #9d174d; font-weight: 700; font-size: 9px; }
  .block { border: 1px solid #f5d0fe; border-radius: 10px; padding: 12px; margin-bottom: 12px; background: #fdf2f8; }
  .kpi-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  .kpi { background: #fff; border: 1px solid #fbcfe8; border-radius: 8px; padding: 8px; }
  .kpi .head { display: flex; align-items: center; gap: 6px; margin-bottom: 4px; }
  .dot { width: 10px; height: 10px; border-radius: 9999px; display:inline-block; }
  .dot.good { background:#16a34a; }
  .dot.warn { background:#f59e0b; }
  .dot.bad { background:#dc2626; }
  .dot.neutral { background:#64748b; }
  .kpi .label { font-size: 9px; font-weight: 700; color: #475569; text-transform: uppercase; }
  .kpi .value { font-size: 15px; font-weight: 700; color: #0f172a; }
  .kpi .mom { font-size: 10px; color: #64748b; }
  ul { margin: 6px 0 0 18px; padding: 0; }
  li { margin: 5px 0; font-size: 11px; }
  .subtle { font-size: 10px; color: #64748b; }
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

      <div class="block">
        <h2>1. Resumen Ejecutivo</h2>
        <div class="subtle">Período actual: ${f(executive.monthLabel)} · comparado con ${f(
    executive.prevMonthLabel
  )}</div>
        <div class="kpi-grid" style="margin-top:8px;">
          ${executive.metrics
            .map(
              (k) => `<div class="kpi">
                <div class="head"><span class="${statusClass(k.status)}"></span><span class="label">${f(
                k.label
              )}</span></div>
                <div class="value">${f(k.value)}</div>
                <div class="mom">${f(momText(k.mom))}</div>
              </div>`
            )
            .join("")}
        </div>
        <ul>${executive.insights.map((i) => `<li>${f(i)}</li>`).join("")}</ul>
      </div>

      <div class="block">
        <h2>2. Performance Operativa</h2>
        <div class="subtle">Lectura rápida de volumen, velocidad de resolución y cumplimiento SLA para decisiones operativas.</div>
      </div>

      <div class="block">
        <h2>3. Calidad / Impacto</h2>
        <div class="subtle">Seguimiento de reaperturas, estabilidad de servicio y señales de riesgo para la experiencia del cliente.</div>
      </div>

      <div class="block">
        <h2>4. Plan de Acción</h2>
        <ul>
          <li>Priorizar focos de backlog y reaperturas con objetivos de reducción para el próximo mes.</li>
          <li>Definir acciones concretas para sostener (o recuperar) el cumplimiento SLA.</li>
          <li>Alinear capacidad del equipo según el comportamiento de demanda observado.</li>
        </ul>
      </div>
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
            <div className="text-2xl font-semibold tracking-tight text-slate-900">
              {value}
            </div>
            {badge ? <div className="mt-2">{badge}</div> : null}
            {subtitle ? <div className={"mt-1 " + UI.subtle}>{subtitle}</div> : null}
          </div>
          {right ? <div className={"text-right " + UI.subtle}>{right}</div> : null}
        </div>
      </CardContent>
    </Card>
  );
}

function HealthBadge({ label, color }: { label: string; color: string }) {
  return (
    <span
      className="inline-flex items-center rounded-full px-2.5 py-1 text-xs font-semibold"
      style={{ backgroundColor: "#f1f5f9", color }}
    >
      {label}
    </span>
  );
}

function YearBars({
  rows,
  maxTickets,
}: {
  rows: Array<{ year: string; tickets: number; partialLabel: string }>;
  maxTickets: number;
}) {
  return (
    <div className="space-y-3">
      {rows.map((r) => {
        const pctW = maxTickets
          ? Math.max(0, Math.min(100, (r.tickets / maxTickets) * 100))
          : 0;
        return (
          <div key={r.year} className="flex items-center gap-3">
            <div className="w-12 text-sm text-slate-700 font-semibold">{r.year}</div>
            <div className="flex-1">
              <div className="h-3 rounded-full bg-slate-100 overflow-hidden">
                <div
                  className="h-3 rounded-full"
                  style={{ width: `${pctW}%`, backgroundColor: UI.primary }}
                />
              </div>
            </div>
            <div className="w-40 text-right text-sm text-slate-700">
              <span className="font-semibold">{formatInt(r.tickets)}</span>
              <span className="text-xs text-slate-500">{r.partialLabel}</span>
            </div>
          </div>
        );
      })}
    </div>
  );
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
  console.assert((parseSlaHours("1:30") || 0) > 0, "positive SLA should be > 0");
  console.assert((parseSlaHours("-1:30") || 0) < 0, "negative SLA should be < 0");

  const arr: string[] = [];
  const min = arr.length ? arr[0] : undefined;
  console.assert(min === undefined, "safe min when empty");
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
  const [exporting, setExporting] = useState(false);
  const [showExecutiveReport, setShowExecutiveReport] = useState(false);

  // Filters: rango por mes (YYYY-MM)
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
            // Excluir Block/Hold
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
              "No pude parsear filas con fecha 'Creada'. Revisa que el CSV tenga columna 'Creada' y formato tipo 19/ene/26 12:47 PM."
            );
            return;
          }

          parsed.sort((a, b) => a.creada.getTime() - b.creada.getTime());
          setRows(parsed);

          const minMonth = parsed[0].month;
          const maxMonth = parsed[parsed.length - 1].month;
          setAutoRange({ minMonth, maxMonth });
          setFromMonth(minMonth);
          setToMonth(maxMonth);

          if (badDate > 0) {
            setError(`Aviso: ${badDate} filas fueron omitidas porque la fecha 'Creada' no era interpretable.`);
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

  const minMonthBound =
    autoRange.minMonth ?? (filterOptions.months.length ? filterOptions.months[0] : undefined);
  const maxMonthBound =
    autoRange.maxMonth ??
    (filterOptions.months.length ? filterOptions.months[filterOptions.months.length - 1] : undefined);

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
        ? rated.reduce((s, r) => s + (r.satisfaction == null ? 0 : r.satisfaction), 0) / rated.length
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
      if (!monthsSorted.length) return [] as string[];
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

    // Tickets por año
    const byYear = new Map<string, { year: string; tickets: number }>();
    for (const r of filtered) {
      const y = String(r.year);
      const cur = byYear.get(y) || { year: y, tickets: 0 };
      cur.tickets += 1;
      byYear.set(y, cur);
    }
    const ticketsByYear = Array.from(byYear.values()).sort((a, b) => Number(a.year) - Number(b.year));

    // Estado por año
    const yearStatus = new Map<string, any>();
    for (const r of filtered) {
      const y = String(r.year);
      const obj = yearStatus.get(y) || { year: y };
      const s = r.estado || "(Sin estado)";
      obj[s] = (obj[s] || 0) + 1;
      yearStatus.set(y, obj);
    }
    const estadoByYear = Array.from(yearStatus.values()).sort((a, b) => Number(a.year) - Number(b.year));

    // SLA response por año
    const slaYear = new Map<string, any>();
    for (const r of filtered) {
      const y = String(r.year);
      const obj =
        slaYear.get(y) ||
        ({ year: y, Total: 0, Cumplido: 0, Incumplido: 0, CumplidoPct: 0, IncumplidoPct: 0 } as any);
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

    // Count helper
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

    // Pie top 5 orgs + otros
    const allOrgs = count((r) => r.organization);
    const top5 = allOrgs.slice(0, 5);
    const top5Sum = top5.reduce((s, x) => s + x.tickets, 0);
    const others = Math.max(0, totalTickets - top5Sum);
    const topOrgsPie = [...top5];
    if (others > 0) topOrgsPie.push({ name: "Otros", tickets: others });

    // CSAT por año
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
      .map((x) => ({ year: x.year, csatAvg: x.cnt ? x.sum / x.cnt : null, responses: x.cnt }))
      .sort((a, b) => Number(a.year) - Number(b.year));

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

    // Heatmap por hora
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
      return { data, max };
    })();

    // Heatmap semana (día vs hora)
    const weekHeatMap = (() => {
      const days = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]; // ISO
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
      for (const row of matrix) for (const d of days) max = Math.max(max, row[d] || 0);

      return { days, matrix, max };
    })();

    return {
      ticketsByMonth,
      ticketsByYear,
      estadoByYear,
      slaByYear,
      csatByYear,
      topAssignees,
      topOrgsPie,
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
    const isPartialYear = !!maxCreated && !(maxCreated.getMonth() === 11 && maxCreated.getDate() === 31);

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

  const heatMaxMonthState = useMemo(() => {
    let max = 0;
    for (const r of series.heatMap.rows) {
      for (const s of series.heatMap.states) max = Math.max(max, Number(r[s] || 0));
    }
    return max;
  }, [series.heatMap]);

  const executiveReportData = useMemo(() => {
    const monthsSorted = Array.from(new Set(filtered.map((r) => r.month))).sort();
    const currentMonth = monthsSorted.length ? monthsSorted[monthsSorted.length - 1] : null;
    const previousMonth = monthsSorted.length > 1 ? monthsSorted[monthsSorted.length - 2] : null;

    const currentRows = currentMonth ? filtered.filter((r) => r.month === currentMonth) : [];
    const previousRows = previousMonth ? filtered.filter((r) => r.month === previousMonth) : [];

    const closedStatuses = new Set([
      "done",
      "closed",
      "resuelto",
      "resuelta",
      "solucionado",
      "solucionada",
      "resuelto/a",
      "completado",
      "completada",
    ]);

    const canceledStatuses = new Set([
      "cancelado",
      "cancelada",
      "cancelled",
      "canceled",
      "anulado",
      "anulada",
    ]);

    const withoutCanceled = (rowsSubset: Row[]) =>
      rowsSubset.filter((r) => !canceledStatuses.has(String(r.estado || "").trim().toLowerCase()));

    const resolvedCount = (rowsSubset: Row[]) =>
      rowsSubset.filter((r) => closedStatuses.has(String(r.estado || "").trim().toLowerCase())).length;

    const backlogCount = (rowsSubset: Row[]) =>
      rowsSubset.filter((r) => !closedStatuses.has(String(r.estado || "").trim().toLowerCase())).length;

    const backlogByStatus = (rowsSubset: Row[]) => {
      const map = new Map<string, { count: number; keys: string[] }>();
      rowsSubset.forEach((r) => {
        const statusRaw = String(r.estado || "").trim();
        const statusKey = statusRaw.toLowerCase();
        if (!statusRaw || closedStatuses.has(statusKey)) return;
        const current = map.get(statusRaw) || { count: 0, keys: [] };
        current.count += 1;
        if (r.key) current.keys.push(r.key);
        map.set(statusRaw, current);
      });
      return Array.from(map.entries())
        .map(([status, data]) => ({
          status: toTitleCaseWords(status),
          count: data.count,
          keys: data.keys.sort((a, b) => a.localeCompare(b)),
        }))
        .sort((a, b) => b.count - a.count || a.status.localeCompare(b.status));
    };

    const currentRowsNoCanceled = withoutCanceled(currentRows);
    const previousRowsNoCanceled = withoutCanceled(previousRows);

    const ticketsCurrent = currentRowsNoCanceled.length;
    const ticketsPrev = previousRowsNoCanceled.length;
    const resolvedCurrent = resolvedCount(currentRowsNoCanceled);
    const resolvedPrev = resolvedCount(previousRowsNoCanceled);
    const slaCurrent =
      100 - pct(currentRowsNoCanceled.filter((r) => r.slaResponseStatus === "Incumplido").length, ticketsCurrent);
    const slaPrev =
      100 - pct(previousRowsNoCanceled.filter((r) => r.slaResponseStatus === "Incumplido").length, ticketsPrev);
    const backlogCurrent = backlogCount(currentRowsNoCanceled);
    const backlogPrev = backlogCount(previousRowsNoCanceled);
    const backlogStatusCurrent = backlogByStatus(currentRowsNoCanceled);

    const metricStatus = (metric: string, value: number) => {
      if (!Number.isFinite(value)) return "neutral" as const;
      if (metric === "sla") {
        if (value >= 95) return "good" as const;
        if (value >= 90) return "warn" as const;
        return "bad" as const;
      }
      if (metric === "backlog") {
        if (value <= 25) return "good" as const;
        if (value <= 60) return "warn" as const;
        return "bad" as const;
      }
      return "neutral" as const;
    };

    const safeInsights = (() => {
      if (!currentMonth) {
        return [
          "No hay suficientes datos filtrados para construir insights del mes.",
          "Carga un CSV y selecciona un cliente para ver comparativos mensuales.",
          "La sección prioriza conclusiones para acelerar decisiones ejecutivas.",
        ];
      }

      const ticketsMom = monthDeltaPct(ticketsCurrent, ticketsPrev);
      const backlogMom = monthDeltaPct(backlogCurrent, backlogPrev);

      const momSummary = (v: number | null, up: string, down: string) => {
        if (v == null) return "sin base comparativa";
        if (Math.abs(v) < 0.05) return "sin variación";
        if (v > 0) return `${up} ${v.toFixed(1)}%`;
        return `${down} ${Math.abs(v).toFixed(1)}%`;
      };

      return [
        `Volumen de tickets ${momSummary(ticketsMom, "al alza", "a la baja")} en ${monthLabel(currentMonth)}.`,
        `Cumplimiento SLA ${slaCurrent >= 95 ? "estable" : "en riesgo"} en ${slaCurrent.toFixed(1)}%, foco en continuidad operativa.`,
        `Backlog ${momSummary(backlogMom, "aumenta", "disminuye")} y requiere foco por estado operativo.`,
        "Se concentran esfuerzos para entregar esos desarrollos dentro de la semana.",
      ];
    })();

    return {
      monthLabel: currentMonth ? monthLabel(currentMonth) : "Sin datos",
      prevMonthLabel: previousMonth ? monthLabel(previousMonth) : "Sin mes anterior",
      metrics: [
        {
          label: "🎫 Tickets recibidos",
          value: formatInt(ticketsCurrent),
          mom: monthDeltaPct(ticketsCurrent, ticketsPrev),
          status: "neutral" as const,
        },
        {
          label: "✅ Tickets resueltos",
          value: formatInt(resolvedCurrent),
          mom: monthDeltaPct(resolvedCurrent, resolvedPrev),
          status: "neutral" as const,
        },
        {
          label: "⏱️ SLA cumplimiento",
          value: formatPct(slaCurrent),
          mom: monthDeltaPct(slaCurrent, slaPrev),
          status: metricStatus("sla", slaCurrent),
        },
        {
          label: "🔴 Backlog al cierre",
          value: formatInt(backlogCurrent),
          mom: monthDeltaPct(backlogCurrent, backlogPrev),
          status: metricStatus("backlog", backlogCurrent),
        },
      ],
      backlogByStatus: backlogStatusCurrent,
      resolvedMom: monthDeltaPct(resolvedCurrent, resolvedPrev),
      backlogMom: monthDeltaPct(backlogCurrent, backlogPrev),
      insights: safeInsights,
    };
  }, [filtered]);

  const clearAll = () => {
    setRows([]);
    setError(null);
    setFromMonth("all");
    setToMonth("all");
    setOrgFilter("all");
    setAssigneeFilter("all");
    setStatusFilter("all");
    setAutoRange({ minMonth: null, maxMonth: null });
  };

  return (
    <div className={`min-h-screen ${UI.pageBg} p-4 md:p-8`}>
      <div className="mx-auto max-w-7xl">
        <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
          <div>
            <h1 className="text-2xl md:text-3xl font-semibold tracking-tight">
              Janis Commerce -  Care Executive Dashboard
            </h1>
            <p className="text-sm text-slate-500 mt-1">
              Sube tu CSV y verás KPIs y gráficos. Regla SLA Response (Time to first response): &gt;= 0 (o vacío) =
              cumplido; &lt; 0 = incumplido. Dotación: 5 (Jun-2024 a Jun-2025), 3 (Jul-2025+).
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
              disabled={exporting || !filtered.length}
              onClick={async () => {
                setExporting(true);
                setError(null);

                try {
                  if (!filtered.length) {
                    setError("No hay datos filtrados para exportar.");
                    return;
                  }

                  setError("Generando PDF…");

                  const now = new Date();
                  const y = now.getFullYear();
                  const m = String(now.getMonth() + 1).padStart(2, "0");
                  const d = String(now.getDate()).padStart(2, "0");
                  const filename = `Informe_Ejecutivo_Janis_Care_${y}${m}${d}.pdf`;

                  const html = buildExecutiveReportHtml({
                    title: "Janis Commerce -  Care Executive Dashboard",
                    generatedAt: now,
                    filters: {
                      fromMonth: fromMonth === "all" ? autoRange.minMonth || "all" : fromMonth,
                      toMonth: toMonth === "all" ? autoRange.maxMonth || "all" : toMonth,
                      org: orgFilter === "all" ? "Todas" : orgFilter,
                      assignee: assigneeFilter === "all" ? "Todos" : assigneeFilter,
                      status: statusFilter === "all" ? "Todos" : statusFilter,
                    },
                    autoRange,
                    executive: executiveReportData,
                  });

                  await exportExecutivePdfDirect({ html, filename });
                  setError(null);
                } catch (e: any) {
                  console.error(e);
                  setError(
                    (e && (e.message || String(e))) ||
                      "No se pudo exportar (descarga directa). Si tu entorno no incluye html2canvas/jspdf, hay que agregarlos."
                  );
                } finally {
                  setExporting(false);
                }
              }}
            >
              {exporting ? "Exportando…" : "Exportar Informe (General)"}
            </Button>

            <Button variant="outline" onClick={clearAll}>
              Limpiar
            </Button>
          </div>
        </div>

        {error ? (
          <div className="mt-4 rounded-xl border border-slate-200 bg-white p-3 text-sm text-slate-700">
            {error}
          </div>
        ) : null}

        {/* Filters */}
        <div className="mt-6 grid grid-cols-1 gap-3 md:grid-cols-5">
          <Card className={UI.card}>
            <CardContent className="p-4">
              <div className={UI.subtle}>Desde (mes)</div>
              <Input
                type="month"
                value={fromMonth === "all" ? (autoRange.minMonth || "") : fromMonth}
                min={minMonthBound || undefined}
                max={maxMonthBound || undefined}
                onChange={(e) => setFromMonth(e.target.value || "all")}
              />
            </CardContent>
          </Card>
          <Card className={UI.card}>
            <CardContent className="p-4">
              <div className={UI.subtle}>Hasta (mes)</div>
              <Input
                type="month"
                value={toMonth === "all" ? (autoRange.maxMonth || "") : toMonth}
                min={minMonthBound || undefined}
                max={maxMonthBound || undefined}
                onChange={(e) => setToMonth(e.target.value || "all")}
              />
            </CardContent>
          </Card>

          <Card className={UI.card}>
            <CardContent className="p-4">
              <div className={UI.subtle}>Organización</div>
              <Select value={orgFilter} onValueChange={setOrgFilter}>
                <SelectTrigger>
                  <SelectValue placeholder="Todas" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">Todas</SelectItem>
                  {filterOptions.orgs.map((o) => (
                    <SelectItem key={o} value={o}>
                      {o}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </CardContent>
          </Card>

          <Card className={UI.card}>
            <CardContent className="p-4">
              <div className={UI.subtle}>Asignado</div>
              <Select value={assigneeFilter} onValueChange={setAssigneeFilter}>
                <SelectTrigger>
                  <SelectValue placeholder="Todos" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">Todos</SelectItem>
                  {filterOptions.assignees.map((a) => (
                    <SelectItem key={a} value={a}>
                      {a}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </CardContent>
          </Card>

          <Card className={UI.card}>
            <CardContent className="p-4">
              <div className={UI.subtle}>Estado</div>
              <Select value={statusFilter} onValueChange={setStatusFilter}>
                <SelectTrigger>
                  <SelectValue placeholder="Todos" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">Todos</SelectItem>
                  {filterOptions.estados.map((s) => (
                    <SelectItem key={s} value={s}>
                      {s}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </CardContent>
          </Card>
        </div>

        {/* KPIs */}
        <div className="mt-6 grid grid-cols-1 gap-3 md:grid-cols-5">
          {kpiCard("Tickets (vista)", formatInt(kpis.total))}
          {kpiCard(
            "Tickets último mes (vista)",
            formatInt(kpis.monthCount),
            kpis.latestMonth ? monthLabel(kpis.latestMonth) : "—"
          )}
          {kpiCard(
            "Cumplimiento SLA Response",
            formatPct(kpis.respOkPct),
            `${formatInt(kpis.respInc)} incumplidos`
          )}
          {kpiCard(
            "CSAT promedio (por año)",
            kpis.csatAvg == null ? "—" : kpis.csatAvg.toFixed(2),
            `Cobertura: ${formatPct(kpis.csatCoverage)}`
          )}
          {kpiCard(
            "Tickets / Persona (prom. 6 meses)",
            kpis.tpp6m == null ? "—" : kpis.tpp6m.toFixed(1),
            "(excluye mes actual si no está cerrado)",
            undefined,
            <HealthBadge label={kpis.tppHealth.label} color={kpis.tppHealth.color} />
          )}
        </div>

        {/* Charts */}
        <div className="mt-6 grid grid-cols-1 gap-3 md:grid-cols-2">
          <Card className={UI.card}>
            <CardHeader>
              <CardTitle className={UI.title}>Tickets por Mes</CardTitle>
            </CardHeader>
            <CardContent className="h-72">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={series.ticketsByMonth}>
                  <CartesianGrid stroke={UI.grid} />
                  <XAxis dataKey="month" tickFormatter={monthLabel as any} />
                  <YAxis />
                  <Tooltip labelFormatter={(l) => monthLabel(String(l))} />
                  <Line type="monotone" dataKey="tickets" stroke={UI.primary} strokeWidth={2} dot={false} />
                </LineChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>

          <Card className={UI.card}>
            <CardHeader>
              <CardTitle className={UI.title}>Tickets por Año</CardTitle>
            </CardHeader>
            <CardContent>
              <YearBars rows={ticketsByYearBars.rows} maxTickets={ticketsByYearBars.maxTickets} />
            </CardContent>
          </Card>
        </div>

        <div className="mt-6 grid grid-cols-1 gap-3 md:grid-cols-2">
          <Card className={UI.card}>
            <CardHeader>
              <CardTitle className={UI.title}>SLA Response por Año (porcentaje)</CardTitle>
            </CardHeader>
            <CardContent className="h-72">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={series.slaByYear}>
                  <CartesianGrid stroke={UI.grid} />
                  <XAxis dataKey="year" />
                  <YAxis domain={[0, 100]} tickFormatter={(v) => `${v}%`} />
                  <Tooltip
                    formatter={(v: any, n: any) => [`${Number(v).toFixed(2)}%`, n]}
                    labelFormatter={(l) => `Año ${l}`}
                  />
                  <Legend />
                  <Bar dataKey="CumplidoPct" name="Cumplido" stackId="a" fill={UI.primary} />
                  <Bar dataKey="IncumplidoPct" name="Incumplido" stackId="a" fill={UI.warning} />
                </BarChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>

          <Card className={UI.card}>
            <CardHeader>
              <CardTitle className={UI.title}>CSAT promedio por Año</CardTitle>
            </CardHeader>
            <CardContent className="h-72">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={series.csatByYear}>
                  <CartesianGrid stroke={UI.grid} />
                  <XAxis dataKey="year" />
                  <YAxis />
                  <Tooltip
                    formatter={(v: any) => [Number(v).toFixed(2), "CSAT"]}
                    labelFormatter={(l) => `Año ${l}`}
                  />
                  <Bar dataKey="csatAvg" name="CSAT" fill={UI.primary} />
                </BarChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>
        </div>

        <div className="mt-6 grid grid-cols-1 gap-3 md:grid-cols-2">
          <Card className={UI.card}>
            <CardHeader>
              <CardTitle className={UI.title}>Top 5 Organizaciones (torta) + Otros</CardTitle>
            </CardHeader>
            <CardContent className="h-80">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Tooltip formatter={pieTooltipFormatter as any} />
                  <Pie
                    data={series.topOrgsPie}
                    dataKey="tickets"
                    nameKey="name"
                    outerRadius={100}
                    innerRadius={45}
                    paddingAngle={2}
                  >
                    {series.topOrgsPie.map((_, i) => (
                      <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />
                    ))}
                  </Pie>
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>

          <Card className={UI.card}>
            <CardHeader>
              <CardTitle className={UI.title}>Top 10 Asignados</CardTitle>
            </CardHeader>
            <CardContent className="h-80">
              <ResponsiveContainer width="80%" height="100%">
                <BarChart data={series.topAssignees} layout="vertical" margin={{ left: 50 }}>
                  <CartesianGrid stroke={UI.grid} />
                  <XAxis type="number" />
                  <YAxis type="category" dataKey="name" width={100} />
                  <Tooltip formatter={(v: any) => [formatInt(v), "Tickets"]} />
                  <Bar dataKey="tickets" fill={UI.primary} />
                </BarChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>
        </div>

        {/* Heatmaps */}
        <div className="mt-6 grid grid-cols-1 gap-3">
          <Card className={UI.card}>
            <CardHeader>
              <CardTitle className={UI.title}>Heatmap Mes vs Estado (últimos 6 meses)</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-sm">
                  <thead>
                    <tr className="text-left">
                      <th className="p-2 border border-slate-200 bg-slate-50">Mes</th>
                      {series.heatMap.states.map((s) => (
                        <th key={s} className="p-2 border border-slate-200 bg-slate-50">
                          {s}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {series.heatMap.rows.map((r: any) => (
                      <tr key={r.month}>
                        <td className="p-2 border border-slate-200 font-semibold text-slate-700">
                          {monthLabel(r.month)}
                        </td>
                        {series.heatMap.states.map((s) => {
                          const v = Number(r[s] || 0);
                          const style = heatBg(v, heatMaxMonthState);
                          return (
                            <td key={s} className="p-2 border border-slate-200 text-center" style={style}>
                              {v ? formatInt(v) : ""}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </CardContent>
          </Card>

          <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
            <Card className={UI.card}>
              <CardHeader>
                <CardTitle className={UI.title}>Heatmap Horario (por hora)</CardTitle>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-6 gap-2">
                  {series.hourHeatMap.data.map((x) => (
                    <div
                      key={x.hour}
                      className="rounded-lg border border-slate-200 p-2 text-center"
                      style={heatBg(x.tickets, series.hourHeatMap.max)}
                    >
                      <div className="text-xs font-semibold">{String(x.hour).padStart(2, "0")}:00</div>
                      <div className="text-sm">{x.tickets ? formatInt(x.tickets) : ""}</div>
                    </div>
                  ))}
                </div>
              </CardContent>
            </Card>

            <Card className={UI.card}>
              <CardHeader>
                <CardTitle className={UI.title}>Heatmap Semana (día vs hora)</CardTitle>
              </CardHeader>
              <CardContent>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-sm">
                    <thead>
                      <tr className="text-left">
                        <th className="p-2 border border-slate-200 bg-slate-50">Hora</th>
                        {series.weekHeatMap.days.map((d) => (
                          <th key={d} className="p-2 border border-slate-200 bg-slate-50">
                            {d}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {series.weekHeatMap.matrix.map((row: any) => (
                        <tr key={row.hour}>
                          <td className="p-2 border border-slate-200 font-semibold text-slate-700">
                            {String(row.hour).padStart(2, "0")}:00
                          </td>
                          {series.weekHeatMap.days.map((d) => {
                            const v = Number(row[d] || 0);
                            const style = heatBg(v, series.weekHeatMap.max);
                            return (
                              <td key={d} className="p-2 border border-slate-200 text-center" style={style}>
                                {v ? formatInt(v) : ""}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </CardContent>
            </Card>
          </div>
        </div>

        <div className="mt-6 text-xs text-slate-500">
          Sugerencia: aplica enfoque Pareto 80/20 sobre Top Organizaciones/Asignados para reducir demanda recurrente.
        </div>

        <div className="mt-6">
          <Card className="rounded-xl border border-[#1e4b8f] bg-gradient-to-br from-[#001640] via-[#002862] to-[#0a3c86] text-white shadow-lg shadow-[#001640]/35">
            <CardHeader>
              <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                <CardTitle className="text-base font-semibold text-white">
                  Reporte Ejecutivo por Cliente (v1)
                </CardTitle>
                <Button
                  className="self-start border border-orange-300/50 bg-gradient-to-r from-[#ff8f2b] to-[#ff7600] text-white shadow-sm shadow-[#ff7600]/40 hover:from-[#ff9c43] hover:to-[#ff8b1f]"
                  disabled={!filtered.length}
                  onClick={() => setShowExecutiveReport((prev) => !prev)}
                >
                  {showExecutiveReport ? "Ocultar Reporte Ejecutivo" : "Generar Reporte Ejecutivo"}
                </Button>
              </div>
              <p className="text-sm text-blue-100">
                Enfoque en 4 bloques: Resumen Ejecutivo, Performance Operativa, Calidad/Impacto y Plan de Acción.
                Vista simple, visual y comparativa vs mes anterior para decisiones rápidas.
              </p>
            </CardHeader>
            <CardContent>
              {showExecutiveReport ? (
                <div className="rounded-lg border border-blue-200/60 bg-white p-4 text-slate-900">
                  <div className="mb-3">
                    <div className="text-sm font-semibold text-[#0a2f6f]">1️⃣ Resumen Ejecutivo</div>
                    <div className="text-xs text-slate-600">
                      {executiveReportData.monthLabel} vs {executiveReportData.prevMonthLabel}
                    </div>
                  </div>

                  <div className="grid grid-cols-1 gap-3 md:grid-cols-2 xl:grid-cols-3">
                    {executiveReportData.metrics
                      .filter((metric) => !metric.label.includes("resueltos") && !metric.label.includes("Backlog"))
                      .map((metric) => {
                      const dotColor =
                        metric.status === "good"
                          ? "bg-emerald-500"
                          : metric.status === "warn"
                            ? "bg-amber-500"
                            : metric.status === "bad"
                              ? "bg-red-500"
                              : "bg-slate-400";

                      const momLabel =
                        metric.mom == null
                          ? "Sin comparativo"
                          : `${metric.mom > 0 ? "+" : ""}${metric.mom.toFixed(1)}% vs mes anterior`;

                      return (
                        <div key={metric.label} className="rounded-lg border border-blue-100 bg-gradient-to-b from-blue-50 to-white p-3">
                          <div className="mb-1 flex items-center gap-2 text-xs font-semibold text-[#0a2f6f]">
                            <span className={`h-2.5 w-2.5 rounded-full ${dotColor}`} />
                            {metric.label}
                          </div>
                          <div className="text-xl font-semibold text-slate-900">{metric.value}</div>
                          <div className="text-xs text-blue-700">{momLabel}</div>
                        </div>
                      );
                      })}

                    <div className="rounded-lg border-2 border-red-500 bg-gradient-to-b from-blue-50 to-white p-3 md:col-span-2 xl:col-span-1">
                      <div className="mb-1 flex items-center gap-2 text-xs font-semibold text-[#0a2f6f]">
                        <span className="h-2.5 w-2.5 rounded-full bg-emerald-500" />
                        ✅ Tickets resueltos + 🔴 Backlog al cierre
                      </div>
                      <div className="space-y-1 text-xs text-slate-700">
                        <div>
                          <span className="font-semibold text-slate-900">Resueltos:</span>{" "}
                          {executiveReportData.metrics.find((m) => m.label.includes("resueltos"))?.value || "0"}
                          <span className="ml-2 text-blue-700">
                            {(() => {
                              const mom = executiveReportData.resolvedMom;
                              if (mom == null) return "Sin comparativo";
                              return `${mom > 0 ? "+" : ""}${mom.toFixed(1)}% vs mes anterior`;
                            })()}
                          </span>
                        </div>
                        <div>
                          <span className="font-semibold text-slate-900">Backlog:</span>{" "}
                          {executiveReportData.metrics.find((m) => m.label.includes("Backlog"))?.value || "0"}
                          <span className="ml-2 text-blue-700">
                            {(() => {
                              const mom = executiveReportData.backlogMom;
                              if (mom == null) return "Sin comparativo";
                              return `${mom > 0 ? "+" : ""}${mom.toFixed(1)}% vs mes anterior`;
                            })()}
                          </span>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="mt-4 rounded-lg border border-blue-100 bg-white p-3">
                    <div className="mb-2 text-sm font-semibold text-[#0a2f6f]">Backlog al cierre · Apertura por estados</div>
                    {executiveReportData.backlogByStatus.length ? (
                      <div className="grid grid-cols-1 gap-2 sm:grid-cols-2 lg:grid-cols-3">
                        {executiveReportData.backlogByStatus.map((item) => (
                          <div key={item.status} className="rounded-md border border-blue-100 bg-blue-50/60 px-3 py-2 text-sm">
                            <span className="font-medium text-slate-800">{item.status}</span>
                            <span className="ml-2 font-semibold text-[#0a2f6f]">{formatInt(item.count)}</span>
                            {item.keys.length ? (
                              <div className="mt-1 text-xs text-slate-600">
                                <span className="font-semibold">Keys:</span> {item.keys.join(", ")}
                              </div>
                            ) : null}
                          </div>
                        ))}
                      </div>
                    ) : (
                      <p className="text-xs text-slate-600">Sin backlog para el período seleccionado.</p>
                    )}
                  </div>

                  <div className="mt-4 grid grid-cols-1 gap-3 md:grid-cols-3 text-sm text-slate-700">
                    <div className="rounded-lg border border-blue-100 bg-blue-50/60 p-3">
                      <div className="mb-1 font-semibold text-[#0a2f6f]">2️⃣ Performance Operativa</div>
                      <p className="text-xs">Volumen, velocidad y cumplimiento SLA para decisiones de capacidad.</p>
                    </div>
                    <div className="rounded-lg border border-blue-100 bg-blue-50/60 p-3">
                      <div className="mb-1 font-semibold text-[#0a2f6f]">3️⃣ Calidad / Impacto</div>
                      <p className="text-xs">Seguimiento de reaperturas y estabilidad del servicio para reducir fricción.</p>
                    </div>
                    <div className="rounded-lg border border-blue-100 bg-blue-50/60 p-3">
                      <div className="mb-1 font-semibold text-[#0a2f6f]">4️⃣ Plan de Acción</div>
                      <p className="text-xs">Priorizar backlog, sostener SLA y ajustar capacidad del equipo.</p>
                    </div>
                  </div>

                  <div className="mt-4 rounded-lg border border-[#ffb477]/70 bg-gradient-to-r from-[#fff7ef] to-white p-3">
                    <div className="text-sm font-semibold text-[#c45d00]">Insights ejecutivos</div>
                    <ul className="mt-2 list-disc space-y-1 pl-5 text-sm text-slate-700">
                      {executiveReportData.insights.map((insight, idx) => (
                        <li key={idx}>{insight}</li>
                      ))}
                    </ul>
                  </div>
                </div>
              ) : null}
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  );
}
