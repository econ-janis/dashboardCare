/**
 * Parseo y formato de los CSV de Jira y Janis (order-report), y los tipos y
 * helpers de fecha/formato que dependen de ese parseo.
 *
 * Extraído de App.tsx para que tanto el dashboard principal como cualquier
 * otra vista (ej. /agentPerformance) reutilicen exactamente la misma lógica
 * de parseo en vez de reimplementarla.
 *
 * Reglas importantes (documentadas originalmente en App.tsx):
 * - Fechas: 19/ene/26 12:47 PM (meses en español)
 * - SLA Response: Cumplido si valor >= 0 o vacío. Incumplido solo si valor < 0.
 *   (Incluye 0, 0:00, 00:00 como Cumplido.)
 * - Columna SLA a considerar: "Campo personalizado (Time to first response)" (con o sin punto final).
 * - Organizaciones: basarse en "Campo personalizado (Organizations)"
 * - Excluir estados Block/Hold del conteo
 * - Dotación: 5 personas (Jun-2024 a Jun-2025), 3 personas (Jul-2025+)
 */

export type Row = {
  key: string;
  organization: string;
  estado: string;
  asignado: string;
  linkedKeys: string[];
  creada: Date;
  year: number;
  month: string;
  slaResponseHours: number | null;
  slaResponseStatus: "Cumplido" | "Incumplido";
  satisfaction: number | null;
  // Columna "Cuenta Janis" del CSV de Jira. "1 - PMC" identifica clientes
  // con Plan de Soporte Contratado (ver isPmcAccount en App.tsx).
  cuentaJanis: string;
};

export type JanisRow = {
  clientCode: string;
  month: string; // YYYY-MM
  year: number;
  totalOrders: number;
};

// Mapeo manual, configurable en Settings: Organizacion de Jira (clave
// normalizada) -> lista de Clientes de Janis Data (nombres tal cual
// aparecen en el order-report) asociados a esa organizacion.
export type OrgMapping = Record<string, string[]>;

export function coalesce(a: any, b: any) {
  return a === null || a === undefined ? b : a;
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

export function parseCreated(dateStr: string) {
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

export function ym(date: Date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

export function teamSizeForMonth(monthStr: string) {
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

export function parseSlaHours(s: any) {
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

export function pct(n: number, d: number) {
  if (!d) return 0;
  return (n / d) * 100;
}

export function formatInt(n: any) {
  return new Intl.NumberFormat("es-CL").format(Number(n) || 0);
}

export function formatPct(n: any) {
  const val = Number(n) || 0;
  return `${val.toFixed(2)}%`;
}

const MONTH_SHORT_NAMES = [
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

export function monthShortName(m: string) {
  if (!m || !m.includes("-")) return m;
  const [, mm] = m.split("-");
  const idx = Number(mm) - 1;
  return idx >= 0 && idx < MONTH_SHORT_NAMES.length ? MONTH_SHORT_NAMES[idx] : mm;
}

export function monthLabel(m: string) {
  // m = YYYY-MM
  if (!m || !m.includes("-")) return m;
  const [y] = m.split("-");
  return `${monthShortName(m)} ${y}`;
}

export function getField(row: Record<string, any>, candidates: string[]) {
  for (const c of candidates) {
    const v = row[c];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  for (const c of candidates) {
    if (Object.prototype.hasOwnProperty.call(row, c)) return row[c];
  }
  return undefined;
}

export function normalizeOrgKey(value: string) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

export type Shift = "Mañana" | "Tarde" | "Guardia";

// Turnos: Mañana 06:00–14:00 y Tarde 14:00–23:00 de lunes a viernes;
// Guardia es el resto del horario en día de semana (23:00–06:00) y
// sábado/domingo completo (24 hs).
export function classifyShift(d: Date): Shift {
  const day = d.getDay(); // 0=domingo, 6=sábado
  const hour = d.getHours();
  const isWeekend = day === 0 || day === 6;
  if (isWeekend) return "Guardia";
  if (hour >= 6 && hour < 14) return "Mañana";
  if (hour >= 14 && hour < 23) return "Tarde";
  return "Guardia";
}

export function isNormalSchedule(d: Date) {
  const day = d.getDay(); // 0=dom, 6=sáb
  const hour = d.getHours();
  const isWeekday = day >= 1 && day <= 5;
  return isWeekday && hour >= 6 && hour < 23;
}

// Resultado de parsear el CSV de Jira: filas válidas + cuántas se
// descartaron por no tener una fecha "Creada" interpretable.
export type ParsedJiraCsv = {
  rows: Row[];
  badDate: number;
};

// Arma las filas tipadas `Row[]` a partir del resultado crudo de
// Papa.parse (header lowercased/trimmed vía transformHeader) para el CSV
// de Jira. Misma lógica que usaba originalmente App.tsx en el `onFile`.
export function parseJiraCsvRows(data: any[]): ParsedJiraCsv {
  const parsed: Row[] = [];
  let badDate = 0;

  for (const r of data || []) {
    if (!r) continue;
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

    const cuentaJanis = String(
      coalesce(
        getField(r, [
          "campo personalizado (cuenta janis)",
          "cuenta janis",
        ]),
        ""
      )
    ).trim();

    const estado = String(coalesce(r["estado"], "")).trim();
    // Excluir Block/Hold
    if (/\b(block|hold)\b/i.test(estado)) continue;

    const linkedColumns = Object.keys(r).filter((k) => {
      const key = String(k || "").toLowerCase();
      return (
        key.includes("actividades vinculadas") ||
        key.includes("actividad vinculada") ||
        key.includes("linked activit") ||
        key.includes("enlace a la incidencia") ||
        key.includes("enlace de incidencia")
      );
    });

    const linkedMatches = linkedColumns.flatMap((col) => {
      const raw = String(coalesce(r[col], "")).trim();
      if (!raw) return [] as string[];
      const matches = raw.match(/\bHDI-\d+\b/gi);
      return matches ? matches : [];
    });

    const linkedKeys = Array.from(
      new Set(linkedMatches.map((x) => String(x).toUpperCase().trim()).filter(Boolean))
    );

    parsed.push({
      key: String(coalesce(r["clave de incidencia"], coalesce(r["key"], ""))).trim(),
      organization: org,
      estado,
      asignado: String(coalesce(r["persona asignada"], "")).trim(),
      linkedKeys,
      creada,
      year: creada.getFullYear(),
      month: ym(creada),
      slaResponseHours: slaResp,
      slaResponseStatus: respStatus,
      satisfaction: sat,
      cuentaJanis,
    });
  }

  parsed.sort((a, b) => a.creada.getTime() - b.creada.getTime());
  return { rows: parsed, badDate };
}

// Arma las filas tipadas `JanisRow[]` a partir del resultado crudo de
// Papa.parse para el CSV de Janis (order-report). Misma lógica que usaba
// originalmente App.tsx en el `onJanisFile`.
export function parseJanisCsvRows(data: any[]): JanisRow[] {
  const parsed: JanisRow[] = [];
  for (const raw of data || []) {
    const clientCode = String(coalesce(raw?.clientCode, "")).trim();
    const monthNum = Number(String(coalesce(raw?.month, "")).trim());
    const yearNum = Number(String(coalesce(raw?.year, "")).trim());
    const totalOrdersNum = Number(String(coalesce(raw?.totalOrders, "")).trim());

    if (!clientCode || !Number.isFinite(monthNum) || !Number.isFinite(yearNum) || !Number.isFinite(totalOrdersNum)) {
      continue;
    }
    if (monthNum < 1 || monthNum > 12) continue;

    const month = `${yearNum}-${String(monthNum).padStart(2, "0")}`;
    parsed.push({
      clientCode,
      month,
      year: yearNum,
      totalOrders: totalOrdersNum,
    });
  }
  return parsed;
}
