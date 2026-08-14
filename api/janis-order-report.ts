// Proxy serverless (Vercel) para el order-report de Janis.
//
// Por qué existe: la API de Janis requiere credenciales (janis-client /
// janis-api-key / janis-api-secret). Si el frontend llamara a oms.janis.in
// directamente, esas credenciales quedarían visibles en el bundle JS y en
// el Network tab del navegador de cualquiera que abra el dashboard. Este
// endpoint corre server-side en Vercel, guarda las credenciales como
// variables de entorno (nunca expuestas al cliente) y devuelve al frontend
// solo el JSON ya filtrado.
//
// Variables de entorno requeridas (Vercel → Project Settings →
// Environment Variables, ambiente Production):
//   JANIS_CLIENT
//   JANIS_API_KEY
//   JANIS_API_SECRET

const JANIS_ORDER_REPORT_URL = "https://oms.janis.in/api/order-report";
// Salvaguarda ante un posible loop de paginación: 60 páginas cubre con
// margen la historia completa del reporte (hoy son ~1600 filas totales).
const MAX_PAGES = 60;

type JanisApiRow = {
  clientCode?: unknown;
  month?: unknown;
  year?: unknown;
  totalOrders?: unknown;
  totalAmount?: unknown;
  cpoTotalAmount?: unknown;
  cpoTotalOrders?: unknown;
  npoTotalAmount?: unknown;
  npoTotalOrders?: unknown;
};

function extractRows(body: any): JanisApiRow[] {
  if (Array.isArray(body)) return body;
  if (Array.isArray(body?.rows)) return body.rows;
  if (Array.isArray(body?.data)) return body.data;
  if (Array.isArray(body?.result)) return body.result;
  return [];
}

function extractTotal(headers: Headers): number | null {
  const raw =
    headers.get("x-janis-total") ||
    headers.get("x-total-count") ||
    headers.get("x-total");
  const n = raw ? Number(raw) : NaN;
  return Number.isFinite(n) ? n : null;
}

async function fetchAllRows(authHeaders: Record<string, string>) {
  const rows: JanisApiRow[] = [];
  let page = 1;
  let expectedTotal: number | null = null;

  while (page <= MAX_PAGES) {
    const url = `${JANIS_ORDER_REPORT_URL}?sortBy=dateCreated&sortDirection=desc&page=${page}`;
    const upstream = await fetch(url, { headers: authHeaders });

    if (!upstream.ok) {
      const detail = await upstream.text().catch(() => "");
      throw new Error(
        `Janis API respondió ${upstream.status} ${upstream.statusText}${
          detail ? ` — ${detail.slice(0, 300)}` : ""
        }`
      );
    }

    if (expectedTotal === null) {
      expectedTotal = extractTotal(upstream.headers);
    }

    const body = await upstream.json();
    const pageRows = extractRows(body);
    if (!pageRows.length) break;

    rows.push(...pageRows);

    // Sin señal de paginación (endpoint devuelve todo en una sola
    // respuesta, como sugiere el ejemplo provisto): no seguimos pidiendo
    // páginas adicionales.
    if (expectedTotal === null) break;
    if (rows.length >= expectedTotal) break;

    page += 1;
  }

  return rows;
}

export default async function handler(req: any, res: any) {
  if (req.method && req.method !== "GET") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  const JANIS_CLIENT = process.env.JANIS_CLIENT;
  const JANIS_API_KEY = process.env.JANIS_API_KEY;
  const JANIS_API_SECRET = process.env.JANIS_API_SECRET;

  if (!JANIS_CLIENT || !JANIS_API_KEY || !JANIS_API_SECRET) {
    res.status(500).json({
      error:
        "Faltan variables de entorno JANIS_CLIENT / JANIS_API_KEY / JANIS_API_SECRET en el servidor.",
    });
    return;
  }

  try {
    const rawRows = await fetchAllRows({
      "janis-client": JANIS_CLIENT,
      "janis-api-key": JANIS_API_KEY,
      "janis-api-secret": JANIS_API_SECRET,
    });

    // Arranca abierto con los últimos 2 años (año actual + el anterior).
    const currentYear = new Date().getFullYear();
    const minYear = currentYear - 1;

    const rows = rawRows
      .map((r) => ({
        clientCode: String(r?.clientCode ?? "").trim(),
        month: Number(r?.month),
        year: Number(r?.year),
        totalOrders: Number(r?.totalOrders) || 0,
        totalAmount: Number(r?.totalAmount) || 0,
        cpoTotalAmount: Number(r?.cpoTotalAmount) || 0,
        cpoTotalOrders: Number(r?.cpoTotalOrders) || 0,
        npoTotalAmount: Number(r?.npoTotalAmount) || 0,
        npoTotalOrders: Number(r?.npoTotalOrders) || 0,
      }))
      .filter(
        (r) =>
          r.clientCode &&
          Number.isFinite(r.month) &&
          Number.isFinite(r.year) &&
          r.year >= minYear
      );

    // Cache corto a nivel CDN: reduce llamadas repetidas a Janis cuando
    // varias personas abren el dashboard casi al mismo tiempo.
    res.setHeader(
      "Cache-Control",
      "private, s-maxage=120, stale-while-revalidate=300"
    );
    res.status(200).json({ rows, fetchedAt: new Date().toISOString() });
  } catch (e: any) {
    res
      .status(502)
      .json({ error: e?.message || "Error consultando la API de Janis" });
  }
}
