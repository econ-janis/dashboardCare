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

// El endpoint pagina por HEADERS, no por query string: la query string se
// valida estrictamente y rechaza con 400 cualquier parámetro no esperado
// (confirmado con un 400 real al mandar "?page=1"). La app oficial
// (app.janis.in) pagina mandando X-Janis-Page / X-Janis-Page-Size /
// X-Janis-Totals como headers — replicamos ese mismo contrato acá.
const PAGE_SIZE = 60;
// Salvaguarda ante un loop inesperado: 200 páginas * 60 = 12.000 filas,
// muy por encima del historial actual (~1600 filas).
const MAX_PAGES = 200;

async function fetchPage(url: string, headers: Record<string, string>) {
  const upstream = await fetch(url, { headers });
  if (!upstream.ok) {
    const detail = await upstream.text().catch(() => "");
    throw new Error(
      `Janis API respondió ${upstream.status} ${upstream.statusText}${
        detail ? ` — ${detail.slice(0, 300)}` : ""
      }`
    );
  }
  return upstream;
}

async function fetchAllRows(authHeaders: Record<string, string>) {
  const url = `${JANIS_ORDER_REPORT_URL}?sortBy=dateCreated&sortDirection=desc`;
  const rows: JanisApiRow[] = [];
  let pagesFetched = 0;

  // Criterio simple y confirmado: la API pagina de forma prolija (60
  // filas por página, última página incompleta). Pedimos página 1, 2, 3…
  // siempre con X-Janis-Totals: false, y cortamos apenas una página
  // devuelve menos de PAGE_SIZE filas (o ninguna). Nada de sondear un
  // "total" vía headers: content-length es el tamaño en bytes de la
  // respuesta, no una cantidad de filas, y usarlo como tal daba números
  // sin sentido (llegó a calcular 342 páginas de un total real de 27).
  for (let page = 1; page <= MAX_PAGES; page++) {
    const upstream = await fetchPage(url, {
      ...authHeaders,
      "X-Janis-Page": String(page),
      "X-Janis-Page-Size": String(PAGE_SIZE),
      "X-Janis-Totals": "false",
    });
    const body = await upstream.json();
    const pageRows = extractRows(body);
    rows.push(...pageRows);
    pagesFetched += 1;
    if (pageRows.length < PAGE_SIZE) break;
  }

  return { rows, pagesFetched };
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
    const { rows: rawRows, pagesFetched } = await fetchAllRows({
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
    res.status(200).json({
      rows,
      fetchedAt: new Date().toISOString(),
      // Diagnóstico: cuántas páginas se recorrieron y cuántas filas crudas
      // (antes del filtro de años) reportó la API — útil para detectar si
      // en algún momento vuelve a cortarse antes de tiempo.
      meta: { pagesFetched, rawRowCount: rawRows.length },
    });
  } catch (e: any) {
    res
      .status(502)
      .json({ error: e?.message || "Error consultando la API de Janis" });
  }
}
