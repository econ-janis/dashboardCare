// Proxy serverless (Vercel) para la actividad de agentes de Janis en Jira
// (proyecto JCS): comentarios y cambios de estado por día, más una matriz
// día-de-semana x hora, para alimentar la tab "Team Data" del dashboard.
//
// Por qué existe: el export CSV/Excel del buscador de Jira no trae el
// changelog de transiciones de estado, y el campo "Comment" viene como
// texto plano sin estructura por autor/fecha. Hay que ir directo contra la
// API REST de Jira (search con expand=changelog, con fallback a
// /issue/{key}/changelog cuando el changelog viene truncado).
//
// IMPORTANTE: Atlassian eliminó /rest/api/3/search (HTTP 410 desde
// septiembre 2025, ver https://developer.atlassian.com/changelog/#CHANGE-2046).
// El endpoint vigente es /rest/api/3/search/jql, que pagina con
// `nextPageToken` en vez de `startAt`/`total` (la respuesta ya no trae
// `total`: se sabe que no hay más páginas cuando `nextPageToken` no viene
// en la respuesta).
//
// Variables de entorno requeridas (Vercel -> Project Settings ->
// Environment Variables):
//   JIRA_EMAIL     - cuenta de Atlassian con acceso al proyecto JCS
//   JIRA_API_TOKEN - generado en https://id.atlassian.com/manage-profile/security/api-tokens
// Opcionales:
//   JIRA_SITE      - default "janiscommerce.atlassian.net"
//   JIRA_PROJECT   - default "JCS"
//
// El token nunca llega al browser: este endpoint es el único lugar que lo
// usa, el frontend sólo llama a /api/jira-agent-activity.

const AGENTS: Record<string, string> = {
  "Erwin Concha": "Erwin",
  "Joel Lechuga Garcia": "Joel Lechuga Garcia (Joe)",
  "Mauricio Oruezabal": "Mauricio Oruezabal",
};
const AGENT_LABELS = Array.from(new Set(Object.values(AGENTS)));

const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const MAX_DAYS = 92;
// Resguardo para no dejar la función serverless corriendo indefinidamente
// en rangos con volumen enorme de tickets.
const MAX_ISSUES = 2000;

function addDays(dateStr: string, days: number) {
  const [y, m, d] = dateStr.split("-").map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  dt.setUTCDate(dt.getUTCDate() + days);
  return dt.toISOString().slice(0, 10);
}

function enumerateDays(from: string, toExclusive: string) {
  const days: string[] = [];
  let cur = from;
  while (cur < toExclusive) {
    days.push(cur);
    cur = addDays(cur, 1);
  }
  return days;
}

// Día de semana (0=lunes...6=domingo) y hora, leídos tal cual del string
// ISO de Jira -- SIN convertir zona horaria: igual que
// datetime.fromisoformat(...).hour en Python, que respeta el offset
// embebido en el timestamp en vez de convertirlo a otro huso.
function weekdayAndHour(createdIso: string) {
  const match = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(createdIso || "");
  if (!match) return null;
  const [, y, mo, d, h] = match;
  const utcDay = new Date(Date.UTC(Number(y), Number(mo) - 1, Number(d))).getUTCDay(); // 0=domingo
  const weekdayMon0 = (utcDay + 6) % 7; // 0=lunes
  return { weekday: weekdayMon0, hour: Number(h) };
}

type JiraUser = { displayName?: string };
type JiraComment = { author?: JiraUser; created?: string };
type ChangelogItem = { field?: string };
type ChangelogHistory = { author?: JiraUser; created?: string; items?: ChangelogItem[] };

async function jiraFetch(
  site: string,
  path: string,
  params: Record<string, string | number>,
  authHeader: string
) {
  const url = new URL(`https://${site}${path}`);
  for (const [k, v] of Object.entries(params)) url.searchParams.set(k, String(v));

  for (let attempt = 0; ; attempt += 1) {
    const resp = await fetch(url.toString(), {
      headers: { Authorization: authHeader, Accept: "application/json" },
    });
    if (resp.status === 429 && attempt < 3) {
      const retryAfter = Number(resp.headers.get("Retry-After")) || 1;
      await new Promise((r) => setTimeout(r, retryAfter * 1000));
      continue;
    }
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      throw new Error(`Jira respondió ${resp.status} en ${path}: ${body.slice(0, 300)}`);
    }
    return resp.json();
  }
}

export default async function handler(req: any, res: any) {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  const JIRA_EMAIL = process.env.JIRA_EMAIL;
  const JIRA_API_TOKEN = process.env.JIRA_API_TOKEN;
  if (!JIRA_EMAIL || !JIRA_API_TOKEN) {
    res.status(500).json({ error: "Faltan las variables de entorno JIRA_EMAIL / JIRA_API_TOKEN en el servidor." });
    return;
  }
  const site = process.env.JIRA_SITE || "janiscommerce.atlassian.net";
  const project = process.env.JIRA_PROJECT || "JCS";

  const from = String(req.query?.from || "");
  const to = String(req.query?.to || "");
  if (!DATE_RE.test(from) || !DATE_RE.test(to)) {
    res.status(400).json({ error: "Parámetros 'from' y 'to' requeridos, formato YYYY-MM-DD." });
    return;
  }
  if (from > to) {
    res.status(400).json({ error: "'from' no puede ser posterior a 'to'." });
    return;
  }

  const toExclusive = addDays(to, 1);
  const days = enumerateDays(from, toExclusive);
  if (days.length > MAX_DAYS) {
    res.status(400).json({ error: `Rango demasiado grande (máx. ${MAX_DAYS} días). Elegí un período más acotado.` });
    return;
  }

  const authHeader = `Basic ${Buffer.from(`${JIRA_EMAIL}:${JIRA_API_TOKEN}`).toString("base64")}`;

  const comments: Record<string, Record<string, number>> = {};
  const statusChanges: Record<string, Record<string, number>> = {};
  const heatmap: Record<string, number[][]> = {};
  for (const label of AGENT_LABELS) {
    comments[label] = {};
    statusChanges[label] = {};
    heatmap[label] = Array.from({ length: 7 }, () => Array(24).fill(0));
  }

  const bump = (bucket: Record<string, Record<string, number>>, label: string, day: string) => {
    bucket[label][day] = (bucket[label][day] || 0) + 1;
  };

  const bucketHeatmap = (label: string, createdIso: string) => {
    const wh = weekdayAndHour(createdIso);
    if (wh) heatmap[label][wh.weekday][wh.hour] += 1;
  };

  // No usar comillas simples de más: el valor de jql va siempre entre
  // comillas dobles (fechas), tal cual el JQL de referencia.
  const jql = `project = ${project} AND updated >= "${from}" AND updated <= "${toExclusive}"`;

  try {
    let pageToken: string | undefined;
    let issuesSeen = 0;
    let truncated = false;

    while (true) {
      const params: Record<string, string | number> = {
        jql,
        fields: "comment",
        expand: "changelog",
        maxResults: 100,
      };
      if (pageToken) params.nextPageToken = pageToken;

      const data: any = await jiraFetch(site, "/rest/api/3/search/jql", params, authHeader);
      const issues: any[] = data.issues || [];

      for (const issue of issues) {
        for (const c of (issue.fields?.comment?.comments || []) as JiraComment[]) {
          const author = c.author?.displayName || "";
          const created = c.created || "";
          const day = created.slice(0, 10);
          const label = AGENTS[author];
          if (label && day >= from && day < toExclusive) {
            bump(comments, label, day);
            bucketHeatmap(label, created);
          }
        }

        let histories: ChangelogHistory[] = issue.changelog?.histories || [];
        const changelogTotal = Number(issue.changelog?.total ?? histories.length);
        if (changelogTotal > histories.length) {
          let clStart = histories.length;
          while (clStart < changelogTotal) {
            const clData: any = await jiraFetch(
              site,
              `/rest/api/3/issue/${issue.key}/changelog`,
              { startAt: clStart, maxResults: 100 },
              authHeader
            );
            const values: ChangelogHistory[] = clData.values || [];
            if (!values.length) break;
            histories = histories.concat(values);
            clStart += values.length;
          }
        }

        for (const h of histories) {
          const hasStatusItem = (h.items || []).some((i) => i.field === "status");
          if (!hasStatusItem) continue;
          const author = h.author?.displayName || "";
          const created = h.created || "";
          const day = created.slice(0, 10);
          const label = AGENTS[author];
          if (label && day >= from && day < toExclusive) {
            bump(statusChanges, label, day);
            bucketHeatmap(label, created);
          }
        }
      }

      issuesSeen += issues.length;
      pageToken = data.nextPageToken;
      // Sin `nextPageToken` en la respuesta, no hay más páginas (la API ya
      // no devuelve `total`, ver nota arriba). `issues.length === 0` es un
      // resguardo extra por si la API devuelve una página vacía sin cortar
      // el token antes.
      if (!pageToken || issues.length === 0) break;
      if (issuesSeen >= MAX_ISSUES) {
        truncated = true;
        break;
      }
    }

    res.setHeader("Cache-Control", "private, no-store");
    res.status(200).json({
      from,
      to,
      days,
      agents: AGENT_LABELS,
      comments,
      statusChanges,
      heatmap,
      truncated,
    });
  } catch (e: any) {
    res.status(502).json({ error: e?.message || "Error consultando la API de Jira" });
  }
}
