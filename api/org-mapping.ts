// Proxy serverless (Vercel) para el mapeo Organizacion (Jira) <-> Cliente
// (Janis Data) configurado en Settings.
//
// Por que existe: antes el mapeo vivia solo en localStorage del navegador
// (por-persona, por-dispositivo). Esto lo guarda en Postgres (Neon) para
// que sea compartido entre todos los que usan el dashboard.
//
// Variable de entorno requerida (Vercel -> Project Settings ->
// Environment Variables, ambiente Production):
//   DATABASE_URL
//
// Usa el driver HTTP de Neon (@neondatabase/serverless) en vez de una
// conexion TCP tradicional: es el approach recomendado para funciones
// serverless (sin pooling de conexiones que administrar) y funciona bien
// detras de entornos que solo permiten salida HTTPS.

import { neon } from "@neondatabase/serverless";

function normalizeOrgKey(value: string) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

async function ensureSchema(sql: ReturnType<typeof neon>) {
  await sql`
    CREATE TABLE IF NOT EXISTS org_client_mapping (
      id bigserial PRIMARY KEY,
      jira_org text NOT NULL,
      jira_org_key text NOT NULL,
      janis_client text NOT NULL,
      janis_client_key text NOT NULL,
      created_at timestamptz NOT NULL DEFAULT now(),
      UNIQUE (jira_org_key, janis_client_key)
    )
  `;
  await sql`
    CREATE INDEX IF NOT EXISTS org_client_mapping_jira_org_key_idx
    ON org_client_mapping (jira_org_key)
  `;
}

type MappingRow = { jira_org_key: string; janis_client: string };

function rowsToMapping(rows: MappingRow[]): Record<string, string[]> {
  const mapping: Record<string, string[]> = {};
  for (const r of rows) {
    if (!mapping[r.jira_org_key]) mapping[r.jira_org_key] = [];
    mapping[r.jira_org_key].push(r.janis_client);
  }
  return mapping;
}

export default async function handler(req: any, res: any) {
  const DATABASE_URL = process.env.DATABASE_URL;
  if (!DATABASE_URL) {
    res.status(500).json({ error: "Falta la variable de entorno DATABASE_URL en el servidor." });
    return;
  }

  const sql = neon(DATABASE_URL);

  try {
    if (req.method === "GET") {
      await ensureSchema(sql);
      const rows = (await sql`
        SELECT jira_org_key, janis_client FROM org_client_mapping ORDER BY jira_org_key, janis_client
      `) as MappingRow[];
      res.setHeader("Cache-Control", "private, no-store");
      res.status(200).json({ mapping: rowsToMapping(rows) });
      return;
    }

    if (req.method === "PUT") {
      const body = typeof req.body === "string" ? JSON.parse(req.body || "{}") : req.body || {};
      const jiraOrg = String(body?.jiraOrg ?? "").trim();
      const janisClients: string[] = Array.isArray(body?.janisClients)
        ? Array.from(
            new Set(
              body.janisClients
                .map((c: unknown) => String(c ?? "").trim())
                .filter((c: string) => c)
            )
          )
        : [];

      if (!jiraOrg) {
        res.status(400).json({ error: "Falta 'jiraOrg' en el body." });
        return;
      }

      const jiraOrgKey = normalizeOrgKey(jiraOrg);
      await ensureSchema(sql);

      // Reemplaza todas las asociaciones de esta Organizacion por la lista
      // nueva (borrar + insertar es mas simple que hacer un diff, y el
      // volumen de filas es minimo). Todo en una sola transaccion HTTP para
      // que no quede a mitad de camino si algo falla.
      await sql.transaction([
        sql`DELETE FROM org_client_mapping WHERE jira_org_key = ${jiraOrgKey}`,
        ...janisClients
          .map((janisClient) => ({ janisClient, janisClientKey: normalizeOrgKey(janisClient) }))
          .filter(({ janisClientKey }) => janisClientKey)
          .map(
            ({ janisClient, janisClientKey }) => sql`
              INSERT INTO org_client_mapping (jira_org, jira_org_key, janis_client, janis_client_key)
              VALUES (${jiraOrg}, ${jiraOrgKey}, ${janisClient}, ${janisClientKey})
              ON CONFLICT (jira_org_key, janis_client_key) DO NOTHING
            `
          ),
      ]);

      const rows = (await sql`
        SELECT jira_org_key, janis_client FROM org_client_mapping ORDER BY jira_org_key, janis_client
      `) as MappingRow[];
      res.status(200).json({ mapping: rowsToMapping(rows) });
      return;
    }

    res.status(405).json({ error: "Method not allowed" });
  } catch (e: any) {
    res.status(502).json({ error: e?.message || "Error consultando la base de datos" });
  }
}
