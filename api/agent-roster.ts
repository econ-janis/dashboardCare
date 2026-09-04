// Proxy serverless (Vercel) para el roster de agentes configurado en
// Settings de /agentPerformance: qué valores de "Asignado" cuentan como
// agente para las comparativas de equipo (y cuáles son líderes, bots,
// etc. y deben quedar afuera).
//
// Mismo approach que api/org-mapping.ts: se persiste en Postgres (Neon)
// para que sea compartido entre todos los que usan el dashboard, no algo
// que cada uno configura en su propio navegador.
//
// Modelo: exclude-list. Todos los "Asignado" son agentes por default;
// esta tabla guarda únicamente a quienes se excluyen explícitamente
// (ej. el líder del equipo). Así, una persona nueva en el equipo aparece
// en las comparativas sin tener que tocar este Settings.
//
// Variable de entorno requerida (Vercel -> Project Settings ->
// Environment Variables, ambiente Production):
//   DATABASE_URL

import { neon } from "@neondatabase/serverless";

function normalizeAssigneeKey(value: string) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

async function ensureSchema(sql: ReturnType<typeof neon>) {
  await sql`
    CREATE TABLE IF NOT EXISTS agent_roster_exclusions (
      id bigserial PRIMARY KEY,
      assignee_key text NOT NULL UNIQUE,
      assignee_name text NOT NULL,
      created_at timestamptz NOT NULL DEFAULT now()
    )
  `;
}

type ExclusionRow = { assignee_key: string };

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
        SELECT assignee_key FROM agent_roster_exclusions ORDER BY assignee_key
      `) as ExclusionRow[];
      res.setHeader("Cache-Control", "private, no-store");
      res.status(200).json({ excludedKeys: rows.map((r) => r.assignee_key) });
      return;
    }

    if (req.method === "PUT") {
      const body = typeof req.body === "string" ? JSON.parse(req.body || "{}") : req.body || {};
      const excludedNames: string[] = Array.isArray(body?.excludedNames)
        ? Array.from(
            new Set(
              body.excludedNames
                .map((n: unknown) => String(n ?? "").trim())
                .filter((n: string) => n)
            )
          )
        : [];

      await ensureSchema(sql);

      // Reemplaza toda la lista de exclusiones por la nueva (el volumen de
      // filas es minimo, no vale la pena diffear). Todo en una transaccion
      // para que no quede a mitad de camino si algo falla.
      await sql.transaction([
        sql`DELETE FROM agent_roster_exclusions`,
        ...excludedNames
          .map((name) => ({ name, key: normalizeAssigneeKey(name) }))
          .filter(({ key }) => key)
          .map(
            ({ name, key }) => sql`
              INSERT INTO agent_roster_exclusions (assignee_key, assignee_name)
              VALUES (${key}, ${name})
              ON CONFLICT (assignee_key) DO NOTHING
            `
          ),
      ]);

      const rows = (await sql`
        SELECT assignee_key FROM agent_roster_exclusions ORDER BY assignee_key
      `) as ExclusionRow[];
      res.status(200).json({ excludedKeys: rows.map((r) => r.assignee_key) });
      return;
    }

    res.status(405).json({ error: "Method not allowed" });
  } catch (e: any) {
    res.status(502).json({ error: e?.message || "Error consultando la base de datos" });
  }
}
