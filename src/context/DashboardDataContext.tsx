import React, { createContext, useCallback, useContext, useEffect, useMemo, useState } from "react";
import Papa from "papaparse";
import {
  parseJiraCsvRows,
  parseJanisCsvRows,
  normalizeOrgKey,
  type Row,
  type JanisRow,
} from "@/lib/csvParsing";

/**
 * Estado compartido entre TODAS las vistas del dashboard (principal y
 * /agentPerformance): los CSV de Jira/Janis ya parseados, el filtro de
 * rango de mes, y el roster de agentes (quién cuenta como agente para las
 * comparativas de /agentPerformance). Así, navegar entre vistas no pierde
 * los archivos cargados ni obliga a re-subirlos, y el parseo de CSV vive
 * en un único lugar (src/lib/csvParsing.ts).
 *
 * Lo que NO vive acá (se queda local a la página del dashboard principal
 * porque solo esa vista lo necesita): el mapeo Organizacion (Jira) <->
 * Cliente (Janis) contra /api/org-mapping, y los filtros de Organización /
 * Asignado / Estado propios de esa vista.
 */

type AutoRange = { minMonth: string | null; maxMonth: string | null };

type DashboardDataContextValue = {
  rows: Row[];
  janisRows: JanisRow[];
  error: string | null;
  setError: (error: string | null) => void;
  autoRange: AutoRange;
  fromMonth: string;
  setFromMonth: (value: string) => void;
  toMonth: string;
  setToMonth: (value: string) => void;
  onFile: (file: File) => void;
  onJanisFile: (file: File) => void;
  clearAll: () => void;
  // Roster de agentes (ver src/context/DashboardDataContext.tsx y
  // api/agent-roster.ts): modelo exclude-list, todo "Asignado" es agente
  // salvo que esté en `excludedAgentNames` (ej. el líder del equipo).
  // Compartido vía Postgres para que el ajuste valga para todo el equipo.
  excludedAgentKeys: string[];
  agentRosterLoading: boolean;
  isAgent: (assignee: string) => boolean;
  setAgentIncluded: (assigneeName: string, included: boolean) => void;
};

const AGENT_ROSTER_STORAGE_KEY = "janis-care-dashboard:excluded-agents";

function loadExcludedAgentsFromStorage(): string[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = window.localStorage.getItem(AGENT_ROSTER_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter((v) => typeof v === "string" && v) : [];
  } catch {
    return [];
  }
}

const DashboardDataContext = createContext<DashboardDataContextValue | null>(null);

export function DashboardDataProvider({ children }: { children: React.ReactNode }) {
  const [rows, setRows] = useState<Row[]>([]);
  const [janisRows, setJanisRows] = useState<JanisRow[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [autoRange, setAutoRange] = useState<AutoRange>({ minMonth: null, maxMonth: null });
  const [fromMonth, setFromMonth] = useState<string>("all");
  const [toMonth, setToMonth] = useState<string>("all");

  // Roster de agentes: claves normalizadas (ver normalizeOrgKey) de
  // "Asignado" excluidas de las comparativas de /agentPerformance. Fuente
  // de verdad: Postgres (Neon) via /api/agent-roster, compartido entre
  // todos los que usan el dashboard. Se cachea en localStorage solo para
  // no arrancar sin filtro mientras responde el fetch inicial (o si la
  // API no responde). Los nombres a mostrar en el Settings de la UI
  // siempre se leen en vivo de `rows` (ver AgentPerformancePage), nunca
  // de acá — esto solo guarda las claves para el chequeo `isAgent`.
  const [excludedAgentKeys, setExcludedAgentKeys] = useState<string[]>(() =>
    loadExcludedAgentsFromStorage()
  );
  const [agentRosterLoading, setAgentRosterLoading] = useState(true);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const res = await fetch("/api/agent-roster");
        const body = await res.json().catch(() => null);
        if (!res.ok) throw new Error((body && body.error) || `Error ${res.status} cargando el roster de agentes`);
        if (!cancelled && body && Array.isArray(body.excludedKeys)) {
          setExcludedAgentKeys(body.excludedKeys);
        }
      } catch (e: any) {
        // Sin conexion a la API: seguimos con lo ultimo cacheado en localStorage.
        console.error("No pude cargar el roster de agentes desde la API:", e);
      } finally {
        if (!cancelled) setAgentRosterLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (typeof window === "undefined") return;
    try {
      window.localStorage.setItem(AGENT_ROSTER_STORAGE_KEY, JSON.stringify(excludedAgentKeys));
    } catch {
      // localStorage puede no estar disponible (modo privado, cuota llena, etc.); no es critico.
    }
  }, [excludedAgentKeys]);

  const excludedAgentKeySet = useMemo(() => new Set(excludedAgentKeys), [excludedAgentKeys]);

  const isAgent = useCallback(
    (assignee: string) => !excludedAgentKeySet.has(normalizeOrgKey(assignee)),
    [excludedAgentKeySet]
  );

  // Prende/apaga a una persona del roster de agentes. Actualiza el estado
  // local de forma optimista y lo revierte si el guardado falla, igual que
  // updateOrgMapping en CareDashboardPage. `assigneeName` es el nombre tal
  // cual aparece en el CSV (para que la API pueda, a futuro, guardar
  // también una label legible); acá solo importa su clave normalizada.
  const setAgentIncluded = useCallback(
    (assigneeName: string, included: boolean) => {
      const key = normalizeOrgKey(assigneeName);
      if (!key) return;
      const previous = excludedAgentKeys;
      const next = included ? previous.filter((k) => k !== key) : previous.includes(key) ? previous : [...previous, key];

      setExcludedAgentKeys(next);

      fetch("/api/agent-roster", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ excludedNames: next }),
      })
        .then(async (res) => {
          const body = await res.json().catch(() => null);
          if (!res.ok) throw new Error((body && body.error) || `Error ${res.status} guardando el roster de agentes`);
          if (body && Array.isArray(body.excludedKeys)) setExcludedAgentKeys(body.excludedKeys);
        })
        .catch((e: any) => {
          setError((e && e.message) || "No pude guardar el roster de agentes en la base de datos.");
          setExcludedAgentKeys(previous);
        });
    },
    [excludedAgentKeys]
  );

  const onFile = useCallback((file: File) => {
    setError(null);

    Papa.parse(file, {
      transformHeader: (h) => String(h || "").trim().toLowerCase(),
      header: true,
      skipEmptyLines: true,
      complete: (res: any) => {
        try {
          const data = (res.data || []).filter(Boolean);
          const { rows: parsed, badDate } = parseJiraCsvRows(data);

          if (!parsed.length) {
            setRows([]);
            setError(
              "No pude parsear filas con fecha 'Creada'. Revisa que el CSV tenga columna 'Creada' y formato tipo 19/ene/26 12:47 PM."
            );
            return;
          }

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
  }, []);

  const onJanisFile = useCallback(
    (file: File) => {
      setError(null);
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (res: any) => {
          try {
            const parsed = parseJanisCsvRows(res.data || []);
            setJanisRows(parsed);
            if (!rows.length && parsed.length) {
              const months = Array.from(new Set(parsed.map((r) => r.month))).sort();
              const minMonth = months[0];
              const maxMonth = months[months.length - 1];
              setAutoRange({ minMonth, maxMonth });
              setFromMonth(minMonth);
              setToMonth(maxMonth);
            }
          } catch (e: any) {
            setError((e && e.message) || "Error procesando Janis Data");
            setJanisRows([]);
          }
        },
        error: (err: any) => {
          setError(err.message);
          setJanisRows([]);
        },
      });
    },
    [rows.length]
  );

  const clearAll = useCallback(() => {
    setRows([]);
    setJanisRows([]);
    setError(null);
    setFromMonth("all");
    setToMonth("all");
    setAutoRange({ minMonth: null, maxMonth: null });
  }, []);

  const value = useMemo<DashboardDataContextValue>(
    () => ({
      rows,
      janisRows,
      error,
      setError,
      autoRange,
      fromMonth,
      setFromMonth,
      toMonth,
      setToMonth,
      onFile,
      onJanisFile,
      clearAll,
      excludedAgentKeys,
      agentRosterLoading,
      isAgent,
      setAgentIncluded,
    }),
    [
      rows,
      janisRows,
      error,
      autoRange,
      fromMonth,
      toMonth,
      onFile,
      onJanisFile,
      clearAll,
      excludedAgentKeys,
      agentRosterLoading,
      isAgent,
      setAgentIncluded,
    ]
  );

  return <DashboardDataContext.Provider value={value}>{children}</DashboardDataContext.Provider>;
}

export function useDashboardData() {
  const ctx = useContext(DashboardDataContext);
  if (!ctx) throw new Error("useDashboardData debe usarse dentro de <DashboardDataProvider>");
  return ctx;
}
