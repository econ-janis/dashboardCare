import React, { createContext, useCallback, useContext, useMemo, useState } from "react";
import Papa from "papaparse";
import {
  parseJiraCsvRows,
  parseJanisCsvRows,
  type Row,
  type JanisRow,
} from "@/lib/csvParsing";

/**
 * Estado compartido entre TODAS las vistas del dashboard (principal y
 * /agentPerformance): los CSV de Jira/Janis ya parseados y el filtro de
 * rango de mes. Así, navegar entre vistas no pierde los archivos cargados
 * ni obliga a re-subirlos, y el parseo de CSV vive en un único lugar
 * (src/lib/csvParsing.ts).
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
};

const DashboardDataContext = createContext<DashboardDataContextValue | null>(null);

export function DashboardDataProvider({ children }: { children: React.ReactNode }) {
  const [rows, setRows] = useState<Row[]>([]);
  const [janisRows, setJanisRows] = useState<JanisRow[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [autoRange, setAutoRange] = useState<AutoRange>({ minMonth: null, maxMonth: null });
  const [fromMonth, setFromMonth] = useState<string>("all");
  const [toMonth, setToMonth] = useState<string>("all");

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
    }),
    [rows, janisRows, error, autoRange, fromMonth, toMonth, onFile, onJanisFile, clearAll]
  );

  return <DashboardDataContext.Provider value={value}>{children}</DashboardDataContext.Provider>;
}

export function useDashboardData() {
  const ctx = useContext(DashboardDataContext);
  if (!ctx) throw new Error("useDashboardData debe usarse dentro de <DashboardDataProvider>");
  return ctx;
}
