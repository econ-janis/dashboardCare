import React, { useCallback, useMemo, useRef, useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { UI } from "@/lib/theme";
import { formatInt } from "@/lib/csvParsing";
import { ViewNav } from "@/components/ViewNav";
import { StatCard } from "@/components/StatCard";
import { WeekHourHeatmap } from "@/components/WeekHourHeatmap";
import { exportElementToPdf } from "@/lib/exportPdf";

/**
 * Vista /teamData: actividad de los agentes de Janis en Jira (proyecto
 * JCS) — comentarios y cambios de estado por día, más un heatmap semanal
 * (día x hora) por agente. A diferencia del resto del dashboard, no se
 * alimenta del CSV cargado en DashboardDataContext: va directo contra
 * /api/jira-agent-activity (proxy serverless que llama a la API de Jira
 * con el changelog completo, algo que el export CSV no trae).
 */

type ActivityResponse = {
  from: string;
  to: string;
  days: string[];
  agents: string[];
  comments: Record<string, Record<string, number>>;
  statusChanges: Record<string, Record<string, number>>;
  heatmap: Record<string, number[][]>;
  truncated: boolean;
};

function todayIso() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function isoAddDays(iso: string, days: number) {
  const [y, m, d] = iso.split("-").map(Number);
  const dt = new Date(y, m - 1, d);
  dt.setDate(dt.getDate() + days);
  const yy = dt.getFullYear();
  const mm = String(dt.getMonth() + 1).padStart(2, "0");
  const dd = String(dt.getDate()).padStart(2, "0");
  return `${yy}-${mm}-${dd}`;
}

function formatDayHeader(iso: string) {
  const [, m, d] = iso.split("-");
  return `${d}/${m}`;
}

export default function TeamDataPage() {
  const today = useMemo(() => todayIso(), []);
  const [fromDate, setFromDate] = useState(() => isoAddDays(today, -6));
  const [toDate, setToDate] = useState(today);
  const [data, setData] = useState<ActivityResponse | null>(null);
  const [selectedAgent, setSelectedAgent] = useState<string>("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [exporting, setExporting] = useState(false);
  const reportRef = useRef<HTMLDivElement | null>(null);

  const loadActivity = useCallback(async () => {
    if (!fromDate || !toDate) return;
    setLoading(true);
    setError(null);
    try {
      const params = new URLSearchParams({ from: fromDate, to: toDate });
      const res = await fetch(`/api/jira-agent-activity?${params.toString()}`);
      const body = await res.json().catch(() => null);
      if (!res.ok) throw new Error((body && body.error) || `Error ${res.status} cargando la actividad`);
      const parsed = body as ActivityResponse;
      setData(parsed);
      setSelectedAgent((prev) => (prev && parsed.agents.includes(prev) ? prev : parsed.agents[0] || ""));
    } catch (e: any) {
      setError((e && e.message) || "No pude cargar la actividad desde Jira.");
    } finally {
      setLoading(false);
    }
  }, [fromDate, toDate]);

  const agentTotals = useMemo(() => {
    if (!data || !selectedAgent) return { comments: 0, statusChanges: 0 };
    const commentsByDay = data.comments[selectedAgent] || {};
    const statusByDay = data.statusChanges[selectedAgent] || {};
    const sum = (obj: Record<string, number>) => Object.values(obj).reduce((s, v) => s + v, 0);
    return { comments: sum(commentsByDay), statusChanges: sum(statusByDay) };
  }, [data, selectedAgent]);

  return (
    <div className={`min-h-screen ${UI.pageBg} p-4 md:p-8`}>
      <div className="mx-auto max-w-7xl" ref={reportRef}>
        <ViewNav />

        <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
          <div>
            <h1 className="text-2xl md:text-3xl font-semibold tracking-tight">Team Data</h1>
            <p className="text-sm text-slate-500 mt-1">
              Actividad de agentes en Jira (proyecto JCS): comentarios y cambios de estado por día, y heatmap semanal
              por agente.
            </p>
          </div>

          {data ? (
            <div className="export-hide flex flex-col sm:flex-row gap-2">
              <Button
                className="text-white"
                style={{ backgroundColor: UI.primary }}
                disabled={exporting}
                onClick={async () => {
                  setExporting(true);
                  setError(null);
                  try {
                    if (!reportRef.current) {
                      setError("No se encontró el contenido de la vista para exportar.");
                      return;
                    }
                    const y = today.slice(0, 4);
                    const m = today.slice(5, 7);
                    const d = today.slice(8, 10);
                    await exportElementToPdf({
                      element: reportRef.current,
                      filename: `Team_Data_${y}${m}${d}.pdf`,
                    });
                  } catch (e: any) {
                    setError((e && (e.message || String(e))) || "No se pudo exportar.");
                  } finally {
                    setExporting(false);
                  }
                }}
              >
                {exporting ? "Exportando…" : "Exportar"}
              </Button>
            </div>
          ) : null}
        </div>

        {/* Filtros: rango de fechas + botón de carga, siempre visibles (antes y después de cargar datos) */}
        <div className="export-hide mt-6 grid grid-cols-1 gap-3 md:grid-cols-3">
          <Card className={UI.card}>
            <CardContent className="p-4">
              <div className={UI.subtle}>Desde</div>
              <Input
                type="date"
                value={fromDate}
                max={toDate}
                onChange={(e) => setFromDate(e.target.value)}
              />
            </CardContent>
          </Card>
          <Card className={UI.card}>
            <CardContent className="p-4">
              <div className={UI.subtle}>Hasta</div>
              <Input
                type="date"
                value={toDate}
                min={fromDate}
                max={today}
                onChange={(e) => setToDate(e.target.value)}
              />
            </CardContent>
          </Card>
          <Card className={UI.card}>
            <CardContent className="p-4 flex items-end h-full">
              <Button
                className="w-full text-white"
                style={{ backgroundColor: UI.primary }}
                disabled={loading}
                onClick={loadActivity}
              >
                {loading ? "Cargando…" : data ? "Actualizar" : "Cargar actividad"}
              </Button>
            </CardContent>
          </Card>
        </div>

        {error ? (
          <div className="mt-4 rounded-xl border border-slate-200 bg-white p-3 text-sm text-slate-700">{error}</div>
        ) : null}

        {!data ? (
          <Card className={`${UI.card} mt-6`}>
            <CardContent className="flex flex-col items-center gap-2 p-10 text-center">
              <div className="text-base font-semibold text-slate-700">Todavía no hay datos cargados</div>
              <p className={`mt-1 max-w-md ${UI.subtle}`}>
                Elegí un rango de fechas arriba y hacé click en "Cargar actividad" para traer, desde Jira, los
                comentarios y cambios de estado de cada agente en ese período.
              </p>
            </CardContent>
          </Card>
        ) : (
          <>
            {data.truncated ? (
              <div className="mt-4 rounded-xl border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800">
                El rango elegido tiene demasiados tickets: los datos mostrados están parciales. Probá con un
                período más acotado.
              </div>
            ) : null}

            {/* Tabla comentarios - cambios de estado por día */}
            <Card className={`${UI.card} mt-6`}>
              <CardHeader>
                <CardTitle className={UI.title}>Actividad por día (comentarios - cambios de estado)</CardTitle>
                <p className={`mt-1 ${UI.subtle}`}>
                  {formatDayHeader(data.from)} al {formatDayHeader(data.to)} · formato: comentarios - cambios de
                  estado
                </p>
              </CardHeader>
              <CardContent>
                {data.days.length === 0 ? (
                  <div className={UI.subtle}>Sin días en el rango elegido.</div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="text-left text-slate-500">
                          <th className="py-2 pr-4 font-medium whitespace-nowrap">Agente</th>
                          {data.days.map((d) => (
                            <th key={d} className="py-2 px-2 font-medium text-center whitespace-nowrap">
                              {formatDayHeader(d)}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {data.agents.map((agent) => (
                          <tr key={agent} className="border-t border-slate-100">
                            <td className="py-2 pr-4 font-medium text-slate-700 whitespace-nowrap">{agent}</td>
                            {data.days.map((d) => {
                              const c = data.comments[agent]?.[d] || 0;
                              const s = data.statusChanges[agent]?.[d] || 0;
                              return (
                                <td key={d} className="py-2 px-2 text-center text-slate-600 whitespace-nowrap">
                                  {c} - {s}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </CardContent>
            </Card>

            {/* Selector de agente + heatmap semanal */}
            <div className="export-hide mt-6 flex flex-wrap gap-2">
              {data.agents.map((agent) => (
                <button
                  key={agent}
                  type="button"
                  onClick={() => setSelectedAgent(agent)}
                  className={`rounded-md px-3 py-1.5 text-sm font-medium transition-colors ${
                    selectedAgent === agent
                      ? "text-white"
                      : "bg-white text-slate-600 border border-slate-200 hover:bg-slate-50"
                  }`}
                  style={selectedAgent === agent ? { backgroundColor: UI.primary } : undefined}
                >
                  {agent}
                </button>
              ))}
            </div>

            {selectedAgent ? (
              <>
                <div className="mt-3 grid grid-cols-1 gap-3 md:grid-cols-2">
                  <StatCard title={`Comentarios — ${selectedAgent}`} value={formatInt(agentTotals.comments)} />
                  <StatCard
                    title={`Cambios de estado — ${selectedAgent}`}
                    value={formatInt(agentTotals.statusChanges)}
                  />
                </div>

                <div className="mt-3">
                  <WeekHourHeatmap
                    title={`Heatmap semanal (comentarios + cambios de estado) — ${selectedAgent}`}
                    matrix={data.heatmap[selectedAgent] || []}
                  />
                </div>
              </>
            ) : null}
          </>
        )}
      </div>
    </div>
  );
}
