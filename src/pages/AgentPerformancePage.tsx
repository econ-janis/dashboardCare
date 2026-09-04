import React, { useEffect, useMemo, useState } from "react";
import {
  LineChart,
  Line,
  BarChart,
  Bar,
  Cell,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
  Legend,
  ResponsiveContainer,
  RadarChart,
  PolarGrid,
  PolarAngleAxis,
  PolarRadiusAxis,
  Radar,
} from "recharts";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { UI } from "@/lib/theme";
import { formatInt, formatPct, monthLabel, pct, type Row } from "@/lib/csvParsing";
import { useDashboardData } from "@/context/DashboardDataContext";
import { ViewNav } from "@/components/ViewNav";
import { UploadPrompt } from "@/components/UploadPrompt";
import { StatCard } from "@/components/StatCard";
import { HourHeatmap, buildHourHeatmapData } from "@/components/HourHeatmap";

/**
 * Vista /agentPerformance: performance individual de un Asignado, sobre el
 * mismo CSV de Jira ya cargado en el dashboard principal (vía
 * DashboardDataContext, sin volver a parsear ni re-subir archivos).
 *
 * Filtros propios: rango de mes (compartido con el dashboard principal) +
 * Agente. No reutiliza los filtros de Organización/Estado del dashboard
 * principal: "filtrar todo el dataset por ese Asignado" es sólo fecha +
 * agente.
 */

const SIN_ASIGNAR = "(Sin asignar)";
const RADAR_AXES = ["Volumen", "SLA", "CSAT", "Velocidad", "Cobertura"] as const;

function monthRangeDayCount(
  fromMonth: string,
  toMonth: string,
  autoRange: { minMonth: string | null; maxMonth: string | null }
) {
  const from = fromMonth === "all" ? autoRange.minMonth : fromMonth;
  const to = toMonth === "all" ? autoRange.maxMonth : toMonth;
  if (!from || !to) return 0;
  const [fy, fm] = from.split("-").map(Number);
  const [ty, tm] = to.split("-").map(Number);
  if (!Number.isFinite(fy) || !Number.isFinite(fm) || !Number.isFinite(ty) || !Number.isFinite(tm)) return 0;
  const start = new Date(fy, fm - 1, 1);
  const end = new Date(ty, tm, 0); // último día del mes "hasta"
  if (end < start) return 0;
  return Math.round((end.getTime() - start.getTime()) / 86400000) + 1;
}

// Agrega, para un mismo conjunto de filas ya filtradas por período, las
// métricas base por Asignado que alimentan el ranking de carga y el radar.
function buildAgentAggregates(periodRows: Row[], days: number) {
  const byAgent = new Map<
    string,
    { name: string; tickets: number; cumplido: number; satSum: number; satCount: number; orgs: Set<string> }
  >();

  for (const r of periodRows) {
    const name = (r.asignado || "").trim() || SIN_ASIGNAR;
    const cur =
      byAgent.get(name) ||
      { name, tickets: 0, cumplido: 0, satSum: 0, satCount: 0, orgs: new Set<string>() };
    cur.tickets += 1;
    if (r.slaResponseStatus === "Cumplido") cur.cumplido += 1;
    if (r.satisfaction != null) {
      cur.satSum += r.satisfaction;
      cur.satCount += 1;
    }
    if (r.organization) cur.orgs.add(r.organization);
    byAgent.set(name, cur);
  }

  return Array.from(byAgent.values())
    .map((a) => ({
      name: a.name,
      tickets: a.tickets,
      slaPct: a.tickets ? pct(a.cumplido, a.tickets) : 0,
      csatAvg: a.satCount ? a.satSum / a.satCount : null,
      ticketsPerDay: days > 0 ? a.tickets / days : 0,
      orgsCount: a.orgs.size,
    }))
    .sort((a, b) => b.tickets - a.tickets);
}

export default function AgentPerformancePage() {
  const { rows, autoRange, fromMonth, setFromMonth, toMonth, setToMonth } = useDashboardData();
  const [selectedAgent, setSelectedAgent] = useState<string>("");

  const minMonthBound = autoRange.minMonth ?? undefined;
  const maxMonthBound = autoRange.maxMonth ?? undefined;

  const agentOptions = useMemo(
    () => Array.from(new Set(rows.map((r) => (r.asignado || "").trim() || SIN_ASIGNAR))).sort(),
    [rows]
  );

  const periodRows = useMemo(
    () =>
      rows.filter((r) => {
        if (fromMonth !== "all" && r.month < fromMonth) return false;
        if (toMonth !== "all" && r.month > toMonth) return false;
        return true;
      }),
    [rows, fromMonth, toMonth]
  );

  const days = useMemo(() => monthRangeDayCount(fromMonth, toMonth, autoRange), [fromMonth, toMonth, autoRange]);

  const agentAggregates = useMemo(() => buildAgentAggregates(periodRows, days), [periodRows, days]);

  // Si no hay agente elegido (o el elegido ya no aparece en el período/
  // dataset actual), se preselecciona el de mayor volumen para que la
  // vista no arranque vacía.
  useEffect(() => {
    if (!agentOptions.length) {
      if (selectedAgent) setSelectedAgent("");
      return;
    }
    if (selectedAgent && agentOptions.includes(selectedAgent)) return;
    const top = agentAggregates[0]?.name;
    setSelectedAgent(top && agentOptions.includes(top) ? top : agentOptions[0]);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [agentOptions]);

  const agentRows = useMemo(
    () => periodRows.filter((r) => ((r.asignado || "").trim() || SIN_ASIGNAR) === selectedAgent),
    [periodRows, selectedAgent]
  );

  const agentStats = useMemo(() => agentAggregates.find((a) => a.name === selectedAgent) || null, [
    agentAggregates,
    selectedAgent,
  ]);

  const rankingPosition = useMemo(
    () => agentAggregates.findIndex((a) => a.name === selectedAgent) + 1,
    [agentAggregates, selectedAgent]
  );

  // Tickets resueltos por mes del agente vs. promedio del equipo (total del
  // mes / cantidad de agentes con actividad ese mes).
  const monthlyVsTeam = useMemo(() => {
    const months = Array.from(new Set(periodRows.map((r) => r.month))).sort();
    return months.map((m) => {
      const monthRows = periodRows.filter((r) => r.month === m);
      const agentTickets = monthRows.filter(
        (r) => ((r.asignado || "").trim() || SIN_ASIGNAR) === selectedAgent
      ).length;
      const activeAgents = new Set(monthRows.map((r) => (r.asignado || "").trim() || SIN_ASIGNAR));
      const teamAvg = activeAgents.size ? monthRows.length / activeAgents.size : 0;
      return { month: m, agente: agentTickets, equipo: Number(teamAvg.toFixed(2)) };
    });
  }, [periodRows, selectedAgent]);

  // % SLA Response por mes del agente.
  const slaByMonth = useMemo(() => {
    const months = Array.from(new Set(agentRows.map((r) => r.month))).sort();
    return months.map((m) => {
      const monthRows = agentRows.filter((r) => r.month === m);
      const cumplido = monthRows.filter((r) => r.slaResponseStatus === "Cumplido").length;
      return { month: m, slaPct: Number(pct(cumplido, monthRows.length).toFixed(1)) };
    });
  }, [agentRows]);

  const topOrganizations = useMemo(() => {
    const counts = new Map<string, number>();
    for (const r of agentRows) {
      const org = (r.organization || "").trim() || "(Sin organización)";
      counts.set(org, (counts.get(org) || 0) + 1);
    }
    const total = agentRows.length;
    return Array.from(counts.entries())
      .map(([name, tickets]) => ({ name, tickets, sharePct: total ? (tickets / total) * 100 : 0 }))
      .sort((a, b) => b.tickets - a.tickets)
      .slice(0, 8);
  }, [agentRows]);

  const hourHeatmap = useMemo(() => buildHourHeatmapData(agentRows), [agentRows]);

  const radarData = useMemo(() => {
    if (!agentStats || !agentAggregates.length) return [];
    const bestOf = (fn: (a: (typeof agentAggregates)[number]) => number) =>
      Math.max(0, ...agentAggregates.map(fn));
    const bestVolume = bestOf((a) => a.tickets);
    const bestSla = bestOf((a) => a.slaPct);
    const bestCsat = bestOf((a) => a.csatAvg ?? 0);
    const bestSpeed = bestOf((a) => a.ticketsPerDay);
    const bestCoverage = bestOf((a) => a.orgsCount);
    const ratio = (value: number, best: number) => (best > 0 ? Math.min(100, (value / best) * 100) : 0);

    const values = [
      ratio(agentStats.tickets, bestVolume),
      ratio(agentStats.slaPct, bestSla),
      ratio(agentStats.csatAvg ?? 0, bestCsat),
      ratio(agentStats.ticketsPerDay, bestSpeed),
      ratio(agentStats.orgsCount, bestCoverage),
    ];
    return RADAR_AXES.map((axis, i) => ({ axis, value: Number(values[i].toFixed(1)) }));
  }, [agentStats, agentAggregates]);

  return (
    <div className={`min-h-screen ${UI.pageBg} p-4 md:p-8`}>
      <div className="mx-auto max-w-7xl">
        <ViewNav />

        <div>
          <h1 className="text-2xl md:text-3xl font-semibold tracking-tight">Performance por Agente</h1>
          <p className="text-sm text-slate-500 mt-1">
            Performance individual de cada persona asignada, sobre el mismo CSV de Jira cargado en el dashboard.
          </p>
        </div>

        {rows.length === 0 ? (
          <UploadPrompt />
        ) : (
          <>
            {/* Filtros */}
            <div className="mt-6 grid grid-cols-1 gap-3 md:grid-cols-3">
              <Card className={UI.card}>
                <CardContent className="p-4">
                  <div className={UI.subtle}>Desde (mes)</div>
                  <Input
                    type="month"
                    value={fromMonth === "all" ? autoRange.minMonth || "" : fromMonth}
                    min={minMonthBound}
                    max={maxMonthBound}
                    onChange={(e) => setFromMonth(e.target.value || "all")}
                  />
                </CardContent>
              </Card>
              <Card className={UI.card}>
                <CardContent className="p-4">
                  <div className={UI.subtle}>Hasta (mes)</div>
                  <Input
                    type="month"
                    value={toMonth === "all" ? autoRange.maxMonth || "" : toMonth}
                    min={minMonthBound}
                    max={maxMonthBound}
                    onChange={(e) => setToMonth(e.target.value || "all")}
                  />
                </CardContent>
              </Card>
              <Card className={UI.card}>
                <CardContent className="p-4">
                  <div className={UI.subtle}>Agente</div>
                  <Select value={selectedAgent} onValueChange={setSelectedAgent}>
                    <SelectTrigger>
                      <SelectValue placeholder="Elegí un agente" />
                    </SelectTrigger>
                    <SelectContent>
                      {agentOptions.map((name) => (
                        <SelectItem key={name} value={name}>
                          {name}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </CardContent>
              </Card>
            </div>

            {!agentStats ? (
              <Card className={`${UI.card} mt-6`}>
                <CardContent className="p-6 text-sm text-slate-500">
                  Sin datos para el agente y período seleccionados.
                </CardContent>
              </Card>
            ) : (
              <>
                {/* Stat cards */}
                <div className="mt-6 grid grid-cols-1 gap-3 md:grid-cols-3 xl:grid-cols-5">
                  <StatCard title="Tickets (vista)" value={formatInt(agentStats.tickets)} />
                  <StatCard title="Cumplimiento SLA Response" value={formatPct(agentStats.slaPct)} />
                  <StatCard
                    title="CSAT promedio"
                    value={agentStats.csatAvg != null ? agentStats.csatAvg.toFixed(2) : "Sin dato"}
                  />
                  <StatCard title="Tickets/día promedio" value={agentStats.ticketsPerDay.toFixed(2)} />
                  <StatCard
                    title="Ranking de carga"
                    value={`#${rankingPosition} de ${agentAggregates.length}`}
                    subtitle="Posición por volumen de tickets vs. el resto del equipo"
                  />
                </div>

                {/* Tickets/mes agente vs equipo */}
                <Card className={`${UI.card} mt-6`}>
                  <CardHeader>
                    <CardTitle className={UI.title}>Tickets resueltos por mes — Agente vs. Equipo</CardTitle>
                  </CardHeader>
                  <CardContent className="h-80">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={monthlyVsTeam}>
                        <CartesianGrid stroke={UI.grid} />
                        <XAxis dataKey="month" tickFormatter={monthLabel as any} interval="preserveStartEnd" />
                        <YAxis allowDecimals={false} width={50} />
                        <Tooltip labelFormatter={(l) => monthLabel(String(l))} />
                        <Legend />
                        <Line
                          type="monotone"
                          dataKey="agente"
                          name={selectedAgent}
                          stroke={UI.primary}
                          strokeWidth={2}
                          dot={{ r: 3 }}
                        />
                        <Line
                          type="monotone"
                          dataKey="equipo"
                          name="Promedio equipo"
                          stroke={UI.warning}
                          strokeWidth={2}
                          strokeDasharray="5 5"
                          dot={false}
                        />
                      </LineChart>
                    </ResponsiveContainer>
                  </CardContent>
                </Card>

                {/* SLA % por mes */}
                <Card className={`${UI.card} mt-6`}>
                  <CardHeader>
                    <CardTitle className={UI.title}>% SLA Response por mes</CardTitle>
                  </CardHeader>
                  <CardContent className="h-80">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={slaByMonth}>
                        <CartesianGrid stroke={UI.grid} />
                        <XAxis dataKey="month" tickFormatter={monthLabel as any} interval="preserveStartEnd" />
                        <YAxis domain={[0, 100]} tickFormatter={(v: any) => `${v}%`} width={50} />
                        <Tooltip
                          labelFormatter={(l) => monthLabel(String(l))}
                          formatter={(v: any) => [`${v}%`, "SLA Response"]}
                        />
                        <Line type="monotone" dataKey="slaPct" name="SLA Response" stroke={UI.ok} strokeWidth={2} dot={{ r: 3 }} />
                      </LineChart>
                    </ResponsiveContainer>
                  </CardContent>
                </Card>

                <div className="mt-6 grid grid-cols-1 gap-3 lg:grid-cols-2">
                  {/* Ranking de carga */}
                  <Card className={UI.card}>
                    <CardHeader>
                      <CardTitle className={UI.title}>Ranking de carga (volumen de tickets)</CardTitle>
                    </CardHeader>
                    <CardContent style={{ height: Math.max(240, agentAggregates.length * 32) }}>
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={agentAggregates} layout="vertical" margin={{ left: 60 }}>
                          <CartesianGrid stroke={UI.grid} />
                          <XAxis type="number" allowDecimals={false} />
                          <YAxis type="category" dataKey="name" width={120} />
                          <Tooltip formatter={(v: any) => [formatInt(v), "Tickets"]} />
                          <Bar dataKey="tickets">
                            {agentAggregates.map((a) => (
                              <Cell key={a.name} fill={a.name === selectedAgent ? UI.primary : "#cbd5e1"} />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </CardContent>
                  </Card>

                  {/* Radar */}
                  <Card className={UI.card}>
                    <CardHeader>
                      <CardTitle className={UI.title}>Radar de performance (vs. mejor del equipo)</CardTitle>
                    </CardHeader>
                    <CardContent className="h-80">
                      <ResponsiveContainer width="100%" height="100%">
                        <RadarChart data={radarData}>
                          <PolarGrid stroke={UI.grid} />
                          <PolarAngleAxis dataKey="axis" tick={{ fontSize: 12 }} />
                          <PolarRadiusAxis domain={[0, 100]} tick={{ fontSize: 10 }} />
                          <Radar
                            name={selectedAgent}
                            dataKey="value"
                            stroke={UI.primary}
                            fill={UI.primary}
                            fillOpacity={0.35}
                          />
                          <Tooltip formatter={(v: any) => [`${v}`, "vs. mejor del equipo"]} />
                        </RadarChart>
                      </ResponsiveContainer>
                    </CardContent>
                  </Card>
                </div>

                <div className="mt-6 grid grid-cols-1 gap-3 lg:grid-cols-2">
                  <HourHeatmap
                    title={`Heatmap Horario — ${selectedAgent}`}
                    data={hourHeatmap.data}
                    max={hourHeatmap.max}
                  />

                  {/* Top organizaciones del agente */}
                  <Card className={UI.card}>
                    <CardHeader>
                      <CardTitle className={UI.title}>Top organizaciones atendidas</CardTitle>
                    </CardHeader>
                    <CardContent>
                      {topOrganizations.length === 0 ? (
                        <div className={UI.subtle}>Sin datos en el período</div>
                      ) : (
                        <div className="space-y-2.5">
                          {topOrganizations.map((org) => (
                            <div key={org.name}>
                              <div className="flex items-center justify-between text-xs text-slate-600">
                                <span className="truncate pr-2" title={org.name}>
                                  {org.name}
                                </span>
                                <span className="shrink-0 font-medium text-slate-700">
                                  {formatInt(org.tickets)} ({org.sharePct.toFixed(1)}%)
                                </span>
                              </div>
                              <div className="mt-1 h-2 rounded-full bg-slate-100 overflow-hidden">
                                <div
                                  className="h-2 rounded-full"
                                  style={{ width: `${org.sharePct}%`, backgroundColor: UI.primary }}
                                />
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </CardContent>
                  </Card>
                </div>
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}
