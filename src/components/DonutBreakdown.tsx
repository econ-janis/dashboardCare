import React from "react";
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer } from "recharts";
import { formatInt } from "@/lib/csvParsing";

/**
 * Torta (donut) + leyenda lateral, mismo patrón visual que "Top 10
 * Organizaciones" del dashboard principal (CareDashboardPage): porcentaje
 * dibujado sobre cada porción, tooltip con valor + %, y una lista lateral
 * con el detalle. Extraído a componente compartido para no reimplementar
 * el mismo look en cada vista nueva (ej. distribución por turno de
 * /agentPerformance).
 */

export type DonutBreakdownEntry = {
  name: string;
  value: number;
  color: string;
};

// Dibuja el porcentaje directamente sobre cada porción de la torta. Las
// porciones muy angostas (<4%) se omiten para no amontonar texto; esos
// valores igual quedan disponibles en el tooltip y en la lista lateral.
function renderPieSliceLabel(props: any) {
  const { cx, cy, midAngle, innerRadius, outerRadius, percent } = props || {};
  if (!percent || percent < 0.04) return null;
  const RADIAN = Math.PI / 180;
  const radius = innerRadius + (outerRadius - innerRadius) * 0.6;
  const x = cx + radius * Math.cos(-midAngle * RADIAN);
  const y = cy + radius * Math.sin(-midAngle * RADIAN);
  return (
    <text x={x} y={y} fill="#ffffff" textAnchor="middle" dominantBaseline="central" fontSize={11} fontWeight={600}>
      {`${(percent * 100).toFixed(0)}%`}
    </text>
  );
}

export function DonutBreakdown({
  data,
  valueFormatter = formatInt,
}: {
  data: DonutBreakdownEntry[];
  valueFormatter?: (n: number) => string;
}) {
  const total = data.reduce((s, x) => s + (Number(x.value) || 0), 0);

  const tooltipFormatter = (value: any, _name: any, props: any) => {
    const v = Number(value) || 0;
    const p = total ? (v / total) * 100 : 0;
    const label = props && props.payload && props.payload.name ? props.payload.name : "";
    return [`${valueFormatter(v)} (${p.toFixed(2)}%)`, label];
  };

  return (
    <div className="flex h-full flex-col gap-3 md:flex-row md:items-center">
      <div className="h-72 w-full shrink-0 md:h-full md:flex-1">
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Tooltip formatter={tooltipFormatter as any} />
            <Pie
              data={data}
              dataKey="value"
              nameKey="name"
              outerRadius={95}
              innerRadius={50}
              paddingAngle={2}
              label={renderPieSliceLabel}
              labelLine={false}
            >
              {data.map((entry) => (
                <Cell key={entry.name} fill={entry.color} />
              ))}
            </Pie>
          </PieChart>
        </ResponsiveContainer>
      </div>
      <div className="flex w-full flex-col gap-1.5 overflow-y-auto md:h-full md:w-48 md:shrink-0">
        {data.map((entry) => {
          const p = total ? (entry.value / total) * 100 : 0;
          return (
            <div key={entry.name} className="flex items-center justify-between gap-2 text-xs" title={entry.name}>
              <span className="flex min-w-0 items-center gap-1.5">
                <span className="h-2.5 w-2.5 shrink-0 rounded-full" style={{ backgroundColor: entry.color }} />
                <span className="truncate text-slate-600">{entry.name}</span>
              </span>
              <span className="shrink-0 font-medium text-slate-700">
                {valueFormatter(entry.value)} ({p.toFixed(1)}%)
              </span>
            </div>
          );
        })}
      </div>
    </div>
  );
}
