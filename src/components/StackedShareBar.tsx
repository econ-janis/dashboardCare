import React from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, LabelList, ResponsiveContainer } from "recharts";
import { formatInt } from "@/lib/csvParsing";

/**
 * Una única barra horizontal 100%-apilada, dividida en segmentos de color
 * proporcionales al share de cada categoría, con el % centrado en cada
 * segmento (si entra) y una leyenda debajo con el detalle. Pensada para
 * mostrar una distribución de pocas categorías (ej. turnos) de forma más
 * compacta que una torta, sin salir de Recharts.
 */

export type StackedShareBarEntry = {
  name: string;
  value: number;
  color: string;
};

function renderSegmentLabel(props: any, percentText: string) {
  const { x, y, width, height, value } = props || {};
  if (!value || width < 30) return null;
  return (
    <text
      x={x + width / 2}
      y={y + height / 2}
      fill="#ffffff"
      textAnchor="middle"
      dominantBaseline="central"
      fontSize={12}
      fontWeight={700}
    >
      {percentText}
    </text>
  );
}

export function StackedShareBar({
  data,
  valueFormatter = formatInt,
}: {
  data: StackedShareBarEntry[];
  valueFormatter?: (n: number) => string;
}) {
  const total = data.reduce((s, x) => s + (Number(x.value) || 0), 0);
  const chartData = [
    Object.fromEntries([["category", ""], ...data.map((d) => [d.name, d.value])]),
  ];

  const tooltipFormatter = (value: any, name: any) => {
    const v = Number(value) || 0;
    const p = total ? (v / total) * 100 : 0;
    return [`${valueFormatter(v)} (${p.toFixed(1)}%)`, name];
  };

  return (
    <div className="flex h-full flex-col justify-center gap-5">
      <div className="h-16 w-full">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={chartData} layout="vertical" barSize={48} margin={{ top: 0, right: 8, bottom: 0, left: 8 }}>
            <XAxis type="number" domain={[0, total || 1]} hide />
            <YAxis type="category" dataKey="category" hide width={0} />
            <Tooltip formatter={tooltipFormatter as any} />
            {data.map((entry, i) => {
              const percentText = total ? `${((entry.value / total) * 100).toFixed(0)}%` : "";
              const radius: [number, number, number, number] =
                i === 0 ? [8, 0, 0, 8] : i === data.length - 1 ? [0, 8, 8, 0] : [0, 0, 0, 0];
              return (
                <Bar key={entry.name} dataKey={entry.name} stackId="share" fill={entry.color} radius={radius}>
                  <LabelList dataKey={entry.name} content={(props: any) => renderSegmentLabel(props, percentText)} />
                </Bar>
              );
            })}
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="flex flex-wrap items-center justify-center gap-x-6 gap-y-2">
        {data.map((entry) => {
          const p = total ? (entry.value / total) * 100 : 0;
          return (
            <div key={entry.name} className="flex items-center gap-1.5 text-xs">
              <span className="h-2.5 w-2.5 shrink-0 rounded-full" style={{ backgroundColor: entry.color }} />
              <span className="text-slate-600">{entry.name}</span>
              <span className="font-medium text-slate-700">
                {valueFormatter(entry.value)} ({p.toFixed(1)}%)
              </span>
            </div>
          );
        })}
      </div>
    </div>
  );
}
