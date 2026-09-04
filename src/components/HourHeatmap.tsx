import React from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { UI, heatBg } from "@/lib/theme";
import { formatInt } from "@/lib/csvParsing";

/**
 * Heatmap horario (24 franjas) reutilizado por el dashboard principal y
 * por /agentPerformance, para no duplicar el markup del grid ni el
 * criterio de color (heatBg).
 */
export function HourHeatmap({
  title = "Heatmap Horario (por hora)",
  data,
  max,
}: {
  title?: string;
  data: Array<{ hour: number; tickets: number }>;
  max: number;
}) {
  return (
    <Card className={UI.card}>
      <CardHeader>
        <CardTitle className={UI.title}>{title}</CardTitle>
      </CardHeader>
      <CardContent>
        <div className="grid grid-cols-6 gap-2">
          {data.map((x) => (
            <div
              key={x.hour}
              className="rounded-lg border border-slate-200 p-2 text-center"
              style={heatBg(x.tickets, max)}
            >
              <div className="text-xs font-semibold">{String(x.hour).padStart(2, "0")}:00</div>
              <div className="text-sm">{x.tickets ? formatInt(x.tickets) : ""}</div>
            </div>
          ))}
        </div>
      </CardContent>
    </Card>
  );
}

// Arma los datos { hour, tickets } + max a partir de un conjunto de filas
// con fecha `creada`, mismo cálculo que usaba App.tsx inline.
export function buildHourHeatmapData(rowsWithDate: Array<{ creada: Date }>) {
  const hours = Array.from({ length: 24 }, (_, i) => i);
  const counts = new Map<number, number>();
  for (const r of rowsWithDate) {
    const h = r.creada instanceof Date ? r.creada.getHours() : null;
    if (h == null) continue;
    counts.set(h, (counts.get(h) || 0) + 1);
  }
  const data = hours.map((h) => ({ hour: h, tickets: counts.get(h) || 0 }));
  const max = data.reduce((m, x) => Math.max(m, x.tickets || 0), 0);
  return { data, max };
}
