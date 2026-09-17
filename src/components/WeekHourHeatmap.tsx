import React from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { UI, heatBg } from "@/lib/theme";
import { formatInt } from "@/lib/csvParsing";

/**
 * Heatmap día-de-semana (filas, lunes a domingo) x hora (columnas, 0-23),
 * usado por la tab "Team Data" para la actividad combinada (comentarios +
 * cambios de estado) de cada agente. Mismo criterio de color (heatBg) que
 * HourHeatmap, para que ambos heatmaps del dashboard se vean consistentes.
 *
 * OJO: cada celda suma TODOS los días del rango elegido que caen en ese
 * día-de-semana (ej. "Sáb 08h" = la actividad a las 8am de TODOS los
 * sábados del período, no la de un 08/agosto puntual) — no confundir esta
 * grilla hora-del-día con la tabla de arriba, que sí es por día calendario.
 */

const DAY_LABELS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"];
const HOURS = Array.from({ length: 24 }, (_, i) => i);

export function WeekHourHeatmap({
  title = "Heatmap semanal",
  matrix,
}: {
  title?: string;
  matrix: number[][]; // 7 x 24
}) {
  const max = matrix.reduce((m, row) => Math.max(m, ...row), 0);

  return (
    <Card className={UI.card}>
      <CardHeader>
        <CardTitle className={UI.title}>{title}</CardTitle>
        <p className={`mt-1 ${UI.subtle}`}>
          Cada celda suma la actividad a esa hora en TODOS los días del período que caen en ese día de la semana
          (no es un día calendario puntual).
        </p>
      </CardHeader>
      <CardContent>
        <div className="overflow-x-auto">
          <table className="border-separate" style={{ borderSpacing: 3 }}>
            <thead>
              <tr>
                <th className="text-xs font-normal text-slate-400"> </th>
                {HOURS.map((h) => (
                  <th key={h} className="px-0 text-center text-[10px] font-normal text-slate-400">
                    {String(h).padStart(2, "0")}h
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {DAY_LABELS.map((label, dayIdx) => (
                <tr key={label}>
                  <td className="pr-2 text-xs font-medium text-slate-500 whitespace-nowrap">{label}</td>
                  {HOURS.map((h) => {
                    const value = matrix[dayIdx]?.[h] || 0;
                    return (
                      <td key={h}>
                        <div
                          className="flex h-6 w-6 items-center justify-center rounded text-[9px]"
                          style={heatBg(value, max)}
                          title={`${label} ${String(h).padStart(2, "0")}:00 — ${formatInt(value)}`}
                        >
                          {value ? formatInt(value) : ""}
                        </div>
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </CardContent>
    </Card>
  );
}
