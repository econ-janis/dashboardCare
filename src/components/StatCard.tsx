import React from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { UI } from "@/lib/theme";

/**
 * Tarjeta de KPI reutilizada por el dashboard principal y por cualquier
 * otra vista (ej. /agentPerformance). Extraída del helper `kpiCard(...)`
 * que ya existía en App.tsx, sin cambiar su estructura visual.
 */

export function KpiPreviousPeriod({ children }: { children: React.ReactNode }) {
  return <div className="mt-3 border-t border-slate-100 pt-2 text-xs text-slate-500">{children}</div>;
}

export function StatCard({
  title,
  value,
  subtitle,
  right,
  badge,
  previousPeriod,
}: {
  title: string;
  value: React.ReactNode;
  subtitle?: React.ReactNode;
  right?: string;
  badge?: React.ReactNode;
  previousPeriod?: React.ReactNode;
}) {
  return (
    <Card className={UI.card}>
      <CardHeader className="pb-2">
        <CardTitle className={UI.title}>{title}</CardTitle>
      </CardHeader>
      <CardContent>
        <div className="flex items-start justify-between gap-3">
          <div>
            <div className="text-2xl font-semibold tracking-tight text-slate-900">{value}</div>
            {badge ? <div className="mt-2">{badge}</div> : null}
            {subtitle ? <div className={"mt-1 " + UI.subtle}>{subtitle}</div> : null}
          </div>
          {right ? <div className={"text-right " + UI.subtle}>{right}</div> : null}
        </div>
        {previousPeriod ? <KpiPreviousPeriod>{previousPeriod}</KpiPreviousPeriod> : null}
      </CardContent>
    </Card>
  );
}
