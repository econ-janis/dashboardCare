import React, { useRef } from "react";
import { Card, CardContent } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { UI } from "@/lib/theme";
import { useDashboardData } from "@/context/DashboardDataContext";

/**
 * Estado vacío: invita a subir el CSV de Jira y de Janis (order-report).
 * Un único componente reutilizado por el dashboard principal y por
 * /agentPerformance para que ambas vistas muestren exactamente el mismo
 * placeholder cuando todavía no hay datos cargados.
 */
export function UploadPrompt() {
  const { onFile, onJanisFile } = useDashboardData();
  const jiraFileInputRef = useRef<HTMLInputElement | null>(null);
  const janisFileInputRef = useRef<HTMLInputElement | null>(null);

  return (
    <Card className={`${UI.card} mt-6`}>
      <CardContent className="flex flex-col items-center gap-4 p-10 text-center">
        <div>
          <div className="text-base font-semibold text-slate-700">Todavía no hay datos cargados</div>
          <p className={`mt-1 ${UI.subtle}`}>
            Sube tu CSV de Jira y de Janis (order-report) para visualizar esta vista.
          </p>
        </div>

        <div className="flex flex-col gap-2 sm:flex-row">
          <input
            ref={jiraFileInputRef}
            type="file"
            accept=".csv,text/csv"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files && e.target.files[0];
              if (f) onFile(f);
            }}
          />
          <Button variant="outline" onClick={() => jiraFileInputRef.current?.click()}>
            Jira Data
          </Button>

          <input
            ref={janisFileInputRef}
            type="file"
            accept=".csv,text/csv"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files && e.target.files[0];
              if (f) onJanisFile(f);
            }}
          />
          <Button variant="outline" onClick={() => janisFileInputRef.current?.click()}>
            Janis Data
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}
