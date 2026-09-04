import React from "react";
import { NavLink } from "react-router-dom";
import { UI } from "@/lib/theme";

/**
 * Navegación entre las vistas del dashboard (principal y
 * /agentPerformance), mostrada arriba en ambas páginas.
 */
export function ViewNav() {
  const tabClass = ({ isActive }: { isActive: boolean }) =>
    `rounded-md px-3 py-1.5 text-sm font-medium transition-colors ${
      isActive ? "text-white" : "bg-white text-slate-600 border border-slate-200 hover:bg-slate-50"
    }`;
  const tabStyle = ({ isActive }: { isActive: boolean }) =>
    isActive ? { backgroundColor: UI.primary } : undefined;

  return (
    <nav className="export-hide mb-4 flex flex-wrap gap-2">
      <NavLink to="/" end className={tabClass} style={tabStyle}>
        Dashboard
      </NavLink>
      <NavLink to="/agentPerformance" className={tabClass} style={tabStyle}>
        Performance por agente
      </NavLink>
    </nav>
  );
}
