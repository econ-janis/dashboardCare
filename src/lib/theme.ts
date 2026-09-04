/**
 * Paleta de colores y helpers visuales compartidos entre el dashboard
 * principal y cualquier otra vista (ej. /agentPerformance), para no crear
 * estilos nuevos por afuera de lo ya validado.
 */

export const UI = {
  pageBg: "bg-[#eef2f7]",
  card: "bg-white border border-slate-200 rounded-xl shadow-none",
  title: "text-sm font-semibold text-slate-700",
  subtle: "text-xs text-slate-500",
  primary: "#2563eb", // azul principal (no-SLA y SLA cumplido)
  primaryLight: "#60a5fa",
  warning: "#f59e0b", // naranjo (SLA incumplido)
  danger: "#ef4444",
  ok: "#22c55e",
  grid: "#e5e7eb",
};

export const PIE_COLORS = [
  "#2563eb",
  "#60a5fa",
  "#1d4ed8",
  "#93c5fd",
  "#3b82f6",
  "#94a3b8", // Otros
];

// Paleta categórica (hasta 8 series distinguibles, orden fijo validado
// contra confusión por daltonismo). El 9no asignado en adelante se agrupa
// en "Otros" en vez de generar un color nuevo.
export const ASSIGNEE_COLORS = [
  "#2a78d6", // blue
  "#eb6834", // orange
  "#1baf7a", // aqua
  "#eda100", // yellow
  "#e87ba4", // magenta
  "#008300", // green
  "#4a3aa7", // violet
  "#e34948", // red
];
export const ASSIGNEE_OTHERS_COLOR = "#94a3b8";
export const MAX_ASSIGNEE_SERIES = ASSIGNEE_COLORS.length;

export function hexToRgb(hex: string) {
  const h = String(hex || "").replace("#", "").trim();
  if (h.length !== 6) return null;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  if (![r, g, b].every((x) => Number.isFinite(x))) return null;
  return { r, g, b };
}

// Blanco -> Azul más oscuro (según repetición)
export function heatBg(count: number, max: number) {
  if (!count || !max) return { backgroundColor: "#ffffff", color: "#0f172a" };
  const rgb = hexToRgb(UI.primary) || { r: 37, g: 99, b: 235 };
  const ratio = Math.max(0, Math.min(1, count / max));
  const alpha = 0.06 + ratio * 0.82;
  const bg = `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${alpha})`;
  const text = alpha > 0.55 ? "#ffffff" : "#0f172a";
  return { backgroundColor: bg, color: text };
}
