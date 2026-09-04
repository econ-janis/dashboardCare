import { BrowserRouter, Navigate, Route, Routes } from "react-router-dom";
import { DashboardDataProvider } from "@/context/DashboardDataContext";
import CareDashboardPage from "@/pages/CareDashboardPage";
import AgentPerformancePage from "@/pages/AgentPerformancePage";

export default function App() {
  return (
    <DashboardDataProvider>
      <BrowserRouter>
        <Routes>
          <Route path="/" element={<CareDashboardPage />} />
          <Route path="/agentPerformance" element={<AgentPerformancePage />} />
          <Route path="*" element={<Navigate to="/" replace />} />
        </Routes>
      </BrowserRouter>
    </DashboardDataProvider>
  );
}
