import React, { useState } from "react";
import { BarChart3 } from "lucide-react";
import { Home } from "./pages/Home";
import { EfficiencyPage } from "./pages/EfficiencyPage";
import { WeeklyPage } from "./pages/WeeklyPage";

export function App() {
  const [page, setPage] = useState("home");
  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <BarChart3 size={28} />
          <div>
            <strong>WorkSight Pro</strong>
            <span>JD Logistics</span>
          </div>
        </div>
        <button className={page === "home" ? "nav active" : "nav"} onClick={() => setPage("home")}>Overview</button>
        <button className={page === "efficiency" ? "nav active" : "nav"} onClick={() => setPage("efficiency")}>Efficiency Dashboard</button>
        <button className={page === "weekly" ? "nav active" : "nav"} onClick={() => setPage("weekly")}>Order / Unit Analysis</button>
      </aside>
      <main>
        <div hidden={page !== "home"}>
          <Home onNavigate={setPage} />
        </div>
        <div hidden={page !== "efficiency"}>
          <EfficiencyPage />
        </div>
        <div hidden={page !== "weekly"}>
          <WeeklyPage />
        </div>
      </main>
    </div>
  );
}
