import React, { useState } from "react";
import { BarChart3 } from "lucide-react";
import { Home } from "./pages/Home";
import { EfficiencyPage } from "./pages/EfficiencyPage";
import { WeeklyPage } from "./pages/WeeklyPage";
import { MiniAppsPage } from "./pages/MiniAppsPage";
import { FeedbackPage } from "./pages/FeedbackPage";

export function App() {
  const [page, setPage] = useState("home");
  const pageOrder = ["home", "efficiency", "weekly", "miniApps", "feedback"];
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
        <div className="nav-group" style={{ "--active-index": pageOrder.indexOf(page), "--nav-count": pageOrder.length }}>
          <button className={page === "home" ? "nav active" : "nav"} onClick={() => setPage("home")}>Overview</button>
          <button className={page === "efficiency" ? "nav active" : "nav"} onClick={() => setPage("efficiency")}>Efficiency Dashboard</button>
          <button className={page === "weekly" ? "nav active" : "nav"} onClick={() => setPage("weekly")}>Order / Unit Analysis</button>
          <button className={page === "miniApps" ? "nav active" : "nav"} onClick={() => setPage("miniApps")}>Mini Programs</button>
          <button className={page === "feedback" ? "nav active" : "nav"} onClick={() => setPage("feedback")}>Feedback</button>
        </div>
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
        <div hidden={page !== "miniApps"}>
          <MiniAppsPage />
        </div>
        <div hidden={page !== "feedback"}>
          <FeedbackPage />
        </div>
      </main>
    </div>
  );
}
