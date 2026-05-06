import React, { useEffect, useState } from "react";
import { BarChart3 } from "lucide-react";
import { Home } from "./pages/Home";
import { EfficiencyPage } from "./pages/EfficiencyPage";
import { WeeklyPage } from "./pages/WeeklyPage";
import { MiniAppsPage } from "./pages/MiniAppsPage";
import { FeedbackPage } from "./pages/FeedbackPage";

export function App() {
  const routes = {
    home: "/",
    efficiency: "/efficiency",
    weekly: "/weekly",
    miniApps: "/mini-programs",
    feedback: "/feedback"
  };
  const routePages = Object.fromEntries(Object.entries(routes).map(([key, path]) => [path, key]));
  const normalizePath = (path) => (path !== "/" ? path.replace(/\/+$/, "") : path);
  const pageFromPath = () => routePages[normalizePath(window.location.pathname)] || "home";
  const [page, setPage] = useState(pageFromPath);
  const pageOrder = ["home", "efficiency", "weekly", "miniApps", "feedback"];

  useEffect(() => {
    const onPopState = () => setPage(pageFromPath());
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, []);

  function navigate(nextPage) {
    const nextPath = routes[nextPage] || routes.home;
    if (window.location.pathname !== nextPath) {
      window.history.pushState({}, "", nextPath);
    }
    setPage(nextPage);
  }

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
          <button className={page === "home" ? "nav active" : "nav"} onClick={() => navigate("home")}>Overview</button>
          <button className={page === "efficiency" ? "nav active" : "nav"} onClick={() => navigate("efficiency")}>Efficiency Dashboard</button>
          <button className={page === "weekly" ? "nav active" : "nav"} onClick={() => navigate("weekly")}>Order / Unit Analysis</button>
          <button className={page === "miniApps" ? "nav active" : "nav"} onClick={() => navigate("miniApps")}>Mini Programs</button>
          <button className={page === "feedback" ? "nav active" : "nav"} onClick={() => navigate("feedback")}>Feedback</button>
        </div>
      </aside>
      <main>
        <div hidden={page !== "home"}>
          <Home onNavigate={navigate} />
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
