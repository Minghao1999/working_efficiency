import React from "react";
import { Package, Users } from "lucide-react";

export function Home({ onNavigate }) {
  return (
    <section className="page">
      <header className="page-head">
        <h1>WorkSight Pro</h1>
        <p>Warehouse operations analytics for productivity, orders, units, and trends</p>
      </header>
      <div className="landing-grid">
        <button className="feature-tile" onClick={() => onNavigate("efficiency")}>
          <Users size={32} />
          <h2>Efficiency Dashboard</h2>
          <p>Review employee productivity, Gantt timelines, and Idle / Work distribution</p>
        </button>
        <button className="feature-tile" onClick={() => onNavigate("weekly")}>
          <Package size={32} />
          <h2>Order / Unit Analysis</h2>
          <p>Summarize daily order and unit trends with Excel export</p>
        </button>
      </div>
    </section>
  );
}
