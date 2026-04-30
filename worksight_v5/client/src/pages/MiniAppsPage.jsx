import React, { useState } from "react";
import { ArrowLeft, Download, FileSpreadsheet, Sparkles } from "lucide-react";

const NEW_PRODUCT_TOOL_URL = "/downloads/WorkSight-NewProduct-Automation.exe";

export function MiniAppsPage() {
  const [activeTool, setActiveTool] = useState("");

  return (
    <section className="page">
      <header className="page-head">
        <div>
          <h1>Mini Programs</h1>
          <p>Small warehouse operation tools</p>
        </div>
      </header>

      {!activeTool && (
        <div className="landing-grid mini-program-grid">
          <button className="feature-tile" onClick={() => setActiveTool("new-product-maintenance")}>
            <FileSpreadsheet size={32} />
            <h2>新品维护自动化</h2>
            <p>New product maintenance automation</p>
          </button>
          {["Inventory Helper", "Label Tool", "Exception Tracker", "Batch Update"].map((name) => (
            <button className="feature-tile coming-soon-tile" key={name} disabled>
              <Sparkles size={32} />
              <h2>{name}</h2>
              <p>Coming soon</p>
            </button>
          ))}
        </div>
      )}

      {activeTool === "new-product-maintenance" && (
        <div className="panel">
          <div className="tool-head">
            <button className="ghost-btn" onClick={() => setActiveTool("")}>
              <ArrowLeft size={16} /> Back
            </button>
            <div className="panel-title">
              <FileSpreadsheet size={20} />
              <span>新品维护自动化</span>
            </div>
          </div>
          <div className="automation-download">
            <div className="download-hero">
              <FileSpreadsheet size={40} />
              <div>
                <h2>Download Windows Automation Tool</h2>
                <p>Run the local tool on your computer to open iWMS and fill new product fields from Excel.</p>
              </div>
              <a className="primary-btn download-btn" href={NEW_PRODUCT_TOOL_URL} download>
                <Download size={18} /> Download EXE
              </a>
            </div>

            <div className="automation-layout">
              <div className="automation-card">
                <h2>Required Excel Columns</h2>
                <p className="hint-line">商品编码, 长, 宽, 高, 重量</p>
              </div>

              <div className="automation-card">
                <h2>How It Works</h2>
                <ol className="step-list">
                  <li>Download and run the Windows tool.</li>
                  <li>Select your Excel file when prompted.</li>
                  <li>The tool opens iWMS in a local browser window.</li>
                  <li>First-time login requires your iWMS account and password.</li>
                  <li>The tool fills product fields automatically, but does not click Save.</li>
                </ol>
              </div>

              <div className="automation-card wide">
                <h2>Safety Note</h2>
                <p className="hint-line">The automation runs locally on your computer. Please review every product form before saving in iWMS.</p>
                <div className="empty">The download will work after the EXE is placed at public/downloads/WorkSight-NewProduct-Automation.exe.</div>
              </div>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
