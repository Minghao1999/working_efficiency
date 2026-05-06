import React from "react";
import { ArrowLeft, ClipboardList, Download, FileSpreadsheet, Repeat } from "lucide-react";
import { NEW_PRODUCT_TOOL_URL, PICKING_TOOL_URL, TRANSFER_TOOL_URL } from "./constants";

export function NewProductTool({ onBack }) {
  return (
    <DownloadTool
      onBack={onBack}
      icon={<FileSpreadsheet size={20} />}
      heroIcon={<FileSpreadsheet size={40} />}
      title="New Product Maintenance Automation"
      heroTitle="Download Windows Automation Tool"
      heroText="Run the local tool on your computer to open iWMS and fill new product fields from Excel."
      href={NEW_PRODUCT_TOOL_URL}
      cards={[
        { title: "Required Excel Columns", body: "Product code, length, width, height, weight" },
        {
          title: "How It Works",
          steps: [
            "Download and run the Windows tool.",
            "Select your Excel file when prompted.",
            "The tool opens iWMS in a local browser window.",
            "First-time login requires your iWMS account and password.",
            "The tool fills product fields automatically, but does not click Save."
          ]
        },
        { title: "Note", body: "The automation runs locally on your computer.", extra: <div className="empty">Still under development. Coming Soon.</div>, wide: true }
      ]}
    />
  );
}

export function AutomaticInventoryTransfer({ onBack }) {
  return (
    <DownloadTool
      onBack={onBack}
      icon={<Repeat size={20} />}
      heroIcon={<Repeat size={40} />}
      title="Inventory Transfer (In-Warehouse)"
      heroTitle="Download Inventory Transfer Tool"
      heroText="Run the local tool to automatically perform relocation tasks on PDA via ADB."
      href={TRANSFER_TOOL_URL}
      cards={[
        { title: "Function", body: "Automatically execute relocation loops between two storage cells." },
        { title: "How It Works", steps: ["Download and run the tool.", "Connect PDA via USB.", "Select device and input Cell A / Cell B.", "Start loop execution."] },
        { title: "Note", body: "Requires ADB environment and PDA USB debugging enabled.", wide: true }
      ]}
    />
  );
}

export function GoodsPickingOutbound({ onBack }) {
  return (
    <DownloadTool
      onBack={onBack}
      icon={<ClipboardList size={20} />}
      heroIcon={<ClipboardList size={40} />}
      title="Goods Picking for Outbound"
      heroTitle="Download Picking Automation Tool"
      heroText="Run the local tool to automate outbound picking operations on PDA."
      href={PICKING_TOOL_URL}
      cards={[
        { title: "Function", body: "Automatically process outbound picking tasks and scan orders." },
        { title: "How It Works", steps: ["Download and run the tool.", "Connect PDA via USB.", "Login iWMS (English mode).", "Start picking automation."] },
        { title: "Note", body: "Ensure PDA is connected and debugging mode is enabled.", wide: true }
      ]}
    />
  );
}

function DownloadTool({ onBack, icon, heroIcon, title, heroTitle, heroText, href, cards }) {
  return (
    <div className="panel">
      <div className="tool-head">
        <button className="ghost-btn" onClick={onBack}>
          <ArrowLeft size={16} /> Back
        </button>
        <div className="panel-title">
          {icon}
          <span>{title}</span>
        </div>
      </div>

      <div className="automation-download">
        <div className="download-hero">
          {heroIcon}
          <div>
            <h2>{heroTitle}</h2>
            <p>{heroText}</p>
          </div>
          <a className="primary-btn download-btn" href={href} download>
            <Download size={18} /> Download EXE
          </a>
        </div>

        <div className="automation-layout">
          {cards.map((card) => (
            <div className={card.wide ? "automation-card wide" : "automation-card"} key={card.title}>
              <h2>{card.title}</h2>
              {card.body && <p className="hint-line">{card.body}</p>}
              {card.steps && (
                <ol className="step-list">
                  {card.steps.map((step) => <li key={step}>{step}</li>)}
                </ol>
              )}
              {card.extra}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
