import React, { useState } from "react";
import { ClipboardList, FileSpreadsheet, Package, Repeat, ScanBarcode, Warehouse } from "lucide-react";
import { AutomaticInventoryTransfer, GoodsPickingOutbound, NewProductTool } from "../components/mini-programs/AutomationDownloadTool";
import { BarcodeGenerator } from "../components/mini-programs/BarcodeGenerator";
import { PickingExceptionTool } from "../components/mini-programs/PickingExceptionTool";
import { StorageAnalysisTool } from "../components/mini-programs/StorageAnalysisTool";

const MINI_PROGRAMS = [
  {
    id: "picking-exception",
    icon: ScanBarcode,
    title: "Picking Exception",
    description: "Find where to put back one extra item after an over-pick."
  },
  {
    id: "new-product-maintenance",
    icon: FileSpreadsheet,
    title: "New Product Maintenance",
    description: "New product maintenance automation"
  },
  {
    id: "barcode",
    icon: Package,
    title: "Barcode Generator",
    description: "Generate barcode from input"
  },
  {
    id: "goods-picking",
    icon: ClipboardList,
    title: "Goods Picking for Outbound",
    description: "PDA automatic data processing"
  },
  {
    id: "automatic-transfer",
    icon: Repeat,
    title: "Inventory transfer For In-Warehouse",
    description: "PDA automatic data processing"
  },
  {
    id: "storage-analysis",
    icon: Warehouse,
    title: "Storage analysis",
    description: "Calculate the available storage space in A1-A24(NJ5 only)"
  }
];

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
          {MINI_PROGRAMS.map((tool) => {
            const Icon = tool.icon;
            return (
              <button className="feature-tile" key={tool.id} onClick={() => setActiveTool(tool.id)}>
                <Icon size={32} />
                <h2>{tool.title}</h2>
                <p>{tool.description}</p>
              </button>
            );
          })}

        </div>
      )}

      {activeTool === "picking-exception" && <PickingExceptionTool onBack={() => setActiveTool("")} />}
      {activeTool === "new-product-maintenance" && <NewProductTool onBack={() => setActiveTool("")} />}
      {activeTool === "barcode" && <BarcodeGenerator onBack={() => setActiveTool("")} />}
      {activeTool === "automatic-transfer" && <AutomaticInventoryTransfer onBack={() => setActiveTool("")} />}
      {activeTool === "goods-picking" && <GoodsPickingOutbound onBack={() => setActiveTool("")} />}
      {activeTool === "storage-analysis" && <StorageAnalysisTool onBack={() => setActiveTool("")} />}
    </section>
  );
}
