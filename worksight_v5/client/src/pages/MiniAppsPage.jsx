import React, { useRef, useState } from "react";
import {
  ArrowLeft,
  Boxes,
  CheckCircle2,
  ChevronDown,
  ChevronUp,
  Download,
  FileSpreadsheet,
  Loader2,
  ScanBarcode,
  Search,
  Sparkles,
  TriangleAlert
} from "lucide-react";
import { API } from "../constants";
import { GlassSelect } from "../components/controls";
import Barcode from "react-barcode";

const NEW_PRODUCT_TOOL_URL = "/downloads/WorkSight-NewProduct-Automation.exe";
const PICKING_WAREHOUSES = [
  { value: "2", label: "EWR-SM-2-US(C0000002427)" },
  { value: "1", label: "EWR-LG-1-US(C0000000389)" },
  { value: "5", label: "EWR-LG-5-US(C0000009943)" }
];

function PickingExceptionTool({ onBack }) {
  const [barcode, setBarcode] = useState("");
  const [containerNo, setContainerNo] = useState("");
  const [warehouse, setWarehouse] = useState("2");
  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [showLogs, setShowLogs] = useState(false);
  const containerRef = useRef(null);

  async function searchLocation() {
    const cleanBarcode = barcode.trim();
    const cleanContainerNo = containerNo.trim();

    setResult(null);
    setError("");
    setShowLogs(false);

    if (!cleanBarcode) {
      setError("Please scan or enter the product barcode.");
      return;
    }

    if (!cleanContainerNo) {
      setError("Please scan or enter the container number.");
      return;
    }

    setLoading(true);

    try {
      const response = await fetch(`${API}/api/picking-exception/search-location`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          barcode: cleanBarcode,
          container_no: cleanContainerNo,
          warehouse
        })
      });

      const data = await response.json();

      if (!response.ok || !data.success) {
        setError(data.message || "Query failed.");
        return;
      }

      setResult(data);
    } catch (err) {
      setError(err.message || "Request failed.");
    } finally {
      setLoading(false);
    }
  }

  function clearForm() {
    setBarcode("");
    setContainerNo("");
    setResult(null);
      setError("");
    setShowLogs(false);
  }

  const isExtra = result?.location?.toLowerCase() === "extra";

  return (
    <div className="panel picking-tool">
      <div className="tool-head">
        <button className="ghost-btn" onClick={onBack}>
          <ArrowLeft size={16} /> Back
        </button>
        <div className="panel-title">
          <ScanBarcode size={20} />
          <span>Picking Exception</span>
        </div>
      </div>

      <div className="picking-layout">
        <div className="automation-card picking-form-card">
          <div>
            <h2>Check Storage Space</h2>
            <p className="hint-line">For over-picked items, scan the barcode and container to find the best location to return one unit.</p>
          </div>

          <label className="ws-field">
            <span>Warehouse</span>
            <GlassSelect value={warehouse} options={PICKING_WAREHOUSES} onChange={setWarehouse} className="warehouse-select" />
          </label>

          <label className="ws-field">
            <span>Product Barcode</span>
            <input
              autoFocus
              value={barcode}
              onChange={(event) => setBarcode(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  containerRef.current?.focus();
                }
              }}
              placeholder="Scan or enter barcode"
            />
          </label>

          <label className="ws-field">
            <span>Container No</span>
            <input
              ref={containerRef}
              value={containerNo}
              onChange={(event) => setContainerNo(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  searchLocation();
                }
              }}
              placeholder="Scan or enter container no"
            />
          </label>

          <div className="button-row picking-actions">
            <button className="primary-btn" onClick={searchLocation} disabled={loading}>
              {loading ? <Loader2 className="spin-icon" size={18} /> : <Search size={18} />}
              {loading ? "Searching" : "Check Location"}
            </button>
            <button className="ghost-btn" onClick={clearForm} disabled={loading}>
              Clear
            </button>
          </div>

          {error && (
            <div className="error">
              <TriangleAlert size={16} /> {error}
            </div>
          )}
        </div>

        <div className={result ? "picking-result-card active" : "picking-result-card"}>
          {!result && (
            <div className="empty-state">
              <Boxes size={34} />
              <strong>Ready for scan</strong>
              <span>The recommended location will appear here after WMS responds.</span>
            </div>
          )}

          {result && (
            <>
              <div className="result-status">
                {isExtra ? <TriangleAlert size={22} /> : <CheckCircle2 size={22} />}
                <span>{isExtra ? "Inventory Exception" : "Recommended Location"}</span>
              </div>
              <strong className="recommendation">{result.location}</strong>
              {!isExtra && <span className="result-qty">Inventory qty: {result.qty ?? "-"}</span>}
            </>
          )}
        </div>
      </div>

      {result?.logs?.length > 0 && (
        <div className="automation-card log-card">
          <button className="log-toggle" onClick={() => setShowLogs((value) => !value)}>
            <span>Query Log</span>
            {showLogs ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
          </button>
          {showLogs && (
            <div className="picking-log">
              {result.logs.map((line, index) => (
                <div key={`${line}-${index}`}>{line}</div>
              ))}
            </div>
          )}
          {!showLogs && (
            <div className="log-collapsed">
              {result.logs.length} lines hidden
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function BarcodeGenerator({ onBack }) {
  const [value, setValue] = useState("");
  const barcodeRef = useRef(null);

  function downloadBarcode() {
  const svg = barcodeRef.current?.querySelector("svg");
  if (!svg) return;

  const serializer = new XMLSerializer();
  const svgString = serializer.serializeToString(svg);

  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");

  const img = new Image();
  img.onload = function () {
    canvas.width = img.width;
    canvas.height = img.height;
    ctx.drawImage(img, 0, 0);

    const link = document.createElement("a");
    link.download = `${value || "barcode"}.png`;
    link.href = canvas.toDataURL("image/png");
    link.click();
  };

  img.src = "data:image/svg+xml;base64," + btoa(svgString);
}

  return (
    <div className="panel">
      <div className="tool-head">
        <button className="ghost-btn" onClick={onBack}>
          <ArrowLeft size={16} /> Back
        </button>

        <div className="panel-title">
          <ScanBarcode size={20} />
          <span>Barcode Generator</span>
        </div>
      </div>

      <div className="picking-layout">
  
        {/* 左边：操作区 */}
        <div className="automation-card picking-form-card">
          <div>
            <h2>Generate Barcode</h2>
            <p className="hint-line">
              Enter a barcode or location code to generate a label.
            </p>
          </div>

          <label className="ws-field">
            <span>Input Number / Barcode</span>

            <input
              value={value}
              onChange={(e) => setValue(e.target.value)}
            />
          </label>

          <div className="button-row picking-actions">
            <button
              className="primary-btn"
              onClick={downloadBarcode}
              disabled={!value}
            >
              <Download size={18} />
              Download PNG
            </button>

            <button
              className="ghost-btn"
              onClick={() =>
                alert("🚧 Batch import is under development.\nPlease leave a request in Feedback.")
              }
            >
              Batch Import
            </button>
          </div>
        </div>

        {/* 右边：结果区 */}
        <div className={value ? "picking-result-card active" : "picking-result-card"}>
          
          {!value && (
            <div className="empty-state">
              <ScanBarcode size={34} />
              <strong>Ready to generate</strong>
              <span>The barcode will appear here.</span>
            </div>
          )}

          {value && (
            <>
              <div className="result-status">
                <CheckCircle2 size={22} />
                <span>Generated Barcode</span>
              </div>

              <div ref={barcodeRef} style={{ textAlign: "center" }}>
                <Barcode
                  value={value}
                  width={2.2}
                  height={140}
                  fontSize={16}
                />
              </div>
            </>
          )}

        </div>
      </div>
    </div>
  );
}

function AutomaticInventoryTransfer({ onBack }) {
  const [devices, setDevices] = useState([]);
  const [selectedDevice, setSelectedDevice] = useState("");
  const [cellA, setCellA] = useState("A1-R1-L1-B1");
  const [cellB, setCellB] = useState("A1-R1-L1-B2");
  const [limit, setLimit] = useState("400");
  const [running, setRunning] = useState(false);
  const [logs, setLogs] = useState([
    "Ready. Connect Android device and click Refresh Devices."
  ]);

  function addLog(line) {
    setLogs((current) => [
      ...current,
      `[${new Date().toLocaleTimeString()}] ${line}`
    ]);
  }

  function refreshDevices() {
    // 先不接后端，先做前端假数据
    const mockDevices = ["USB_DEVICE_001"];
    setDevices(mockDevices);
    setSelectedDevice(mockDevices[0]);
    addLog("Devices refreshed.");
  }

  function startTransfer() {
    if (!selectedDevice) {
      addLog("Please select a device first.");
      return;
    }

    if (!cellA.trim() || !cellB.trim()) {
      addLog("Cell A and Cell B are required.");
      return;
    }

    setRunning(true);
    addLog(`Start relocation loop: ${cellA} → ${cellB}`);
    addLog(`Loop count: ${limit || "0"} ${limit === "0" ? "(infinite)" : ""}`);
    addLog("Frontend ready. Backend connection will be added later.");
  }

  function stopTransfer() {
    setRunning(false);
    addLog("Stop requested.");
  }

  return (
    <div className="panel picking-tool">
      <div className="tool-head">
        <button className="ghost-btn" onClick={onBack}>
          <ArrowLeft size={16} /> Back
        </button>

        <div className="panel-title">
          <Boxes size={20} />
          <span>Automatic Inventory Transfer</span>
        </div>
      </div>

      <div className="picking-layout">
        <div className="automation-card picking-form-card">
          <div>
            <h2>Relocation Control</h2>
            <p className="hint-line">
              In-Warehouse -&gt; change -&gt; Transfer -&gt; Relocation by Cell (Please select English as your iWMS language.)
            </p>
          </div>

          <label className="ws-field">
            <span>Device</span>
            <select
              value={selectedDevice}
              onChange={(e) => setSelectedDevice(e.target.value)}
            >
              <option value="">No device selected</option>
              {devices.map((device) => (
                <option key={device} value={device}>
                  {device}
                </option>
              ))}
            </select>
          </label>

          <button className="ghost-btn" onClick={refreshDevices}>
            Refresh Devices
          </button>

          <label className="ws-field">
            <span>Cell A</span>
            <input
              value={cellA}
              onChange={(e) => setCellA(e.target.value)}
              placeholder="A1-R1-L1-B1"
            />
          </label>

          <label className="ws-field">
            <span>Cell B</span>
            <input
              value={cellB}
              onChange={(e) => setCellB(e.target.value)}
              placeholder="A1-R1-L1-B2"
            />
          </label>

          <label className="ws-field">
            <span>Loop Count</span>
            <input
              type="number"
              min="0"
              value={limit}
              onChange={(e) => setLimit(e.target.value)}
              placeholder="0 means infinite"
            />
          </label>

          <div className="button-row picking-actions">
            <button
              className="primary-btn"
              onClick={startTransfer}
              disabled={running}
            >
              {running ? <Loader2 className="spin-icon" size={18} /> : <CheckCircle2 size={18} />}
              {running ? "Running" : "Start"}
            </button>

            <button
              className="ghost-btn"
              onClick={stopTransfer}
              disabled={!running}
            >
              Stop
            </button>
          </div>
        </div>

        <div className="picking-result-card active transfer-log-card">
          <div className="result-status">
            {running ? <Loader2 className="spin-icon" size={22} /> : <Boxes size={22} />}
            <span>{running ? "Running Transfer" : "Transfer Console"}</span>
          </div>

          <div className="transfer-summary">
            <strong>{cellA || "-"}</strong>
            <span>→</span>
            <strong>{cellB || "-"}</strong>
          </div>

          <div className="picking-log transfer-log">
            {logs.map((line, index) => (
              <div key={`${line}-${index}`}>{line}</div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

function GoodsPickingOutbound({ onBack }) {
  const [devices, setDevices] = useState([]);
  const [selectedDevice, setSelectedDevice] = useState("");
  const [cellA, setCellA] = useState("A1-R1-L1-B1");
  const [cellB, setCellB] = useState("A1-R1-L1-B2");
  const [limit, setLimit] = useState("400");
  const [running, setRunning] = useState(false);
  const [logs, setLogs] = useState([
    "Ready. Connect Android device and click Refresh Devices."
  ]);

  function addLog(line) {
    setLogs((current) => [
      ...current,
      `[${new Date().toLocaleTimeString()}] ${line}`
    ]);
  }

  function refreshDevices() {
    // 先不接后端，先做前端假数据
    const mockDevices = ["USB_DEVICE_001"];
    setDevices(mockDevices);
    setSelectedDevice(mockDevices[0]);
    addLog("Devices refreshed.");
  }

  function startTransfer() {
    if (!selectedDevice) {
      addLog("Please select a device first.");
      return;
    }

    if (!cellA.trim() || !cellB.trim()) {
      addLog("Cell A and Cell B are required.");
      return;
    }

    setRunning(true);
    addLog(`Start relocation loop: ${cellA} → ${cellB}`);
    addLog(`Loop count: ${limit || "0"} ${limit === "0" ? "(infinite)" : ""}`);
    addLog("Frontend ready. Backend connection will be added later.");
  }

  function stopTransfer() {
    setRunning(false);
    addLog("Stop requested.");
  }

  return (
    <div className="panel picking-tool">
      <div className="tool-head">
        <button className="ghost-btn" onClick={onBack}>
          <ArrowLeft size={16} /> Back
        </button>

        <div className="panel-title">
          <Boxes size={20} />
          <span>Goods Picking for Outbound</span>
        </div>
      </div>

      <div className="picking-layout">
        <div className="automation-card picking-form-card">
          <div>
            <h2>Picking Automation Control</h2>
            <p className="hint-line">
              Outbound -&gt; Pick -&gt; Goods Pick -&gt; Scan Order Number(Please select English as your iWMS language.)
            </p>
          </div>

          <label className="ws-field">
            <span>Device</span>
            <select
              value={selectedDevice}
              onChange={(e) => setSelectedDevice(e.target.value)}
            >
              <option value="">No device selected</option>
              {devices.map((device) => (
                <option key={device} value={device}>
                  {device}
                </option>
              ))}
            </select>
          </label>

          <button className="ghost-btn" onClick={refreshDevices}>
            Refresh Devices
          </button>

          <label className="ws-field">
            <span>Cell A</span>
            <input
              value={cellA}
              onChange={(e) => setCellA(e.target.value)}
              placeholder="A1-R1-L1-B1"
            />
          </label>

          <label className="ws-field">
            <span>Cell B</span>
            <input
              value={cellB}
              onChange={(e) => setCellB(e.target.value)}
              placeholder="A1-R1-L1-B2"
            />
          </label>

          <label className="ws-field">
            <span>Loop Count</span>
            <input
              type="number"
              min="0"
              value={limit}
              onChange={(e) => setLimit(e.target.value)}
              placeholder="0 means infinite"
            />
          </label>

          <div className="button-row picking-actions">
            <button
              className="primary-btn"
              onClick={startTransfer}
              disabled={running}
            >
              {running ? <Loader2 className="spin-icon" size={18} /> : <CheckCircle2 size={18} />}
              {running ? "Running" : "Start"}
            </button>

            <button
              className="ghost-btn"
              onClick={stopTransfer}
              disabled={!running}
            >
              Stop
            </button>
          </div>
        </div>

        <div className="picking-result-card active transfer-log-card">
          <div className="result-status">
            {running ? <Loader2 className="spin-icon" size={22} /> : <Boxes size={22} />}
            <span>{running ? "Running Transfer" : "Transfer Console"}</span>
          </div>

          <div className="transfer-summary">
            <strong>{cellA || "-"}</strong>
            <span>→</span>
            <strong>{cellB || "-"}</strong>
          </div>

          <div className="picking-log transfer-log">
            {logs.map((line, index) => (
              <div key={`${line}-${index}`}>{line}</div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}


function NewProductTool({ onBack }) {
  return (
    <div className="panel">
      <div className="tool-head">
        <button className="ghost-btn" onClick={onBack}>
          <ArrowLeft size={16} /> Back
        </button>
        <div className="panel-title">
          <FileSpreadsheet size={20} />
          <span>New Product Maintenance Automation</span>
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
            <p className="hint-line">Product code, length, width, height, weight</p>
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
            <h2>Note</h2>
            <p className="hint-line">The automation runs locally on your computer. </p>
            <div className="empty">Still under development. Coming Soon.</div>
          </div>
        </div>
      </div>
    </div>
  );
}

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
          <button className="feature-tile" onClick={() => setActiveTool("picking-exception")}>
            <ScanBarcode size={32} />
            <h2>Picking Exception</h2>
            <p>Find where to put back one extra item after an over-pick.</p>
          </button>
          <button className="feature-tile" onClick={() => setActiveTool("new-product-maintenance")}>
            <FileSpreadsheet size={32} />
            <h2>New Product Maintenance</h2>
            <p>New product maintenance automation</p>
          </button>
          <button className="feature-tile" onClick={() => setActiveTool("barcode")}>
            <ScanBarcode size={32} />
            <h2>Barcode Generator</h2>
            <p>Generate barcode from input</p>
          </button>
          <button className="feature-tile" onClick={() => setActiveTool("goods-picking")}>
            <ScanBarcode size={32} />
            <h2>Goods Picking for Outbound</h2>
            <p>PDA automatic data processing</p>
          </button>
          <button className="feature-tile" onClick={() => setActiveTool("automatic-transfer")}>
            <Boxes size={32} />
            <h2>Inventory transfer For In-Warehouse</h2>
            <p>PDA automatic data processing</p>
          </button>
          {["Batch Update"].map((name) => (
            <button className="feature-tile coming-soon-tile" key={name} disabled>
              <Sparkles size={32} />
              <h2>{name}</h2>
              <p>Coming soon</p>
            </button>
          ))}
        </div>
      )}

      {activeTool === "picking-exception" && <PickingExceptionTool onBack={() => setActiveTool("")} />}
      {activeTool === "new-product-maintenance" && <NewProductTool onBack={() => setActiveTool("")} />}
      {activeTool === "barcode" && <BarcodeGenerator onBack={() => setActiveTool("")} />}
      {activeTool === "automatic-transfer" && (
        <AutomaticInventoryTransfer onBack={() => setActiveTool("")} />
      )}
      {activeTool === "goods-picking" && (
        <GoodsPickingOutbound onBack={() => setActiveTool("")} />
      )}
    </section>
  );
}
