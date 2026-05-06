import React, { useRef, useState } from "react";
import { ArrowLeft, Boxes, CheckCircle2, ChevronDown, ChevronUp, Loader2, ScanBarcode, Search, TriangleAlert } from "lucide-react";
import { API } from "../../constants";
import { GlassSelect } from "../controls";
import { PICKING_WAREHOUSES } from "./constants";

export function PickingExceptionTool({ onBack }) {
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
