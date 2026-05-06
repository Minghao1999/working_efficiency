import React, { useRef, useState } from "react";
import { ArrowLeft, CheckCircle2, Download, ScanBarcode } from "lucide-react";
import Barcode from "react-barcode";

export function BarcodeGenerator({ onBack }) {
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

    img.src = `data:image/svg+xml;base64,${btoa(svgString)}`;
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
        <div className="automation-card picking-form-card">
          <div>
            <h2>Generate Barcode</h2>
            <p className="hint-line">Enter a barcode or location code to generate a label.</p>
          </div>

          <label className="ws-field">
            <span>Input Number / Barcode</span>
            <input value={value} onChange={(e) => setValue(e.target.value)} />
          </label>

          <div className="button-row picking-actions">
            <button className="primary-btn" onClick={downloadBarcode} disabled={!value}>
              <Download size={18} />
              Download PNG
            </button>

            <button
              className="ghost-btn"
              onClick={() => alert("Batch import is under development.\nPlease leave a request in Feedback.")}
            >
              Batch Import
            </button>
          </div>
        </div>

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
                <Barcode value={value} width={2.2} height={140} fontSize={16} />
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
