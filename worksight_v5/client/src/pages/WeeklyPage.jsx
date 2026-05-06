import React, { useEffect, useMemo, useRef, useState } from "react";
import { Download } from "lucide-react";
import { API } from "../constants";
import { DateRangeQuery, GlassSelect, LaborHoursInput, Metric, ProgressBar, UploadBox } from "../components/controls";
import { DailyDetailTable } from "../components/tables";
import { exportRows } from "../utils/export";
import { formatUploadError } from "../utils/formatters";

const WAREHOUSE_OPTIONS = [
  { value: "1", label: "Warehouse 1" },
  { value: "2", label: "Warehouse 2" },
  { value: "5", label: "Warehouse 5" }
];

const TARGET_UPPH_BY_WAREHOUSE = {
  "1": 34.57,
  "2": 37.11,
  "5": 11.3
};

export function WeeklyPage() {
  const [volume, setVolume] = useState(null);
  const [laborMode, setLaborMode] = useState("excel");
  const [isc, setIsc] = useState(null);
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [manualHours, setManualHours] = useState({});
  const [unitRange, setUnitRange] = useState({ from: "", to: "" });
  const [warehouse, setWarehouse] = useState("5");
  const requestSeq = useRef(0);
  const latestDailyRef = useRef([]);

  useEffect(() => {
    latestDailyRef.current = data?.daily || [];
  }, [data?.daily]);

  useEffect(() => {
    const seq = ++requestSeq.current;
    const hasAnyUpload = Boolean(volume || (laborMode === "excel" && isc));
    if (!hasAnyUpload) {
      setError("");
      setLoading(false);
      return;
    }

    const controller = new AbortController();
    const timer = setTimeout(async () => {
      setError("");
      setLoading(true);
      const form = new FormData();
      if (volume) form.append("volume", volume);
      if (laborMode === "excel" && isc) form.append("isc", isc);
      if (!volume && latestDailyRef.current.length) form.append("existingDaily", JSON.stringify(latestDailyRef.current));

      try {
        const res = await fetch(`${API}/api/weekly/analyze`, {
          method: "POST",
          body: form,
          signal: controller.signal
        });
        const json = await res.json();
        if (!res.ok) throw json;
        if (seq !== requestSeq.current) return;
        setData(json);
      } catch (e) {
        if (e.name !== "AbortError" && seq === requestSeq.current) setError(formatUploadError(e));
      } finally {
        if (!controller.signal.aborted && seq === requestSeq.current) setLoading(false);
      }
    }, 250);

    return () => {
      controller.abort();
      clearTimeout(timer);
    };
  }, [volume, isc, laborMode]);

  useEffect(() => {
    if (laborMode === "manual") {
      setIsc(null);
    }
  }, [laborMode]);

  async function queryUnitData() {
    setError("");
    setLoading(true);
    try {
      const res = await fetch(`${API}/api/weekly/query-unit`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ from: unitRange.from, to: unitRange.to, warehouse })
      });
      const json = await res.json();
      if (!res.ok) throw json;
      setData((current) => ({
        ...(current || {}),
        ...json
      }));
    } catch (e) {
      setError(formatUploadError(e));
    } finally {
      setLoading(false);
    }
  }

  const dailyRows = useMemo(() => {
    if (!data?.daily) return [];
    if (laborMode !== "manual") return data.daily;
    return data.daily.map((row) => {
      const date = row.业务日期;
      const rawHours = manualHours[date] ?? "";
      const hours = Number(rawHours);
      const hasHours = rawHours !== "" && Number.isFinite(hours) && hours > 0;
      return {
        ...row,
        总工时: rawHours,
        UPPH: hasHours ? Number((Number(row.件量 || 0) / hours).toFixed(2)) : ""
      };
    });
  }, [data, laborMode, manualHours]);

  const kpi = data?.kpi || {};
  const totalOrders = data ? Math.round(kpi.totalOrders || 0).toLocaleString() : "";
  const totalUnits = data ? Math.round(kpi.totalUnits || 0).toLocaleString() : "";
  const targetUpph = TARGET_UPPH_BY_WAREHOUSE[warehouse] || "";

  return (
    <section className="page">
      <header className="page-head">
        <h1>Weekly Order and Unit Analysis</h1>
        <p>Summarize orders, units, labor hours, and UPPH by business date</p>
      </header>

      <div className="upload-grid weekly-upload-grid">
        <UploadBox
          title="Upload Unit Data"
          caption="iWMS sales order query"
          onChange={setVolume}
          actionSlot={<DateRangeQuery value={unitRange} onChange={setUnitRange} onQuery={queryUnitData} disabled={loading || !unitRange.from || !unitRange.to} />}
        />
        <LaborHoursInput mode={laborMode} setMode={setLaborMode} onFileChange={setIsc} />
      </div>
      {loading && <ProgressBar value={72} label="Analyzing uploaded files..." />}
      {error && <div className="error">{error}</div>}

      <div className="kpi-grid">
        <Metric label="TOTAL ORDERS" value={totalOrders} />
        <Metric label="TOTAL UNITS" value={totalUnits} />
        <WarehouseMetric value={warehouse} onChange={setWarehouse} />
        <Metric label="TARGET UPPH" value={targetUpph} />
      </div>

      {data && (
        <div className="panel">
          <div className="table-head">
            <h2>Daily Detail</h2>
            <button className="ghost-btn" onClick={() => exportRows(dailyRows, "Daily Detail", "order-unit-upph.xlsx")}><Download size={16} /> Download Excel</button>
          </div>
          <DailyDetailTable
            rows={dailyRows}
            lowUpph={targetUpph || null}
            manual={laborMode === "manual"}
            manualHours={manualHours}
            setManualHours={setManualHours}
          />
        </div>
      )}
    </section>
  );
}

function WarehouseMetric({ value, onChange }) {
  return (
    <div className="metric warehouse-metric">
      <span>WAREHOUSE</span>
      <GlassSelect
        value={value}
        options={WAREHOUSE_OPTIONS}
        onChange={onChange}
        className="warehouse-kpi-select"
      />
    </div>
  );
}
