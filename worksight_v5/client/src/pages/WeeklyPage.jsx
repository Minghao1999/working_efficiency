import React, { useEffect, useMemo, useRef, useState } from "react";
import { Download, RotateCcw, Search } from "lucide-react";
import { API } from "../constants";
import { GlassSelect, LaborHoursInput, Metric, ProgressBar, SelectLine, UploadBox } from "../components/controls";
import { DailyDetailTable, EditablePersonTable } from "../components/tables";
import { PersonEfficiencyChart, PickingGanttChart } from "../components/charts";
import { exportRows } from "../utils/export";
import { formatUploadError, personDeleteKey } from "../utils/formatters";

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
  const [pick, setPick] = useState(null);
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [personDate, setPersonDate] = useState("");
  const [personName, setPersonName] = useState("All People");
  const [deleted, setDeleted] = useState(new Set());
  const [manualHours, setManualHours] = useState({});
  const [unitRange, setUnitRange] = useState({ from: "", to: "" });
  const [pickRange, setPickRange] = useState({ from: "", to: "" });
  const [warehouse, setWarehouse] = useState("5");
  const requestSeq = useRef(0);
  const latestDailyRef = useRef([]);

  useEffect(() => {
    latestDailyRef.current = data?.daily || [];
  }, [data?.daily]);

  useEffect(() => {
    const seq = ++requestSeq.current;
    const hasAnyUpload = Boolean(volume || pick || (laborMode === "excel" && isc));
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
      if (pick) form.append("pick", pick);
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
        const firstDate = [...new Set((json.personEfficiency || []).map((r) => r.日期).filter(Boolean))][0] || "";
        setPersonDate(firstDate);
        setPersonName("All People");
        setDeleted(new Set());
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
  }, [volume, isc, pick, laborMode]);

  useEffect(() => {
    if (laborMode === "manual") {
      setIsc(null);
    }
  }, [laborMode]);

  const activePersonEfficiency = (data?.personEfficiency || []).filter((r) => !deleted.has(personDeleteKey(r)));
  const activePickingGantt = (data?.pickingGantt || []).filter((r) => !deleted.has(r.name) && !deleted.has(r.employeeNo));
  const personDates = [...new Set(activePersonEfficiency.map((r) => r.日期).filter(Boolean))].sort();
  const personNames = ["All People", ...new Set(activePersonEfficiency.map((r) => r.姓名).filter(Boolean))].sort();
  const selectedPersonDate = personDate || personDates[0] || "";
  const pickingGanttRows = activePickingGantt
    .filter((r) => personName !== "All People" || !selectedPersonDate || r.date === selectedPersonDate)
    .filter((r) => personName === "All People" || r.name === personName);
  const personRows = activePersonEfficiency
    .filter((r) => personName !== "All People" || !selectedPersonDate || r.日期 === selectedPersonDate)
    .filter((r) => personName === "All People" || r.姓名 === personName)
    .map((r) => {
      const { 低于目标, ...row } = r;
      return row;
    });

  function removePersonFromFilter(name) {
    if (!name || name === "All People") return;
    const nextDeleted = new Set(deleted);
    nextDeleted.add(name);
    setDeleted(nextDeleted);
    if (personName === name) setPersonName("All People");
  }

  function restoreDeletedPeople() {
    setDeleted(new Set());
  }

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
        ...json,
        personEfficiency: current?.personEfficiency || [],
        pickingGantt: current?.pickingGantt || []
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

  const hasAttendance = laborMode === "excel" && Boolean(isc);
  const kpi = data?.kpi || {};
  const totalOrders = data ? Math.round(kpi.totalOrders || 0).toLocaleString() : "";
  const totalUnits = data ? Math.round(kpi.totalUnits || 0).toLocaleString() : "";
  const targetUpph = TARGET_UPPH_BY_WAREHOUSE[warehouse] || "";

  return (
    <section className="page">
      <header className="page-head">
        <h1>Weekly Order and Unit Analysis</h1>
        <p>Summarize orders, units, labor hours, UPPH, and per-person picking efficiency by business date</p>
      </header>

      <div className="upload-grid">
        <UploadBox
          title="Upload Unit Data"
          caption="iWMS 销售单综合查询"
          onChange={setVolume}
          actionSlot={<DateRangeQuery value={unitRange} onChange={setUnitRange} onQuery={queryUnitData} disabled={loading || !unitRange.from || !unitRange.to} />}
        />
        <LaborHoursInput mode={laborMode} setMode={setLaborMode} onFileChange={setIsc} />
        <UploadBox
          title="Upload Picking Data"
          caption="iWMS 拣货结果查询"
          onChange={setPick}
          actionSlot={<DateRangeQuery value={pickRange} onChange={setPickRange} />}
        />
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
        <>
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

          {!!activePersonEfficiency.length && (
            <div className="panel">
              <div className="table-head">
                <h2>Per-Person Daily Picking Efficiency</h2>
                <div className="button-row">
                  <button className="ghost-btn" disabled={!deleted.size} onClick={restoreDeletedPeople} title="Restore deleted people"><RotateCcw size={16} /> Refresh</button>
                </div>
              </div>
              <div className="filter-row">
                <SelectLine label="Select Date" value={selectedPersonDate} options={personDates} onChange={setPersonDate} />
                <SelectLine label="Filter Person" value={personName} options={personNames} onChange={setPersonName} onRemoveOption={removePersonFromFilter} />
              </div>
              {personName !== "All People" && <PersonEfficiencyChart rows={activePersonEfficiency} personName={personName} />}
              <EditablePersonTable rows={personRows} />
            </div>
          )}
          {!!activePickingGantt.length && (
            <div className="panel">
              <div className="table-head">
                <h2>Per-Person Picking Gantt</h2>
              </div>
              <PickingGanttChart rows={pickingGanttRows} groupByDate={personName !== "All People"} />
            </div>
          )}
        </>
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

function DateRangeQuery({ value, onChange, onQuery, disabled = false }) {
  const queryDisabled = disabled || !onQuery;
  return (
    <div className="date-query">
      <span className="date-query-title">Date Range</span>
      <div className="date-query-fields">
        <label>
          <span>From</span>
          <input
            type="date"
            lang="en"
            value={value.from}
            onChange={(event) => onChange((current) => ({ ...current, from: event.target.value }))}
          />
        </label>
        <label>
          <span>To</span>
          <input
            type="date"
            lang="en"
            value={value.to}
            onChange={(event) => onChange((current) => ({ ...current, to: event.target.value }))}
          />
        </label>
      </div>
      <button type="button" className="primary-btn date-query-btn" onClick={onQuery} disabled={queryDisabled}>
        <Search size={16} /> Query
      </button>
    </div>
  );
}
