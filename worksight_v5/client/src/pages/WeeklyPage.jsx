import React, { useEffect, useMemo, useRef, useState } from "react";
import { Download, RotateCcw } from "lucide-react";
import { API } from "../constants";
import { LaborHoursInput, Metric, ProgressBar, SelectLine, UploadBox } from "../components/controls";
import { DailyDetailTable, EditablePersonTable } from "../components/tables";
import { PersonEfficiencyChart, PickingGanttChart } from "../components/charts";
import { exportRows } from "../utils/export";
import { formatUploadError, personDeleteKey } from "../utils/formatters";

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
  const requestSeq = useRef(0);

  useEffect(() => {
    const seq = ++requestSeq.current;
    const hasAnyUpload = Boolean(volume || pick || (laborMode === "excel" && isc));
    if (!hasAnyUpload) {
      setData(null);
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

  return (
    <section className="page">
      <header className="page-head">
        <h1>Weekly Order and Unit Analysis</h1>
        <p>Summarize orders, units, labor hours, UPPH, and per-person picking efficiency by business date</p>
      </header>

      <div className="upload-grid">
        <UploadBox title="Upload Unit Data" caption="iWMS 销售单综合查询" onChange={setVolume} />
        <LaborHoursInput mode={laborMode} setMode={setLaborMode} onFileChange={setIsc} />
        <UploadBox title="Upload Picking Data" caption="iWMS 拣货结果查询" onChange={setPick} />
      </div>
      {loading && <ProgressBar value={72} label="Analyzing uploaded files..." />}
      {error && <div className="error">{error}</div>}

      {data && (
        <>
          <div className="kpi-grid">
            <Metric label="TOTAL ORDERS" value={Math.round(data.kpi.totalOrders).toLocaleString()} />
            <Metric label="TOTAL UNITS" value={Math.round(data.kpi.totalUnits).toLocaleString()} />
            <Metric label="WAREHOUSE" value={hasAttendance ? data.kpi.warehouseName : "Please upload attendance table"} />
            <Metric label="TARGET UPPH" value={hasAttendance ? data.kpi.targetUpph : "Please upload attendance table"} />
          </div>
          <div className="panel">
            <div className="table-head">
              <h2>Daily Detail</h2>
              <button className="ghost-btn" onClick={() => exportRows(dailyRows, "Daily Detail", "order-unit-upph.xlsx")}><Download size={16} /> Download Excel</button>
            </div>
            <DailyDetailTable
              rows={dailyRows}
              lowUpph={hasAttendance ? data.kpi.targetUpph : null}
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
