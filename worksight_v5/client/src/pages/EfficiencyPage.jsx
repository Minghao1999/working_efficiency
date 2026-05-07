import React, { useEffect, useMemo, useRef, useState } from "react";
import { Boxes, FileSpreadsheet, Package, RotateCcw, Timer, Users } from "lucide-react";
import { API } from "../constants";
import { ChartPanel, DateRangeQuery, FilePicker, GlassSelect, Metric, ProgressBar, SelectLine, SelectedFileList, Tabs, UploadBox } from "../components/controls";
import { DataTable, EditablePersonTable } from "../components/tables";
import { Donut, Gantt, IndirectChart, PersonEfficiencyChart, PickingGanttChart, filterShift } from "../components/charts";
import { fileIdentity, isLikelyVolumeFile, mergeEfficiencyVolumeDelta, mergeFiles } from "../utils/files";
import { formatUploadError, personDeleteKey, translateError } from "../utils/formatters";

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

function dateKeysBetween(from, to) {
  if (!from || !to) return [];
  const start = new Date(`${from}T00:00:00`);
  const end = new Date(`${to}T00:00:00`);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end < start) return [];
  const dates = [];
  const current = new Date(start);
  while (current <= end) {
    dates.push(current.toISOString().slice(0, 10));
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

function buildPickingDurationMap(rows) {
  const totals = new Map();
  for (const row of rows || []) {
    if (!row?.date || row.type === "lunch") continue;
    const duration = Number(row.duration) || 0;
    const keys = [
      `${row.date}|${row.employeeNo || ""}`,
      `${row.date}|${row.name || ""}`
    ];
    for (const key of keys) {
      if (!key.endsWith("|")) totals.set(key, (totals.get(key) || 0) + duration);
    }
  }
  return totals;
}

function pickingDurationForRow(row, durationMap) {
  return durationMap.get(`${row.日期}|${row.工号}`) || durationMap.get(`${row.日期}|${row.姓名}`) || Number(row.总时长) || 0;
}

export function EfficiencyPage() {
  const [view, setView] = useState("dashboard");
  const [files, setFiles] = useState([]);
  const [sessionId] = useState(() => `eff-${Date.now()}-${Math.random().toString(16).slice(2)}`);
  const [data, setData] = useState(null);
  const [date, setDate] = useState("");
  const [tab, setTab] = useState("morning");
  const [morningGroup, setMorningGroup] = useState("All Groups");
  const [eveningGroup, setEveningGroup] = useState("All Groups");
  const [category, setCategory] = useState("All Categories");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [progress, setProgress] = useState({ value: 0, label: "" });
  const [submittedFileKeys, setSubmittedFileKeys] = useState(new Set());

  const [pick, setPick] = useState(null);
  const [pickingData, setPickingData] = useState(null);
  const [pickingLoading, setPickingLoading] = useState(false);
  const [pickingError, setPickingError] = useState("");
  const [personDate, setPersonDate] = useState("");
  const [personName, setPersonName] = useState("All People");
  const [deleted, setDeleted] = useState(new Set());
  const [pickRange, setPickRange] = useState({ from: "", to: "" });
  const [warehouse, setWarehouse] = useState("5");
  const pickingRequestSeq = useRef(0);
  const targetUpph = TARGET_UPPH_BY_WAREHOUSE[warehouse] || "";

  useEffect(() => {
    const seq = ++pickingRequestSeq.current;
    if (!pick) {
      setPickingError("");
      setPickingLoading(false);
      return;
    }

    const controller = new AbortController();
    const timer = setTimeout(async () => {
      setPickingError("");
      setPickingLoading(true);
      const form = new FormData();
      form.append("pick", pick);

      try {
        const res = await fetch(`${API}/api/weekly/analyze`, {
          method: "POST",
          body: form,
          signal: controller.signal
        });
        const json = await res.json();
        if (!res.ok) throw json;
        if (seq !== pickingRequestSeq.current) return;
        setPickingData(json);
        resetPickingFilters(json);
      } catch (e) {
        if (e.name !== "AbortError" && seq === pickingRequestSeq.current) setPickingError(formatUploadError(e));
      } finally {
        if (!controller.signal.aborted && seq === pickingRequestSeq.current) setPickingLoading(false);
      }
    }, 250);

    return () => {
      controller.abort();
      clearTimeout(timer);
    };
  }, [pick]);

  function addEfficiencyFiles(nextFiles) {
    setFiles((current) => mergeFiles(current, nextFiles));
  }

  function removeEfficiencyFile(fileKey) {
    setFiles((current) => current.filter((file) => fileIdentity(file) !== fileKey));
    setSubmittedFileKeys((current) => {
      const next = new Set(current);
      next.delete(fileKey);
      return next;
    });
  }

  async function submit() {
    setError("");
    setLoading(true);
    const form = new FormData();
    form.append("sessionId", sessionId);
    form.append("activeFiles", JSON.stringify(files.map(fileIdentity)));
    const filesToUpload = data ? files.filter((file) => !submittedFileKeys.has(fileIdentity(file))) : files;
    const volumeOnlyUpdate = data && filesToUpload.length > 0 && filesToUpload.every((file) => isLikelyVolumeFile(file));
    const progressSteps = volumeOnlyUpdate
      ? [
          { value: 8, label: "Uploading volume data..." },
          { value: 38, label: "Reading the volume workbook..." },
          { value: 68, label: "Aggregating daily units..." },
          { value: 90, label: "Updating existing dashboard..." }
        ]
      : [
          { value: 8, label: "Uploading files..." },
          { value: 22, label: "Classifying uploaded workbooks..." },
          { value: 42, label: "Reading Excel sheets..." },
          { value: 62, label: "Building employee timelines..." },
          { value: 78, label: "Calculating KPIs and history..." },
          { value: 90, label: "Still processing large Excel data..." }
        ];
    setProgress(progressSteps[0]);
    const progressTimers = progressSteps.slice(1).map((step, index) => (
      setTimeout(() => setProgress(step), [450, 1200, 2200, 3600, 5200][index] || 5200)
    ));
    if (volumeOnlyUpdate) form.append("partial", "volume");
    form.append("fileKeys", JSON.stringify(filesToUpload.map(fileIdentity)));
    filesToUpload.forEach((file) => form.append("files", file));
    try {
      const res = await fetch(`${API}/api/efficiency/analyze`, { method: "POST", body: form });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "Dashboard generation failed");
      setProgress({ value: 100, label: "Dashboard ready." });
      if (json.partial === "volume") {
        setData((current) => mergeEfficiencyVolumeDelta(current, json));
      } else {
        setData(json);
        setDate(json.dates[0] || "");
      }
      setSubmittedFileKeys(new Set(files.map(fileIdentity)));
    } catch (e) {
      setError(translateError(e.message));
      setProgress({ value: 100, label: "Failed." });
    } finally {
      progressTimers.forEach(clearTimeout);
      setLoading(false);
    }
  }

  async function queryPickingData() {
    setPickingError("");
    setPickingLoading(true);
    try {
      const res = await fetch(`${API}/api/weekly/query-picking`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ from: pickRange.from, to: pickRange.to, warehouse, targetUpph })
      });
      const json = await res.json();
      if (!res.ok) throw json;
      setPickingData((current) => ({
        ...(current || {}),
        personEfficiency: json.personEfficiency || [],
        pickingGantt: json.pickingGantt || []
      }));
      resetPickingFilters(json);
    } catch (e) {
      setPickingError(formatUploadError(e));
    } finally {
      setPickingLoading(false);
    }
  }

  function resetPickingFilters(nextData) {
    const requestedDates = dateKeysBetween(pickRange.from, pickRange.to);
    const requested = new Set(requestedDates);
    const rows = requestedDates.length
      ? (nextData.personEfficiency || []).filter((r) => requested.has(r.日期))
      : (nextData.personEfficiency || []);
    const firstDate = [...new Set(rows.map((r) => r.日期).filter(Boolean))][0] || "";
    setPersonDate(firstDate);
    setPersonName("All People");
    setDeleted(new Set());
  }

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

  const day = data?.days?.[date];
  const dayRows = useMemo(() => (data?.timeline || []).filter((r) => r.date === date), [data, date]);
  const groups = useMemo(() => ["All Groups", ...new Set(dayRows.map((r) => r.group).filter(Boolean))].sort(), [dayRows]);
  const indirectCategories = useMemo(() => ["All Categories", ...new Set((day?.indirect || []).map((r) => r.操作分类).filter(Boolean))].sort(), [day]);
  const completeness = data?.completeness || {};

  const requestedPickingDates = dateKeysBetween(pickRange.from, pickRange.to);
  const requestedPickingDateSet = new Set(requestedPickingDates);
  const activePersonEfficiency = (pickingData?.personEfficiency || [])
    .filter((r) => !requestedPickingDates.length || requestedPickingDateSet.has(r.日期))
    .filter((r) => !deleted.has(personDeleteKey(r)));
  const activePickingGantt = (pickingData?.pickingGantt || [])
    .filter((r) => !requestedPickingDates.length || requestedPickingDateSet.has(r.date))
    .filter((r) => !deleted.has(r.name) && !deleted.has(r.employeeNo));
  const pickingDurationMap = useMemo(() => buildPickingDurationMap(activePickingGantt), [activePickingGantt]);
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
      const 总时长 = pickingDurationForRow(row, pickingDurationMap);
      const 总效率 = 总时长 ? (Number(row.总件数) || 0) / 总时长 : row.总效率;
      return {
        日期: row.日期,
        工号: row.工号,
        姓名: row.姓名,
        总件数: row.总件数,
        件数: row.件数,
        有效工时: row.有效工时,
        总时长,
        考勤时长: row.考勤时长,
        有效工时占比: row.有效工时占比,
        拣非爆品效率: row.拣非爆品效率,
        总效率
      };
    });

  return (
    <section className="page">
      <header className="page-head">
        <h1>Efficiency Dashboard</h1>
        <HeaderViewToggle value={view} onChange={setView} />
      </header>

      {view === "dashboard" && (
        <>
          <div className="panel">
            <div className="panel-title">
              <FileSpreadsheet size={20} />
              <span>Upload Data Files</span>
            </div>
            <div className="source-list">
              <div>ISC Task Data: ISC 员工操作时长 (Global Export required) · Required</div>
              <div>iAMS Attendance: iAMS班次明细 · Optional</div>
              <div>Volume Data: iWMS 销售单综合查询 · Optional</div>
              <div>Punch Data: iAMS 打卡流水 · Optional</div>
              <div>算力不足，最好一个文件一个文件上传，谢谢！</div>
            </div>
            <FilePicker multiple accept=".xlsx,.xls" files={files} onChange={addEfficiencyFiles} />
            <SelectedFileList files={files} onRemove={removeEfficiencyFile} />
            <button className="primary-btn" disabled={loading || !files.length} onClick={submit}>
              {loading ? "Generating..." : "Generate Dashboard"}
            </button>
            {(loading || progress.label) && <ProgressBar value={progress.value} label={progress.label} />}
            {error && <div className="error">{error}</div>}
          </div>

          {data && day && (
            <>
              <div className="toolbar">
                <label>
                  Date
                  <GlassSelect value={date} options={data.dates} onChange={setDate} />
                </label>
              </div>

              <div className="kpi-grid five">
                <Metric icon={<Boxes />} label="WAREHOUSE" value={day.kpi.warehouse} />
                <Metric icon={<Users />} label="TOTAL WORKERS" value={day.kpi.totalWorkers} />
                <Metric
                  icon={<Timer />}
                  label="TOTAL ATTENDANCE"
                  value={`${day.kpi.totalAttendance.toFixed(1)}h`}
                  note={!completeness.attendance ? "Not accurate: upload iAMS Attendance" : ""}
                />
                <Donut
                  ratio={day.kpi.efficiencyRatio}
                  work={day.kpi.totalWork}
                  attendance={day.kpi.totalAttendance}
                  note={!completeness.attendance ? "Not accurate: upload iAMS Attendance" : ""}
                />
                <Metric icon={<Package />} label="TOTAL VOLUME" value={completeness.volume ? `${day.kpi.volume.toLocaleString()} pcs` : "Incomplete: upload Volume Data"} />
              </div>

              <Tabs tabs={[["morning", "Morning Shift"], ["evening", "Evening Shift"], ["first", "First Action Analysis"], ["noop", "No Operation Analysis"], ["indirect", "Indirect Time"]]} value={tab} onChange={setTab} />

              {tab === "morning" && (
                <ChartPanel title="Morning Shift Timeline">
                  <SelectLine label="Filter Group - Morning" value={morningGroup} options={groups} onChange={setMorningGroup} />
                  <Gantt rows={filterShift(dayRows, "morning", morningGroup)} allRows={dayRows} date={date} shift="morning" summary={day.summary} history={data.history_ratios} />
                </ChartPanel>
              )}
              {tab === "evening" && (
                <ChartPanel title="Evening Shift Timeline">
                  <SelectLine label="Filter Group - Evening" value={eveningGroup} options={groups} onChange={setEveningGroup} />
                  <Gantt rows={filterShift(dayRows, "evening", eveningGroup)} allRows={dayRows} date={date} shift="evening" summary={day.summary} history={data.history_ratios} />
                </ChartPanel>
              )}
              {tab === "first" && (
                <ChartPanel title="First Action Analysis" caption="Wait time from punch-in to first task, with day-over-day trend.">
                  {completeness.attendance ? <DataTable rows={day.firstAction} highlightWait /> : <div className="empty">Incomplete: upload iAMS Attendance to calculate first action wait time.</div>}
                </ChartPanel>
              )}
              {tab === "noop" && (
                <ChartPanel title="No System Operation Analysis" caption="Employees who clocked in but had no system operations.">
                  {!completeness.attendance ? <div className="empty">Incomplete: upload iAMS Attendance to detect employees with no system operations.</div> : day.noOperation.length ? <DataTable rows={day.noOperation} /> : <div className="success">No employees found with punch records but no system operations.</div>}
                </ChartPanel>
              )}
              {tab === "indirect" && (
                <ChartPanel title="Indirect Time Timeline" caption="Indirect work time by employee, filtered by operation category.">
                  <SelectLine label="Filter Category" value={category} options={indirectCategories} onChange={setCategory} />
                  <IndirectChart rows={category === "All Categories" ? day.indirect : day.indirect.filter((r) => r.操作分类 === category)} />
                </ChartPanel>
              )}
            </>
          )}
        </>
      )}

      {view === "picking" && (
        <>
          <div className="upload-grid picking-upload-grid">
            <UploadBox
              title="Upload Picking Data"
              caption="iWMS 拣货结果查询"
              onChange={setPick}
              actionSlot={(
                <div className="picking-action-stack">
                  <label className="date-query-warehouse">
                    <span>Warehouse</span>
                    <GlassSelect
                      value={warehouse}
                      options={WAREHOUSE_OPTIONS}
                      onChange={setWarehouse}
                      className="warehouse-query-select"
                    />
                  </label>
                  <DateRangeQuery value={pickRange} onChange={setPickRange} onQuery={queryPickingData} disabled={pickingLoading || !pickRange.from || !pickRange.to} />
                </div>
              )}
            />
          </div>
          {pickingLoading && <ProgressBar value={72} label="Analyzing picking data..." />}
          {pickingError && <div className="error">{pickingError}</div>}

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
              <EditablePersonTable rows={personRows} hiddenColumns={["è€ƒå‹¤æ—¶é•¿", "æœ‰æ•ˆå·¥æ—¶å æ¯”"]} />
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

          {pickingData && !activePersonEfficiency.length && !activePickingGantt.length && (
            <div className="empty">No picking data found for the selected upload or date range.</div>
          )}
        </>
      )}
    </section>
  );
}

function HeaderViewToggle({ value, onChange }) {
  return (
    <div className="segmented compact sliding page-view-toggle" style={{ "--active-index": value === "dashboard" ? 0 : 1, "--segment-count": 2 }}>
      <button type="button" className={value === "dashboard" ? "active" : ""} onClick={() => onChange("dashboard")}>Work Efficiency</button>
      <button type="button" className={value === "picking" ? "active" : ""} onClick={() => onChange("picking")}>Picking Efficiency</button>
    </div>
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
