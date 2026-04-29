import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import Plot from "react-plotly.js";
import { BarChart3, Boxes, Download, FileSpreadsheet, Package, Timer, Users } from "lucide-react";
import "./styles.css";

const API = "http://127.0.0.1:3001";
const TYPE_COLORS = {
  work: "#7EC398",
  overnight: "#8B7ABD",
  idle: "#C6D7EB",
  break: "#D9E6AF"
};

const COLUMN_LABELS = {
  业务日期: "Business Date",
  单量: "Orders",
  件量: "Units",
  总工时: "Total Hours",
  出勤人数: "Headcount",
  仓库: "Warehouse",
  目标UPPH: "Target UPPH",
  日期: "Date",
  工号: "Employee ID",
  姓名: "Name",
  总件数: "Total Units",
  件数: "Units",
  有效工时: "Effective Hours",
  考勤时长: "Attendance Hours",
  有效工时占比: "Effective Hours Ratio",
  拣非爆品效率: "Non-Burst Picking Efficiency",
  总效率: "Total Efficiency",
  低于目标: "Below Target",
  删除: "Delete",
  上班时间: "Clock In",
  下班时间: "Clock Out",
  employee_no: "Employee ID",
  real_name: "Name",
  操作分类: "Category",
  开始时间: "Start",
  结束时间: "End"
};

function App() {
  const [page, setPage] = useState("home");
  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <BarChart3 size={28} />
          <div>
            <strong>WorkSight Pro</strong>
            <span>Warehouse Operations Analytics</span>
          </div>
        </div>
        <button className={page === "home" ? "nav active" : "nav"} onClick={() => setPage("home")}>Overview</button>
        <button className={page === "efficiency" ? "nav active" : "nav"} onClick={() => setPage("efficiency")}>Efficiency Dashboard</button>
        <button className={page === "weekly" ? "nav active" : "nav"} onClick={() => setPage("weekly")}>Order / Unit Analysis</button>
      </aside>
      <main>
        <div className="page-view" hidden={page !== "home"}>
          <Home onNavigate={setPage} />
        </div>
        <div className="page-view" hidden={page !== "efficiency"}>
          <EfficiencyPage />
        </div>
        <div className="page-view" hidden={page !== "weekly"}>
          <WeeklyPage />
        </div>
      </main>
    </div>
  );
}

function Home({ onNavigate }) {
  return (
    <section className="page">
      <header className="page-head">
        <h1>WorkSight Pro</h1>
        <p>Warehouse operations analytics for productivity, orders, units, and trends</p>
      </header>
      <div className="landing-grid">
        <button className="feature-tile" onClick={() => onNavigate("efficiency")}>
          <Users size={32} />
          <h2>Efficiency Dashboard</h2>
          <p>Review employee productivity, Gantt timelines, and Idle / Work distribution</p>
        </button>
        <button className="feature-tile" onClick={() => onNavigate("weekly")}>
          <Package size={32} />
          <h2>Order / Unit Analysis</h2>
          <p>Summarize daily order and unit trends with Excel export</p>
        </button>
      </div>
    </section>
  );
}

function EfficiencyPage() {
  const [files, setFiles] = useState([]);
  const [data, setData] = useState(null);
  const [date, setDate] = useState("");
  const [tab, setTab] = useState("morning");
  const [morningGroup, setMorningGroup] = useState("All Groups");
  const [eveningGroup, setEveningGroup] = useState("All Groups");
  const [category, setCategory] = useState("All Categories");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [progress, setProgress] = useState({ value: 0, label: "" });

  function addEfficiencyFiles(nextFiles) {
    setFiles((current) => mergeFiles(current, nextFiles));
  }

  function removeEfficiencyFile(fileKey) {
    setFiles((current) => current.filter((file) => fileIdentity(file) !== fileKey));
  }

  async function submit() {
    setError("");
    setLoading(true);
    setProgress({ value: 8, label: "Uploading files..." });
    const progressTimers = [
      setTimeout(() => setProgress({ value: 22, label: "Classifying uploaded workbooks..." }), 450),
      setTimeout(() => setProgress({ value: 42, label: "Reading Excel sheets..." }), 1200),
      setTimeout(() => setProgress({ value: 62, label: "Building employee timelines..." }), 2200),
      setTimeout(() => setProgress({ value: 78, label: "Calculating KPIs and history..." }), 3600),
      setTimeout(() => setProgress({ value: 90, label: "Preparing dashboard data..." }), 5200)
    ];
    const form = new FormData();
    files.forEach((file) => form.append("files", file));
    try {
      const res = await fetch(`${API}/api/efficiency/analyze`, { method: "POST", body: form });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "Dashboard generation failed");
      setProgress({ value: 100, label: "Dashboard ready." });
      setData(json);
      setDate(json.dates[0] || "");
    } catch (e) {
      setError(translateError(e.message));
      setProgress({ value: 100, label: "Failed." });
    } finally {
      progressTimers.forEach(clearTimeout);
      setLoading(false);
    }
  }

  const day = data?.days?.[date];
  const dayRows = useMemo(() => (data?.timeline || []).filter((r) => r.date === date), [data, date]);
  const groups = useMemo(() => ["All Groups", ...new Set(dayRows.map((r) => r.group).filter(Boolean))].sort(), [dayRows]);
  const indirectCategories = useMemo(() => ["All Categories", ...new Set((day?.indirect || []).map((r) => r.操作分类).filter(Boolean))].sort(), [day]);
  const completeness = data?.completeness || {};

  return (
    <section className="page">
      <header className="page-head">
        <h1>Efficiency Dashboard</h1>
        <p>Advanced productivity analytics dashboard for ISC, iAMS, volume, and punch data</p>
      </header>

      <div className="panel">
        <div className="panel-title">
          <FileSpreadsheet size={20} />
          <span>Upload Data Files</span>
        </div>
        <div className="source-list">
          <div>📄 ISC Task Data：来自 ISC员工操作时长（需要使用Global Export） · Required</div>
          <div>👥 iAMS Attendance：来自 iAMS班次明细（实际上下班时间） · Optional</div>
          <div>📦 Volume Data：来自 iWMS销售单综合查询（打包完成时间 & 件数） · Optional</div>
          <div>🧾 Punch Data：来自 iAMS打卡流水（进出仓记录） · Optional</div>
        </div>
        <FilePicker multiple accept=".xlsx,.xls" files={files} onChange={addEfficiencyFiles} />
        <SelectedFiles files={files} onRemove={removeEfficiencyFile} />
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
              <select value={date} onChange={(e) => setDate(e.target.value)}>
                {data.dates.map((d) => <option key={d}>{d}</option>)}
              </select>
            </label>
          </div>

          <div className="kpi-grid five">
            <Metric icon={<Boxes />} label="WAREHOUSE" value={day.kpi.warehouse} />
            <Metric icon={<Users />} label="TOTAL WORKERS" value={day.kpi.totalWorkers} />
            <Metric icon={<Timer />} label="TOTAL ATTENDANCE" value={`${day.kpi.totalAttendance.toFixed(1)}h`} />
            <Donut ratio={day.kpi.efficiencyRatio} work={day.kpi.totalWork} attendance={day.kpi.totalAttendance} />
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
    </section>
  );
}

function WeeklyPage() {
  const [volume, setVolume] = useState(null);
  const [laborMode, setLaborMode] = useState("excel");
  const [isc, setIsc] = useState(null);
  const [pick, setPick] = useState(null);
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [personDate, setPersonDate] = useState("");
  const [deleted, setDeleted] = useState(new Set());
  const [markedForDelete, setMarkedForDelete] = useState(new Set());
  const [manualHours, setManualHours] = useState({});

  useEffect(() => {
    if (!volume) {
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
      form.append("volume", volume);
      if (laborMode === "excel" && isc) form.append("isc", isc);
      if (pick) form.append("pick", pick);

      try {
        const res = await fetch(`${API}/api/weekly/analyze`, {
          method: "POST",
          body: form,
          signal: controller.signal
        });
        const json = await res.json();
        if (!res.ok) throw new Error(json.error || "Analysis failed");
        setData(json);
        const firstDate = [...new Set((json.personEfficiency || []).map((r) => r.日期))][0] || "";
        setPersonDate(firstDate);
        setDeleted(new Set());
        setMarkedForDelete(new Set());
      } catch (e) {
        if (e.name !== "AbortError") setError(translateError(e.message));
      } finally {
        if (!controller.signal.aborted) setLoading(false);
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

  const personDates = [...new Set((data?.personEfficiency || []).map((r) => r.日期))].sort();
  const personRows = (data?.personEfficiency || [])
    .filter((r) => r.日期 === personDate)
    .map((r) => {
      const id = `${r.日期}|${r.工号}`;
      const { 低于目标, ...row } = r;
      return { ...row, 删除: markedForDelete.has(id), _deleted: deleted.has(id) };
    });

  function applyDeleteSort() {
    setDeleted(new Set(markedForDelete));
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

          {!!personDates.length && (
            <div className="panel">
              <div className="table-head">
                <h2>Per-Person Daily Picking Efficiency</h2>
                <button className="ghost-btn" onClick={applyDeleteSort}>Apply Delete</button>
              </div>
              <SelectLine label="Select Date" value={personDate} options={personDates} onChange={setPersonDate} />
              <EditablePersonTable
                rows={personRows}
                markedForDelete={markedForDelete}
                setMarkedForDelete={setMarkedForDelete}
              />
            </div>
          )}
        </>
      )}
    </section>
  );
}

function UploadBox({ title, caption, disabled = false, onChange }) {
  const [fileName, setFileName] = useState("");
  return (
    <label className={disabled ? "upload-box disabled" : "upload-box"}>
      <strong>{title}</strong>
      <span>{caption}</span>
      <FilePicker
        accept=".xlsx,.xls"
        disabled={disabled}
        files={fileName ? [{ name: fileName }] : []}
        onChange={(files) => {
          const file = files[0] || null;
          setFileName(file?.name || "");
          onChange(file);
        }}
      />
    </label>
  );
}

function LaborHoursInput({ mode, setMode, onFileChange }) {
  return (
    <div className="upload-box">
      <strong>Labor Hours</strong>
      <span>ISC 出勤工时(选择Global Export).</span>
      <div className="segmented">
        <button type="button" className={mode === "excel" ? "active" : ""} onClick={() => setMode("excel")}>Excel Upload</button>
        <button type="button" className={mode === "manual" ? "active" : ""} onClick={() => setMode("manual")}>Manual Entry</button>
      </div>
      {mode === "excel" ? (
        <FilePicker accept=".xlsx,.xls" files={[]} onChange={(files) => onFileChange(files[0] || null)} />
      ) : (
        <div className="hint-line">Enter Total Hours directly in the Daily Detail table.</div>
      )}
    </div>
  );
}

function FilePicker({ multiple = false, accept, disabled = false, files = [], onChange }) {
  const label = files.length ? files.map((file) => file.name).join(", ") : "No file selected";
  return (
    <label className={disabled ? "file-picker disabled" : "file-picker"}>
      <span>Choose File{multiple ? "s" : ""}</span>
      <input
        type="file"
        multiple={multiple}
        accept={accept}
        disabled={disabled}
        onChange={(e) => onChange([...e.target.files])}
      />
      <em>{label}</em>
    </label>
  );
}

function SelectedFiles({ files, onRemove }) {
  if (!files.length) return null;
  return (
    <div className="selected-files">
      {files.map((file) => {
        const key = fileIdentity(file);
        return (
          <div className="selected-file" key={key}>
            <span>{file.name}</span>
            <button type="button" onClick={() => onRemove(key)}>Remove</button>
          </div>
        );
      })}
    </div>
  );
}

function fileIdentity(file) {
  return `${file.name}|${file.size ?? 0}|${file.lastModified ?? 0}`;
}

function mergeFiles(current, nextFiles) {
  const map = new Map(current.map((file) => [fileIdentity(file), file]));
  for (const file of nextFiles) {
    map.set(fileIdentity(file), file);
  }
  return [...map.values()];
}

function Metric({ icon, label, value }) {
  return (
    <div className="metric">
      {icon && <span className="metric-icon">{icon}</span>}
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function ProgressBar({ value, label }) {
  return (
    <div className="progress-block" aria-live="polite">
      <div className="progress-meta">
        <span>{label}</span>
        <strong>{Math.round(value)}%</strong>
      </div>
      <div className="progress-track">
        <div className="progress-fill" style={{ width: `${Math.min(100, Math.max(0, value))}%` }} />
      </div>
    </div>
  );
}

function Donut({ ratio, work, attendance }) {
  return (
    <div className="metric donut-card">
      <Plot
        data={[{ type: "pie", labels: ["Work", "Idle"], values: [work, Math.max(0, attendance - work)], hole: 0.65, marker: { colors: ["#2563EB", "#D7E3FF"] }, textinfo: "none", hovertemplate: "%{label}: %{percent:.1%} (%{value:.1f}h)<extra></extra>" }]}
        layout={{ height: 160, margin: { l: 0, r: 0, t: 0, b: 0 }, annotations: [{ text: `${ratio.toFixed(1)}%`, x: 0.5, y: 0.5, showarrow: false, font: { size: 22, color: "#2563EB" } }], showlegend: false }}
        config={{ displayModeBar: false, responsive: true }}
        useResizeHandler
        style={{ width: "100%" }}
      />
    </div>
  );
}

function Tabs({ tabs, value, onChange }) {
  return <div className="tabs">{tabs.map(([id, label]) => <button key={id} className={value === id ? "active" : ""} onClick={() => onChange(id)}>{label}</button>)}</div>;
}

function ChartPanel({ title, caption, children }) {
  return <div className="panel chart-panel"><h2>{title}</h2>{caption && <p className="hint-line">{caption}</p>}{children}</div>;
}

function SelectLine({ label, value, options, onChange }) {
  return <label className="select-line">{label}<select value={value} onChange={(e) => onChange(e.target.value)}>{options.map((o) => <option key={o}>{o}</option>)}</select></label>;
}

function filterShift(rows, shift, group) {
  return rows.filter((r) => r.shift === shift && (group === "All Groups" || r.group === group));
}

function Gantt({ rows, allRows, date, shift, summary, history }) {
  if (!rows.length) return <div className="empty">No {shift} shift data.</div>;
  const stats = buildPersonStats(allRows || rows);
  const names = [...new Set(rows.map((r) => r.name))].sort((a, b) => (summary[b]?.ratio || 0) - (summary[a]?.ratio || 0));
  const traces = Object.keys(TYPE_COLORS).map((type) => {
    const items = rows.filter((r) => r.type === type);
    return {
      type: "bar",
      orientation: "h",
      name: type,
      y: items.map((r) => r.name),
      x: items.map((r) => new Date(r.end) - new Date(r.start)),
      base: items.map((r) => new Date(r.start)),
      marker: { color: TYPE_COLORS[type] },
      customdata: items.map((r) => {
        const person = stats[r.name] || {};
        return [
          r.type,
          r.name,
          timeOf(r.start),
          timeOf(r.end),
          r.duration,
          summary[r.name]?.ratio || 0,
          summary[r.name]?.in || "-",
          summary[r.name]?.out || "-",
          person.idleHours || 0,
          person.workHours || 0,
          person.attendanceHours || 0
        ];
      }),
      hovertemplate:
        "type: %{customdata[0]}<br>" +
        "name: %{customdata[1]}<br>" +
        "clock in: %{customdata[6]}<br>" +
        "clock out: %{customdata[7]}<br>" +
        "task start: %{customdata[2]}<br>" +
        "task end: %{customdata[3]}<br>" +
        "task duration: %{customdata[4]:.2f}h<br>" +
        "idle total: %{customdata[8]:.2f}h<br>" +
        "work total: %{customdata[9]:.2f}h<br>" +
        "attendance total: %{customdata[10]:.2f}h<br>" +
        "efficiency: %{customdata[5]:.1f}%<extra></extra>"
    };
  }).filter((t) => t.y.length);
  const prev = previousDate(date);
  const annotations = names.map((name) => {
    const ratio = summary[name]?.ratio || 0;
    const prevRatio = history?.[prev]?.[name];
    const arrow = prevRatio == null ? "" : ratio > prevRatio ? " ↑" : ratio < prevRatio ? " ↓" : "";
    const color = arrow.includes("↑") ? "#16A34A" : arrow.includes("↓") ? "#DC2626" : "#0F172A";
    return { x: 1.01, y: name, xref: "paper", yref: "y", text: `${ratio.toFixed(1)}%${arrow}`, showarrow: false, xanchor: "left", font: { size: 11, color } };
  });
  return (
    <Plot
      data={traces}
      layout={{ barmode: "stack", height: Math.max(450, names.length * 32), margin: { l: 190, r: 165, t: 20, b: 35 }, plot_bgcolor: "white", paper_bgcolor: "white", yaxis: { autorange: "reversed", categoryorder: "array", categoryarray: names, automargin: true }, xaxis: { title: "Time", type: "date" }, legend: { x: 1.09, y: 1, xanchor: "left", yanchor: "top" }, annotations }}
      config={{ responsive: true }}
      useResizeHandler
      style={{ width: "100%" }}
    />
  );
}

function buildPersonStats(rows) {
  return rows.reduce((acc, row) => {
    const current = acc[row.name] || { idleHours: 0, workHours: 0, attendanceHours: 0 };
    current.attendanceHours += Number(row.duration) || 0;
    if (row.type === "idle") current.idleHours += Number(row.duration) || 0;
    if (row.type === "work" || row.type === "overnight") current.workHours += Number(row.duration) || 0;
    acc[row.name] = current;
    return acc;
  }, {});
}

function IndirectChart({ rows }) {
  if (!rows.length) return <div className="empty">No indirect time data for this date.</div>;
  const cats = [...new Set(rows.map((r) => r.操作分类))];
  const traces = cats.map((cat, idx) => {
    const items = rows.filter((r) => r.操作分类 === cat);
    return { type: "bar", orientation: "h", name: cat, y: items.map((r) => r.employee_no), x: items.map((r) => new Date(r.结束时间) - new Date(r.开始时间)), base: items.map((r) => new Date(r.开始时间)), marker: { color: palette(idx) } };
  });
  return <Plot data={traces} layout={{ barmode: "stack", height: Math.max(450, new Set(rows.map((r) => r.employee_no)).size * 32), margin: { l: 25, r: 60, t: 20, b: 35 }, yaxis: { autorange: "reversed" }, xaxis: { title: "Time", type: "date" }, plot_bgcolor: "white", paper_bgcolor: "white" }} config={{ responsive: true }} useResizeHandler style={{ width: "100%" }} />;
}

function DataTable({ rows, lowUpph, highlightWait }) {
  if (!rows?.length) return <div className="empty">No data</div>;
  const keys = Object.keys(rows[0]).filter((k) => !k.startsWith("_"));
  return (
    <div className="table-wrap">
      <table>
        <thead><tr>{keys.map((k) => <th key={k}>{columnLabel(k)}</th>)}</tr></thead>
        <tbody>
          {rows.map((row, idx) => {
            const upph = Number(row.UPPH);
            const low = lowUpph && Number.isFinite(upph) && upph < lowUpph;
            return <tr key={idx} className={low ? "low-row" : ""}>{keys.map((k) => <td key={k} className={highlightWait && k === "Wait" && parseInt(row[k]) > 5 ? "danger-text" : ""}>{formatCell(row[k])}</td>)}</tr>;
          })}
        </tbody>
      </table>
    </div>
  );
}

function DailyDetailTable({ rows, lowUpph, manual, manualHours, setManualHours }) {
  if (!rows?.length) return <div className="empty">No data</div>;
  const keys = Object.keys(rows[0]).filter((k) => !k.startsWith("_"));
  return (
    <div className="table-wrap">
      <table>
        <thead><tr>{keys.map((k) => <th key={k}>{columnLabel(k)}</th>)}</tr></thead>
        <tbody>
          {rows.map((row, idx) => {
            const upph = Number(row.UPPH);
            const low = lowUpph && Number.isFinite(upph) && upph < lowUpph;
            return (
              <tr key={row.业务日期 || idx} className={low ? "low-row" : ""}>
                {keys.map((k) => (
                  <td key={k}>
                    {manual && k === "总工时" ? (
                      <input
                        className="table-number-input"
                        type="number"
                        min="0"
                        step="0.01"
                        inputMode="decimal"
                        placeholder="Enter hours"
                        value={manualHours[row.业务日期] ?? ""}
                        onChange={(e) => {
                          const value = e.target.value;
                          setManualHours((current) => ({ ...current, [row.业务日期]: value }));
                        }}
                      />
                    ) : (
                      formatCell(row[k])
                    )}
                  </td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function EditablePersonTable({ rows, markedForDelete, setMarkedForDelete }) {
  const [sort, setSort] = useState(null);

  if (!rows.length) return <div className="empty">No data</div>;
  const sorted = sortRows(rows, sort);
  const keys = Object.keys(sorted[0]).filter((key) => !key.startsWith("_"));
  const onSort = (key) => {
    setSort((current) => {
      if (!current || current.key !== key) return { key, direction: "asc" };
      if (current.direction === "asc") return { key, direction: "desc" };
      return null;
    });
  };

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {keys.map((k) => (
              <th key={k}>
                <button className="sort-header" onClick={() => onSort(k)}>
                  <span>{columnLabel(k)}</span>
                  <span className="sort-mark">{sort?.key === k ? (sort.direction === "asc" ? "↑" : "↓") : "↕"}</span>
                </button>
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {sorted.map((row) => {
            const id = `${row.日期}|${row.工号}`;
            return <tr key={id} className={row._deleted ? "deleted-row" : ""}>{keys.map((k) => <td key={k}>{k === "删除" ? <input type="checkbox" checked={markedForDelete.has(id)} onChange={(e) => {
              const next = new Set(markedForDelete);
              e.target.checked ? next.add(id) : next.delete(id);
              setMarkedForDelete(next);
            }} /> : formatCell(row[k])}</td>)}</tr>;
          })}
        </tbody>
      </table>
    </div>
  );
}

function sortRows(rows, sort) {
  return [...rows].sort((a, b) => {
    if (!sort) {
      return Number(a._deleted) - Number(b._deleted) || Number(b.拣非爆品效率 || 0) - Number(a.拣非爆品效率 || 0);
    }

    const result = compareValues(a[sort.key], b[sort.key]);
    return sort.direction === "asc" ? result : -result;
  });
}

function compareValues(a, b) {
  if (a == null && b == null) return 0;
  if (a == null) return 1;
  if (b == null) return -1;

  if (typeof a === "boolean" || typeof b === "boolean") {
    return Number(a) - Number(b);
  }

  const aNum = Number(a);
  const bNum = Number(b);
  if (Number.isFinite(aNum) && Number.isFinite(bNum)) {
    return aNum - bNum;
  }

  const aDate = Date.parse(a);
  const bDate = Date.parse(b);
  if (Number.isFinite(aDate) && Number.isFinite(bDate)) {
    return aDate - bDate;
  }

  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: "base" });
}

async function exportRows(rows, sheetName, fileName) {
  const res = await fetch(`${API}/api/export`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ rows, sheetName, fileName }) });
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}

function formatCell(v) {
  if (v == null) return "";
  if (typeof v === "number") return Number.isInteger(v) ? v.toLocaleString() : v.toFixed(2);
  if (typeof v === "boolean") return v ? "Yes" : "No";
  if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}T/.test(v)) return v.replace("T", " ").slice(0, 19);
  return String(v);
}

function columnLabel(key) {
  return COLUMN_LABELS[key] || key;
}

function translateError(message) {
  return String(message || "")
    .replace("缺少任务单明细（df1）", "Missing task detail file")
    .replace("缺少班次数据（df2）", "Missing attendance shift file")
    .replace("缺少件量数据（df3）", "Missing volume file")
    .replace("缺少打卡数据（df4）", "Missing punch log file")
    .replace("请上传销售单综合数据", "Please upload the sales order file")
    .replace("缺少字段", "Missing fields")
    .replace("找不到 sheet", "Sheet not found");
}

function timeOf(v) {
  const d = new Date(v);
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}:${String(d.getSeconds()).padStart(2, "0")}`;
}

function previousDate(date) {
  const d = new Date(`${date}T00:00:00`);
  d.setDate(d.getDate() - 1);
  return d.toISOString().slice(0, 10);
}

function palette(i) {
  return ["#2563EB", "#16A34A", "#F59E0B", "#DC2626", "#7C3AED", "#0891B2", "#DB2777"][i % 7];
}

createRoot(document.getElementById("root")).render(<App />);
