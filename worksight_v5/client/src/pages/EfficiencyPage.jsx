import React, { useMemo, useState } from "react";
import { Boxes, FileSpreadsheet, Package, Timer, Users } from "lucide-react";
import { API } from "../constants";
import { ChartPanel, FilePicker, Metric, ProgressBar, SelectLine, SelectedFileList, Tabs } from "../components/controls";
import { DataTable } from "../components/tables";
import { Donut, Gantt, IndirectChart, filterShift } from "../components/charts";
import { fileIdentity, isLikelyVolumeFile, mergeEfficiencyVolumeDelta, mergeFiles } from "../utils/files";
import { translateError } from "../utils/formatters";

export function EfficiencyPage() {
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

  const day = data?.days?.[date];
  const dayRows = useMemo(() => (data?.timeline || []).filter((r) => r.date === date), [data, date]);
  const groups = useMemo(() => ["All Groups", ...new Set(dayRows.map((r) => r.group).filter(Boolean))].sort(), [dayRows]);
  const indirectCategories = useMemo(() => ["All Categories", ...new Set((day?.indirect || []).map((r) => r.操作分类).filter(Boolean))].sort(), [day]);
  const completeness = data?.completeness || {};

  return (
    <section className="page">
      <header className="page-head">
        <h1>Efficiency Dashboard</h1>
        <p>Advanced productivity analytics dashboard for ISC, iAMS, and iWMS</p>
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
              <select value={date} onChange={(e) => setDate(e.target.value)}>
                {data.dates.map((d) => <option key={d}>{d}</option>)}
              </select>
            </label>
          </div>

          <div className="kpi-grid five">
            <Metric icon={<Boxes />} label="WAREHOUSE" value={day.kpi.warehouse} />
            <Metric icon={<Users />} label="TOTAL WORKERS" value={day.kpi.totalWorkers} />
            <Metric
              icon={<Timer />}
              label="TOTAL ATTENDANCE"
              value={`${day.kpi.totalAttendance.toFixed(1)}h`}
              note={!completeness.attendance ? "不准确，需要上传班次明细表" : ""}
            />
            <Donut
              ratio={day.kpi.efficiencyRatio}
              work={day.kpi.totalWork}
              attendance={day.kpi.totalAttendance}
              note={!completeness.attendance ? "不准确，需要上传班次明细表" : ""}
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
    </section>
  );
}
