import React from "react";
import Plot from "react-plotly.js";
import { TYPE_COLORS } from "../constants";
import { palette, previousDate, timeOf } from "../utils/formatters";

export function Donut({ ratio, work, attendance, note = "" }) {
  return (
    <div className="metric donut-card">
      <Plot
        data={[{ type: "pie", labels: ["Work", "Idle"], values: [work, Math.max(0, attendance - work)], hole: 0.65, marker: { colors: ["#2563EB", "#D7E3FF"] }, textinfo: "none", hovertemplate: "%{label}: %{percent:.1%} (%{value:.1f}h)<extra></extra>" }]}
        layout={{ height: note ? 140 : 160, margin: { l: 0, r: 0, t: 0, b: 0 }, annotations: [{ text: `${ratio.toFixed(1)}%`, x: 0.5, y: 0.5, showarrow: false, font: { size: 22, color: "#2563EB" } }], showlegend: false }}
        config={{ displayModeBar: false, responsive: true }}
        useResizeHandler
        style={{ width: "100%" }}
      />
      {note && <small>{note}</small>}
    </div>
  );
}

export function filterShift(rows, shift, group) {
  return rows.filter((r) => r.shift === shift && (group === "All Groups" || r.group === group));
}

export function Gantt({ rows, allRows, date, shift, summary, history }) {
  const clippedRows = rows.map((row) => clipGanttRowToClock(row, summary)).filter(Boolean);
  const clippedAllRows = (allRows || rows).map((row) => clipGanttRowToClock(row, summary)).filter(Boolean);
  if (!clippedRows.length) return <div className="empty">No {shift} shift data.</div>;
  const stats = buildPersonStats(clippedAllRows);
  const names = [...new Set(clippedRows.map((r) => r.name))].sort((a, b) => (summary[b]?.ratio || 0) - (summary[a]?.ratio || 0));
  const traces = Object.keys(TYPE_COLORS).map((type) => {
    const items = clippedRows.filter((r) => r.type === type);
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

function clipGanttRowToClock(row, summary) {
  const personSummary = summary?.[row.name];
  const clockIn = personSummary?.in;
  const clockOut = personSummary?.out;
  if (!clockIn || !clockOut || clockIn === "-" || clockOut === "-" || clockIn === "缺失" || clockOut === "缺失") {
    return row;
  }

  const start = new Date(row.start);
  const end = new Date(row.end);
  const shiftStart = dateWithClock(row.date, clockIn);
  let shiftEnd = dateWithClock(row.date, clockOut);
  if (!Number.isFinite(start.getTime()) || !Number.isFinite(end.getTime()) || !shiftStart || !shiftEnd) return row;
  if (shiftEnd <= shiftStart) shiftEnd = new Date(shiftEnd.getTime() + 24 * 60 * 60 * 1000);
  if (end <= shiftStart || start >= shiftEnd) return null;

  const clippedStart = start < shiftStart ? shiftStart : start;
  const clippedEnd = end > shiftEnd ? shiftEnd : end;
  if (clippedStart >= clippedEnd) return null;
  return {
    ...row,
    start: clippedStart.toISOString(),
    end: clippedEnd.toISOString(),
    duration: (clippedEnd - clippedStart) / 3600000
  };
}

function dateWithClock(date, clock) {
  const match = String(clock).match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return null;
  const [year, month, day] = String(date).split("-").map(Number);
  if (!year || !month || !day) return null;
  return new Date(year, month - 1, day, Number(match[1]), Number(match[2]), Number(match[3] || 0));
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

export function IndirectChart({ rows }) {
  if (!rows.length) return <div className="empty">No indirect time data for this date.</div>;
  const cats = [...new Set(rows.map((r) => r.操作分类))];
  const traces = cats.map((cat, idx) => {
    const items = rows.filter((r) => r.操作分类 === cat);
    return { type: "bar", orientation: "h", name: cat, y: items.map((r) => r.employee_no), x: items.map((r) => new Date(r.结束时间) - new Date(r.开始时间)), base: items.map((r) => new Date(r.开始时间)), marker: { color: palette(idx) } };
  });
  return <Plot data={traces} layout={{ barmode: "stack", height: Math.max(450, new Set(rows.map((r) => r.employee_no)).size * 32), margin: { l: 25, r: 60, t: 20, b: 35 }, yaxis: { autorange: "reversed" }, xaxis: { title: "Time", type: "date" }, plot_bgcolor: "white", paper_bgcolor: "white" }} config={{ responsive: true }} useResizeHandler style={{ width: "100%" }} />;
}

export function PersonEfficiencyChart({ rows, personName }) {
  const chartRows = (rows || [])
    .filter((row) => row.姓名 === personName)
    .sort((a, b) => String(a.日期).localeCompare(String(b.日期)));
  if (!chartRows.length) return null;
  const x = chartRows.map((row) => row.日期);
  const traces = [
    {
      type: "scatter",
      mode: "lines+markers",
      name: "Non-Burst Picking Efficiency",
      x,
      y: chartRows.map((row) => Number(row.拣非爆品效率) || 0),
      line: { color: "#2563eb", width: 3 },
      marker: { size: 8 }
    },
    {
      type: "scatter",
      mode: "lines+markers",
      name: "Total Efficiency",
      x,
      y: chartRows.map((row) => Number(row.总效率) || 0),
      line: { color: "#16a34a", width: 2 },
      marker: { size: 7 }
    },
    {
      type: "scatter",
      mode: "lines+markers",
      name: "Effective Hours",
      x,
      y: chartRows.map((row) => Number(row.有效工时) || 0),
      yaxis: "y2",
      line: { color: "#f59e0b", width: 2 },
      marker: { size: 7 }
    }
  ];
  return (
    <div className="chart-strip">
      <Plot
        data={traces}
        layout={{
          height: 320,
          margin: { l: 56, r: 24, t: 18, b: 46 },
          plot_bgcolor: "white",
          paper_bgcolor: "white",
          xaxis: { title: "Date", type: "category" },
          yaxis: { title: "Efficiency", rangemode: "tozero" },
          yaxis2: { title: "Effective Hours", overlaying: "y", side: "right", rangemode: "tozero" },
          legend: { orientation: "h", x: 0, y: 1.12 }
        }}
        config={{ responsive: true, displayModeBar: false }}
        useResizeHandler
        style={{ width: "100%" }}
      />
    </div>
  );
}
