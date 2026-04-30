import express from "express";
import cors from "cors";
import multer from "multer";
import XLSX from "xlsx";

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json({ limit: "20mb" }));

const TARGET_MAP = { "1号仓": 34.57, "2号仓": 37.11, "5号仓": 11.3 };

function sheetRows(file, sheetName = null) {
  const wb = XLSX.read(file.buffer, { type: "buffer", cellDates: true });
  const name = sheetName || wb.SheetNames[0];
  const ws = wb.Sheets[name];
  if (!ws) throw new Error(`找不到 sheet：${name}`);
  return XLSX.utils.sheet_to_json(ws, { defval: null, raw: false }).map((row) => {
    const out = {};
    for (const [key, value] of Object.entries(row)) out[String(key).trim()] = value;
    return out;
  });
}

function workbook(file) {
  return XLSX.read(file.buffer, { type: "buffer", cellDates: true });
}

function sheetRowsFromWorkbook(wb, sheetName = null) {
  const name = sheetName || wb.SheetNames[0];
  const ws = wb.Sheets[name];
  if (!ws) throw new Error(`找不到 sheet：${name}`);
  return XLSX.utils.sheet_to_json(ws, { defval: null, raw: false }).map((row) => {
    const out = {};
    for (const [key, value] of Object.entries(row)) out[String(key).trim()] = value;
    return out;
  });
}

function firstSheetRows(file, n = 5) {
  return sheetRows(file).slice(0, n);
}

function cols(rows) {
  return rows[0] ? Object.keys(rows[0]).map((c) => String(c).trim()) : [];
}

function val(row, key) {
  return row?.[key] ?? null;
}

function colAt(row, idx) {
  const keys = Object.keys(row || {});
  return row?.[keys[idx]] ?? null;
}

function firstPresentColumn(rows, candidates) {
  const available = new Set(cols(rows));
  return candidates.find((column) => available.has(column)) || null;
}

function uploadValidationError(detail) {
  const error = new Error("上传表格校验失败");
  error.status = 400;
  error.details = [detail];
  return error;
}

function validateRows({ file, wb, sheetName = null, label, requiredColumns = [], columnGroups = [] }) {
  const missingSheets = sheetName && !wb.SheetNames.includes(sheetName) ? [sheetName] : [];
  if (missingSheets.length) {
    throw uploadValidationError({ label, file: file.originalname, missingSheets });
  }

  const rows = sheetRowsFromWorkbook(wb, sheetName);
  const available = new Set(cols(rows));
  const missingColumns = requiredColumns.filter((column) => !available.has(column));
  const missingColumnGroups = columnGroups.filter((group) => !group.some((column) => available.has(column)));
  if (missingColumns.length || missingColumnGroups.length) {
    throw uploadValidationError({ label, file: file.originalname, sheet: sheetName || wb.SheetNames[0], missingColumns, missingColumnGroups });
  }
  return rows;
}

function parseDate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;
  if (typeof v === "number") {
    const parsed = XLSX.SSF.parse_date_code(v);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, Math.floor(parsed.S));
  }
  const s = String(v).trim();
  if (!s || s === "nan" || s === "None") return null;
  if (/^\d+(\.\d+)?$/.test(s)) {
    const parsed = XLSX.SSF.parse_date_code(Number(s));
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, Math.floor(parsed.S));
  }
  const normalized = s.replace(/\//g, "-");
  const dateOnly = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (dateOnly) {
    return new Date(Number(dateOnly[1]), Number(dateOnly[2]) - 1, Number(dateOnly[3]));
  }
  const d = new Date(normalized);
  return Number.isNaN(d.getTime()) ? null : d;
}

function num(v, fallback = 0) {
  if (v == null || v === "") return fallback;
  const n = Number(String(v).replace(/,/g, "").replace("%", ""));
  return Number.isFinite(n) ? n : fallback;
}

function dayKey(d) {
  d = parseDate(d);
  if (!d) return null;
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function timeText(d) {
  d = parseDate(d);
  if (!d) return "缺失";
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}:${String(d.getSeconds()).padStart(2, "0")}`;
}

function addDays(d, days) {
  d = parseDate(d);
  if (!d) return null;
  const x = new Date(d);
  x.setDate(x.getDate() + days);
  return x;
}

function hoursBetween(a, b) {
  a = parseDate(a);
  b = parseDate(b);
  if (!a || !b || b <= a) return 0;
  return (b - a) / 3600000;
}

function adjustWorkDate(dt) {
  dt = parseDate(dt);
  if (!dt) return null;
  return dt.getHours() < 6 ? dayKey(addDays(dt, -1)) : dayKey(dt);
}

function businessDate(dt) {
  dt = parseDate(dt);
  if (!dt) return null;
  return dt.getHours() < 5 ? dayKey(addDays(dt, -1)) : dayKey(dt);
}

function extractWarehouse(name) {
  if (!name) return "Unknown";
  const match = String(name).match(/\d+号仓/);
  return match ? match[0] : String(name);
}

function warehouseValue(row) {
  return (
    val(row, "warehouse_name") ??
    val(row, "仓库名称") ??
    val(row, "仓库") ??
    val(row, "仓库名") ??
    val(row, "库房名称") ??
    "Unknown"
  );
}

function mostCommon(values) {
  const counts = new Map();
  for (const value of values.filter((v) => v && v !== "Unknown" && v !== "未知仓库")) {
    counts.set(value, (counts.get(value) || 0) + 1);
  }
  return [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || "Unknown";
}

function groupBy(rows, keyFn) {
  const map = new Map();
  for (const row of rows) {
    const key = keyFn(row);
    if (key == null) continue;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  }
  return map;
}

function sum(rows, fn) {
  return rows.reduce((acc, row) => acc + num(fn(row), 0), 0);
}

function minDate(rows, fn) {
  const dates = rows.map((row) => parseDate(fn(row))).filter(Boolean).sort((a, b) => a - b);
  return dates[0] || null;
}

function maxDate(rows, fn) {
  const dates = rows.map((row) => parseDate(fn(row))).filter(Boolean).sort((a, b) => b - a);
  return dates[0] || null;
}

function classifyFile(file) {
  try {
    const c = cols(firstSheetRows(file));
    const joined = c.join("");
    if (c.includes("任务单开始时间") && c.includes("任务单结束时间")) return "task";
    if (c.includes("实际上班时间") && c.includes("实际下班时间")) return "attendance";
    if (c.includes("打包完成时间") && c.includes("件数")) return "volume";
    if (c.includes("打卡时间") && joined.includes("进出仓标识")) return "punch";
    return "unknown";
  } catch {
    return "unknown";
  }
}

function loadData(taskFile, attendanceFile) {
  const df1 = sheetRows(taskFile);
  const df2Raw = attendanceFile ? sheetRows(attendanceFile) : [];
  const df2 = df2Raw.map((r) => {
    const employee_no = String(colAt(r, 1) ?? "").trim();
    return {
      ...r,
      employee_no,
      班次名称: String(colAt(r, 7) ?? "").trim(),
      real_name: String(colAt(r, 3) ?? "").trim(),
      上班时间: parseDate(colAt(r, 13)),
      下班时间: parseDate(colAt(r, 18)),
      attendance_group: String(colAt(r, 26) ?? "未分组").trim() || "未分组"
    };
  });

  const nameMap = new Map(df2.map((r) => [r.employee_no, r.real_name]));
  const groupMap = new Map(df2.map((r) => [r.employee_no, r.attendance_group]));

  const df = df1.map((r) => {
    let start = parseDate(val(r, "任务单开始时间"));
    let end = parseDate(val(r, "任务单结束时间"));
    if (start && end && start.getHours() < 6) {
      start = addDays(start, 1);
      end = addDays(end, 1);
    }
    const employee_no = String(val(r, "employee_no") ?? "").trim() || null;
    const display_name = employee_no
      ? (nameMap.has(employee_no) ? `${nameMap.get(employee_no)} (${employee_no})` : employee_no)
      : String(val(r, "operator") ?? "Unknown");
    return {
      ...r,
      employee_no,
      operator: val(r, "operator") ?? "Unknown",
      warehouse_name: warehouseValue(r),
      start,
      end,
      display_name,
      type: "work",
      work_date: adjustWorkDate(start)
    };
  });
  return { df, df2, groupMap };
}

function buildIscOnlyTimeline(df) {
  const records = [];
  for (const [key, items] of groupBy(df.filter((r) => r.start && r.end && r.work_date), (r) => `${r.display_name}|${r.work_date}`)) {
    const [display, day] = key.split("|");
    const sorted = [...items].sort((a, b) => a.start - b.start);
    let cursor = sorted[0]?.start;
    const shift = cursor && cursor.getHours() < 15 ? "morning" : "evening";
    const groupName = "ISC Task Data";

    for (const task of sorted) {
      const s = parseDate(task.start);
      const e = parseDate(task.end);
      if (!s || !e || s >= e) continue;
      if (cursor && s > cursor) records.push(makeTimelineRow(display, day, cursor, s, "idle", shift, false, groupName));
      records.push(makeTimelineRow(display, day, s, e, "work", shift, false, groupName));
      if (!cursor || e > cursor) cursor = e;
    }
  }
  return records;
}

function loadVolumeData(file) {
  if (!file) return {};
  const rows = sheetRows(file);
  const cleaned = rows
    .map((r) => ({ date: parseDate(val(r, "打包完成时间")), cancel: val(r, "是否取消"), units: num(colAt(r, 6), 0) }))
    .filter((r) => r.date && !String(r.cancel ?? "").includes("是"));
  const out = {};
  for (const [d, items] of groupBy(cleaned, (r) => dayKey(r.date))) out[d] = sum(items, (r) => r.units);
  return out;
}

function loadBreakData(file) {
  if (!file) return [];
  const rows = sheetRows(file)
    .map((r) => ({
      emp_no: String(colAt(r, 3) ?? "").trim(),
      time: parseDate(colAt(r, 12)),
      action: String(colAt(r, 18) ?? ""),
      group: String(colAt(r, 2) ?? "").trim()
    }))
    .filter((r) => r.time)
    .sort((a, b) => a.emp_no.localeCompare(b.emp_no) || a.time - b.time);
  const breaks = [];
  for (const [emp, items] of groupBy(rows, (r) => r.emp_no)) {
    let outTime = null;
    for (const row of items.sort((a, b) => a.time - b.time)) {
      if (row.action.includes("出仓")) outTime = row.time;
      else if (row.action.includes("进仓") && outTime && row.time > outTime) {
        const lunchStart = new Date(outTime); lunchStart.setHours(11, 0, 0, 0);
        const lunchEnd = new Date(outTime); lunchEnd.setHours(14, 0, 0, 0);
        if (!(row.time < lunchStart || outTime > lunchEnd)) {
          breaks.push({ emp_no: emp, start: outTime, end: row.time, work_date: adjustWorkDate(outTime), type: "break", group: row.group });
        }
        outTime = null;
      }
    }
  }
  return breaks;
}

function loadIndirectData(file) {
  try {
    return sheetRows(file, "间接工时明细-间接工时")
      .map((r) => {
        const start = parseDate(val(r, "开始时间"));
        const end = parseDate(val(r, "结束时间"));
        return {
          employee_no: String(val(r, "employee_no") ?? "").trim(),
          操作分类: String(val(r, "操作分类") ?? "未分类").trim() || "未分类",
          开始时间: start,
          结束时间: end,
          "间接工时(分)": num(val(r, "间接工时(分)"), 0),
          work_date: adjustWorkDate(start),
          duration: hoursBetween(start, end)
        };
      })
      .filter((r) => r.开始时间 && r.结束时间 && r.结束时间 > r.开始时间);
  } catch {
    return [];
  }
}

function buildTimeline(df, df2, groupMap, breaks = []) {
  const records = [];
  const processed = new Set();

  function getIscTime(empNo, day) {
    const rows = df.filter((r) => String(r.employee_no) === String(empNo) && r.work_date === day);
    return { start: minDate(rows, (r) => r.上班时间), end: maxDate(rows, (r) => r.下班时间) };
  }

  for (const row of df2) {
    const empNo = String(row.employee_no);
    const display = `${row.real_name} (${empNo})`;
    const groupName = String(row.attendance_group || "Unknown");
    let workStart = parseDate(row.上班时间);
    let workEnd = parseDate(row.下班时间);
    let day = adjustWorkDate(workStart);
    const isc = getIscTime(empNo, day);
    if (!workStart && isc.start) workStart = isc.start;
    if (!workEnd && isc.end) workEnd = isc.end;
    if (!workStart || !workEnd) continue;
    day = dayKey(workStart);
    processed.add(`${empNo}|${day}`);
    const shiftName = String(row.班次名称 || "");
    const shift = shiftName.includes("早") || shiftName.includes("白") ? "morning" : (shiftName.includes("晚") ? "evening" : (workStart.getHours() < 15 ? "morning" : "evening"));
    let tasks = df
      .filter((r) => String(r.employee_no) === empNo && r.work_date === day && r.start && r.end && r.start < workEnd)
      .map((r) => ({ ...r, start: r.start < workStart ? workStart : r.start, type: "work", employee_no: empNo }));
    tasks.push(...breaks
      .filter((b) => String(b.emp_no).trim() === empNo && b.work_date === day && b.end > workStart && b.start < workEnd)
      .map((b) => ({ ...b, employee_no: empNo, start: b.start < workStart ? workStart : b.start, end: b.end > workEnd ? workEnd : b.end })));
    tasks = tasks.sort((a, b) => a.start - b.start);
    let cursor = workStart;
    for (const t of tasks) {
      const s = t.start > workStart ? t.start : workStart;
      const e = t.end < workEnd ? t.end : workEnd;
      if (s >= e) continue;
      if (s > cursor) records.push(makeTimelineRow(display, day, cursor, s, "idle", shift, false, groupName));
      records.push(makeTimelineRow(display, day, s, e, t.type, shift, false, groupName));
      if (e > cursor) cursor = e;
    }
    if (cursor < workEnd) records.push(makeTimelineRow(display, day, cursor, workEnd, "idle", shift, false, groupName));
  }

  const days = [...new Set(df.map((r) => r.work_date).filter(Boolean))].sort();
  for (const day of days) {
    for (const [empNo, items] of groupBy(df.filter((r) => r.work_date === day), (r) => String(r.employee_no))) {
      if (processed.has(`${empNo}|${day}`)) continue;
      const display = items[0].display_name;
      const groupName = groupMap.get(empNo) || "Unknown";
      let workStart = parseDate(minDate(items, (r) => r.上班时间) || minDate(items, (r) => r.start));
      let workEnd = parseDate(maxDate(items, (r) => r.下班时间) || maxDate(items, (r) => r.end));
      if (!workStart || !workEnd) continue;
      const shiftName = String(items[0].班次名称 || "");
      const shift = shiftName.includes("早") || shiftName.includes("白") ? "morning" : (shiftName.includes("晚") ? "evening" : (workStart.getHours() < 15 ? "morning" : "evening"));
      let tasks = items.filter((r) => r.start && r.end && r.start < workEnd).map((r) => ({ ...r, start: r.start < workStart ? workStart : r.start, type: "work" }));
      tasks.push(...breaks
        .filter((b) => String(b.emp_no).trim() === empNo && b.work_date === day && String(b.group).trim() === groupName && b.end > workStart && b.start < workEnd)
        .map((b) => ({ ...b, start: b.start < workStart ? workStart : b.start, end: b.end > workEnd ? workEnd : b.end })));
      tasks = tasks.sort((a, b) => a.start - b.start);
      let cursor = workStart;
      for (const t of tasks) {
        const s = t.start > workStart ? t.start : workStart;
        const e = t.end < workEnd ? t.end : workEnd;
        if (s >= e) continue;
        if (s > cursor) records.push(makeTimelineRow(display, day, cursor, s, "idle", shift, true, groupName));
        records.push(makeTimelineRow(display, day, s, e, t.end && dayKey(t.end) > dayKey(workEnd) ? "overnight" : t.type, shift, true, groupName));
        if (e > cursor) cursor = e;
      }
      if (cursor < workEnd) records.push(makeTimelineRow(display, day, cursor, workEnd, "idle", shift, true, groupName));
    }
  }
  return removeBreakOverlap(records);
}

function makeTimelineRow(name, date, start, end, type, shift, is_absent, group) {
  return { name, date, start, end, type, shift, is_absent, group, duration: hoursBetween(start, end) };
}

function removeBreakOverlap(timeline) {
  const result = [];
  for (const [, group] of groupBy(timeline, (r) => `${r.name}|${r.date}`)) {
    const breaks = group.filter((r) => r.type === "break");
    const others = group.filter((r) => r.type !== "break");
    for (const row of others) {
      let segments = [[row.start, row.end]];
      for (const br of breaks) {
        const next = [];
        for (const [s, e] of segments) {
          if (e <= br.start || s >= br.end) next.push([s, e]);
          else {
            if (s < br.start) next.push([s, br.start]);
            if (e > br.end) next.push([br.end, e]);
          }
        }
        segments = next;
      }
      for (const [s, e] of segments) if (s < e) result.push({ ...row, start: s, end: e, duration: hoursBetween(s, e) });
    }
    result.push(...breaks);
  }
  return result.sort((a, b) => a.name.localeCompare(b.name) || a.start - b.start || a.type.localeCompare(b.type));
}

function buildHistory(timeline) {
  const ratios = {};
  const wait = {};
  for (const [date, dayRows] of groupBy(timeline, (r) => r.date)) {
    ratios[date] = {};
    wait[date] = {};
    for (const [name, rows] of groupBy(dayRows, (r) => r.name)) {
      const workDuration = sum(rows.filter((r) => ["work", "overnight"].includes(r.type)), (r) => r.duration);
      const totalDuration = sum(rows, (r) => r.duration);
      ratios[date][name] = totalDuration > 0 ? (workDuration / totalDuration) * 100 : 0;
      if (rows.some((r) => r.is_absent)) continue;
      const work = rows.filter((r) => ["work", "overnight"].includes(r.type)).sort((a, b) => a.start - b.start);
      if (work.length) wait[date][empFromName(name)] = (work[0].start - minDate(rows, (r) => r.start)) / 1000;
    }
  }
  return { history_ratios: ratios, history_wait: wait };
}

function empFromName(name) {
  const m = String(name).match(/\(([^)]+)\)\s*$/);
  return m ? m[1].trim() : String(name).trim();
}

function buildSummary(dfDay, df, df2, currentDay) {
  const summary = {};
  for (const [name, rows] of groupBy(dfDay, (r) => r.name)) {
    const work = sum(rows.filter((r) => ["work", "overnight"].includes(r.type)), (r) => r.duration);
    const breaks = sum(rows.filter((r) => r.type === "break"), (r) => r.duration);
    const total = sum(rows, (r) => r.duration);
    const ratio = total - breaks > 0 ? (work / (total - breaks)) * 100 : 0;
    const emp = empFromName(name);
    const iams = (df2 || []).find((r) => String(r.employee_no) === emp && adjustWorkDate(r.上班时间) === currentDay);
    const iscRows = df.filter((r) => String(r.employee_no) === emp && r.work_date === currentDay);
    const inTime = iams?.上班时间 || minDate(iscRows, (r) => r.上班时间);
    const outTime = iams?.下班时间 || maxDate(iscRows, (r) => r.下班时间);
    summary[name] = { ratio, in: timeText(inTime), out: timeText(outTime) };
  }
  return summary;
}

function buildFirstActionTable(dfDay, df, df2, selectedDate, historyWait) {
  if (!df2?.length) return [];
  const rows = [];
  const prevDay = dayKey(addDays(parseDate(selectedDate), -1));
  for (const [name, group] of groupBy(dfDay, (r) => r.name)) {
    const emp = empFromName(name);
    const work = group.filter((r) => ["work", "overnight"].includes(r.type)).sort((a, b) => a.start - b.start);
    if (!work.length) continue;
    const iams = df2.find((r) => String(r.employee_no) === emp && adjustWorkDate(r.上班时间) === selectedDate);
    const iscRows = df.filter((r) => String(r.employee_no) === emp && r.work_date === selectedDate);
    const punchIn = iams?.上班时间 || minDate(iscRows, (r) => r.上班时间) || minDate(group, (r) => r.start);
    const firstTask = work[0].start;
    const waitSeconds = (firstTask - punchIn) / 1000;
    const prev = historyWait?.[prevDay]?.[emp];
    let trend = "-";
    if (prev != null) {
      const diff = waitSeconds - prev;
      const m = Math.floor(Math.abs(diff) / 60);
      const s = Math.floor(Math.abs(diff) % 60);
      trend = diff > 0 ? `↑ ${m}m ${s}s` : diff < 0 ? `↓ ${m}m ${s}s` : "=";
    }
    rows.push({ Name: name, Group: group[0].group, Shift: group[0].shift, "Punch-in": timeText(punchIn), "First Task": timeText(firstTask), Wait: `${Math.floor(waitSeconds / 60)}m ${Math.floor(waitSeconds % 60)}s`, Trend: trend, _wait_seconds: waitSeconds });
  }
  return rows.sort((a, b) => b._wait_seconds - a._wait_seconds).map(({ _wait_seconds, ...r }) => r);
}

function findNoOperationPeople(dfIsc, dfIams) {
  if (!dfIams?.length) return [];
  const iscSet = new Set(dfIsc.filter((r) => r.employee_no).map((r) => String(r.employee_no).trim()));
  return dfIams
    .filter((r) => r.上班时间 && r.下班时间 && !iscSet.has(String(r.employee_no).trim()))
    .map((r) => ({ employee_no: r.employee_no, real_name: r.real_name, 上班时间: r.上班时间, 下班时间: r.下班时间 }));
}

function serialize(row) {
  const out = {};
  for (const [k, v] of Object.entries(row)) out[k] = v instanceof Date ? v.toISOString() : v;
  return out;
}

function analyzeEfficiency(files) {
  let taskFile, attendanceFile, volumeFile, punchFile;
  for (const f of files) {
    const type = classifyFile(f);
    if (type === "task") taskFile = f;
    if (type === "attendance") attendanceFile = f;
    if (type === "volume") volumeFile = f;
    if (type === "punch") punchFile = f;
  }
  if (!taskFile) throw new Error("缺少任务单明细（df1）");
  const { df, df2, groupMap } = loadData(taskFile, attendanceFile);
  const volumeData = loadVolumeData(volumeFile);
  const breakDf = loadBreakData(punchFile);
  const indirectDf = loadIndirectData(taskFile);
  const timeline = attendanceFile ? buildTimeline(df, df2, groupMap, breakDf) : buildIscOnlyTimeline(df);
  const { history_ratios, history_wait } = buildHistory(timeline);
  const currentWarehouse = mostCommon(df.map((r) => extractWarehouse(r.warehouse_name)));
  const dates = [...new Set(timeline.map((r) => r.date))].sort();
  const completeness = {
    attendance: Boolean(attendanceFile),
    volume: Boolean(volumeFile),
    punch: Boolean(punchFile)
  };
  const days = {};
  for (const date of dates) {
    const dfDay = timeline.filter((r) => r.date === date);
    const summary = buildSummary(dfDay, df, df2, date);
    const totalWork = sum(dfDay.filter((r) => ["work", "overnight"].includes(r.type)), (r) => r.duration);
    const totalAttendance = sum(dfDay, (r) => r.duration);
    days[date] = {
      kpi: {
        warehouse: currentWarehouse,
        totalWorkers: new Set(dfDay.map((r) => r.name)).size,
        totalAttendance,
        totalWork,
        efficiencyRatio: totalAttendance > 0 ? (totalWork / totalAttendance) * 100 : 0,
        volume: volumeFile ? Math.round(volumeData[date] || 0) : null
      },
      summary,
      firstAction: buildFirstActionTable(dfDay, df, df2, date, history_wait),
      noOperation: findNoOperationPeople(df.filter((r) => r.work_date === date), df2.filter((r) => adjustWorkDate(r.上班时间) === date)),
      indirect: indirectDf.filter((r) => r.work_date === date).map(serialize)
    };
  }
  return {
    dates,
    currentWarehouse,
    timeline: timeline.map(serialize),
    days,
    history_ratios,
    completeness,
    groups: [...new Set(timeline.map((r) => r.group).filter(Boolean))].sort()
  };
}

function analyzeWeekly({ volumeFile, iscFile, pickFile }) {
  const daily = [];
  const volumeDates = [];
  if (volumeFile) {
    const volumeWb = workbook(volumeFile);
    const rows = validateRows({
      file: volumeFile,
      wb: volumeWb,
      label: "Upload Unit Data",
      requiredColumns: ["打包完成时间", "件数", "京东订单号", "状态"]
    });
    const filtered = rows
      .map((r) => ({ ...r, 打包完成时间: parseDate(val(r, "打包完成时间")), 件数: num(val(r, "件数"), 0) }))
      .filter((r) => r.打包完成时间 && val(r, "状态") === "交接完成" && val(r, "是否取消") !== "是");
    for (const [date, items] of groupBy(filtered, (r) => businessDate(r.打包完成时间))) {
      daily.push({ 业务日期: date, 单量: new Set(items.map((r) => val(r, "京东订单号"))).size, 件量: sum(items, (r) => r.件数) });
      volumeDates.push(date);
    }
  }
  let warehouseName = "";
  if (iscFile) {
    const iscWb = workbook(iscFile);
    const isc = validateRows({
      file: iscFile,
      wb: iscWb,
      sheetName: "日-考勤-日",
      label: "Labor Hours",
      requiredColumns: ["工作日", "考勤时长", "出勤人数"]
    });
    const workDaily = new Map();
    for (const [date, items] of groupBy(isc.map((r) => ({ ...r, 工作日: dayKey(parseDate(val(r, "工作日"))), 考勤时长: num(val(r, "考勤时长"), 0) })), (r) => r.工作日)) {
      workDaily.set(date, {
        总工时: sum(items, (r) => r.考勤时长),
        出勤人数: Math.max(...items.map((r) => num(val(r, "出勤人数"), 0)))
      });
    }
    for (const row of daily) {
      const attendance = workDaily.get(row.业务日期);
      if (attendance) Object.assign(row, attendance);
      else {
        row.总工时 = "No attendance data";
        row.出勤人数 = "No attendance data";
      }
      if (typeof row.总工时 === "number" && row.总工时 > 0) {
        row.UPPH = Number((row.件量 / row.总工时).toFixed(2));
      } else if (!attendance) {
        row.UPPH = "No attendance data";
      }
    }
    try {
      const whRows = sheetRowsFromWorkbook(iscWb, "明细&汇总-图表");
      const wh = whRows.map((r) => val(r, "仓库名称")).find(Boolean);
      warehouseName = extractWarehouse(wh) || "未知仓库";
    } catch {}
  }
  daily.sort((a, b) => String(a.业务日期).localeCompare(String(b.业务日期)));
  const targetUpph = iscFile ? (TARGET_MAP[warehouseName] || 11.3) : "";
  let personEfficiency = [];
  if (pickFile) personEfficiency = analyzePickEfficiency(pickFile, iscFile, targetUpph, volumeDates);
  return {
    kpi: { totalOrders: sum(daily, (r) => r.单量), totalUnits: sum(daily, (r) => r.件量), warehouseName, targetUpph },
    daily,
    personEfficiency
  };
}

function analyzePickEfficiency(pickFile, iscFile, targetUpph, fallbackDates = []) {
  const pickWb = workbook(pickFile);
  const pickRows = validateRows({
    file: pickFile,
    wb: pickWb,
    label: "Upload Picking Data",
    requiredColumns: ["拣货完成时间", "姓名"],
    columnGroups: [
      ["拣货开始时间", "任务领取时间"],
      ["实际拣货量", "拣货数量", "拣货件数", "件数"]
    ]
  });
  const pickStartColumn = firstPresentColumn(pickRows, ["拣货开始时间", "任务领取时间"]);
  const pickQtyColumn = firstPresentColumn(pickRows, ["实际拣货量", "拣货数量", "拣货件数", "件数"]);
  const pickEmployeeColumn = firstPresentColumn(pickRows, ["工号", "员工号"]);
  const pickTaskColumn = firstPresentColumn(pickRows, ["任务单号", "集合单号"]);
  const pick = pickRows
    .map((r, idx) => {
      const name = String(val(r, "姓名") ?? "").trim();
      const employeeNo = String(pickEmployeeColumn ? val(r, pickEmployeeColumn) ?? "" : "").trim();
      return {
        ...r,
        拣货完成时间: parseDate(val(r, "拣货完成时间")),
        拣货开始时间: parseDate(val(r, pickStartColumn)),
        实际拣货量: num(val(r, pickQtyColumn), 0),
        工号: employeeNo || name,
        姓名: name,
        _task_no: pickTaskColumn ? val(r, pickTaskColumn) : `row-${idx}`,
        _row_idx: idx
      };
    })
    .filter((r) => r.拣货完成时间);
  const fallbackDate = fallbackDates[0] || "Unknown Date";
  for (const r of pick) r.日期 = dayKey(r.拣货开始时间 || r.拣货完成时间) || fallbackDate;
  const rawMap = new Map();
  for (const [key, items] of groupBy(pick, (r) => `${r.日期}|${r.工号}`)) rawMap.set(key, sum(items, (r) => r.实际拣货量));
  const unique = [];
  const seen = new Set();
  for (const r of pick) {
    const location = val(r, "储位") ?? val(r, "库位") ?? r._row_idx;
    const k = `${location}|${r.拣货完成时间?.toISOString()}|${r._task_no}|${r.工号}`;
    if (!seen.has(k)) { seen.add(k); unique.push(r); }
  }
  const qtyMap = new Map();
  for (const [key, items] of groupBy(unique, (r) => `${r.日期}|${r.工号}`)) qtyMap.set(key, sum(items, (r) => r.实际拣货量));
  const timeMap = new Map();
  const taskEvents = [];
  for (const [taskKey, items] of groupBy(pick, (r) => `${r.日期}|${r.工号}|${r._task_no}`)) {
    const [日期, 工号] = taskKey.split("|");
    const starts = items.map((r) => parseDate(r.拣货开始时间)).filter(Boolean);
    const ends = items.map((r) => parseDate(r.拣货完成时间)).filter(Boolean);
    const end = ends.sort((a, b) => b - a)[0];
    if (!end) continue;
    const start = starts.sort((a, b) => a - b)[0] || end;
    taskEvents.push({ 日期, 工号, start, end });
  }
  for (const [key, items] of groupBy(taskEvents, (r) => `${r.日期}|${r.工号}`)) {
    const sorted = items.sort((a, b) => a.end - b.end);
    let hours = 0;
    for (let i = 0; i < sorted.length; i++) {
      const task = sorted[i];
      const taskSeconds = Math.min(1200, Math.max(0, (task.end - task.start) / 1000));
      const next = sorted[i + 1];
      const interval = next && next.start > task.end ? Math.min(1200, Math.max(0, (next.start - task.end) / 1000)) : 0;
      hours += (taskSeconds + interval) / 3600;
    }
    timeMap.set(key, hours);
  }
  const attMap = new Map();
  if (iscFile) {
    const iscWb = workbook(iscFile);
    const rawAttRows = validateRows({
      file: iscFile,
      wb: iscWb,
      sheetName: "明细&汇总-图表",
      label: "Labor Hours",
      requiredColumns: ["日", "员工号", "考勤时长"]
    });
    const attRows = rawAttRows.map((r) => ({ 日期: dayKey(parseDate(val(r, "日"))), 工号: String(val(r, "员工号") ?? "").trim(), 考勤时长: num(val(r, "考勤时长"), 0) }));
    for (const [key, items] of groupBy(attRows, (r) => `${r.日期}|${r.工号}`)) attMap.set(key, sum(items, (r) => r.考勤时长));
  }
  const nameMap = new Map(pick.map((r) => [r.工号, r.姓名]));
  return [...qtyMap.keys()].map((key) => {
    const [日期, 工号] = key.split("|");
    const 件数 = qtyMap.get(key) || 0;
    const 总件数 = rawMap.get(key) || 0;
    const hasAttendance = attMap.has(key);
    const 考勤时长 = hasAttendance ? attMap.get(key) || 0 : "";
    const rawEffectiveHours = timeMap.get(key) || 0;
    const effectiveCap = hasAttendance && 考勤时长 ? Math.min(考勤时长, 8) : rawEffectiveHours;
    const 有效工时 = Math.min(rawEffectiveHours, effectiveCap);
    const 拣非爆品效率 = 有效工时 ? 件数 / 有效工时 : 0;
    const 总效率 = 有效工时 ? 总件数 / 有效工时 : 0;
    return {
      日期,
      工号,
      姓名: nameMap.get(工号) || "",
      总件数,
      件数,
      有效工时,
      考勤时长,
      有效工时占比: hasAttendance && 考勤时长 ? 有效工时 / 考勤时长 : "",
      拣非爆品效率,
      总效率,
      低于目标: typeof targetUpph === "number" ? 拣非爆品效率 < targetUpph : ""
    };
  }).sort((a, b) => String(a.日期).localeCompare(String(b.日期)) || b.拣非爆品效率 - a.拣非爆品效率);
}

app.get("/api/health", (_req, res) => res.json({ ok: true }));

app.post("/api/efficiency/analyze", upload.array("files"), (req, res) => {
  try {
    res.json(analyzeEfficiency(req.files || []));
  } catch (error) {
    res.status(error.status || 400).json({ error: error.message, details: error.details || undefined });
  }
});

app.post("/api/weekly/analyze", upload.fields([{ name: "volume", maxCount: 1 }, { name: "isc", maxCount: 1 }, { name: "pick", maxCount: 1 }]), (req, res) => {
  try {
    const volumeFile = req.files?.volume?.[0];
    const iscFile = req.files?.isc?.[0];
    const pickFile = req.files?.pick?.[0];
    if (!volumeFile && !iscFile && !pickFile) throw new Error("请至少上传一个表格");
    res.json(analyzeWeekly({ volumeFile, iscFile, pickFile }));
  } catch (error) {
    res.status(error.status || 400).json({ error: error.message, details: error.details || undefined });
  }
});

app.post("/api/export", (req, res) => {
  const rows = req.body?.rows || [];
  const sheetName = req.body?.sheetName || "每日明细";
  const fileName = req.body?.fileName || "export.xlsx";
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), sheetName);
  const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.send(buffer);
});

app.listen(PORT, () => {
  console.log(`WorkSight API running on http://127.0.0.1:${PORT}`);
});
