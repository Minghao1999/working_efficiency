import {
  addDays,
  adjustWorkDate,
  classifyWorkbook,
  colAt,
  dayKey,
  extractWarehouse,
  groupBy,
  headerIndex,
  hoursBetween,
  maxDate,
  minDate,
  mostCommon,
  num,
  parseDate,
  sheetRows,
  sum,
  timeText,
  val,
  workbook,
  sheetRowsFromWorkbook,
  warehouseValue
} from "../utils/helpers.js";
import { summarizeCompletedVolume } from "./volumeService.js";

function rowsFor(file, wb, sheetName = null) {
  return wb ? sheetRowsFromWorkbook(wb, sheetName) : sheetRows(file, sheetName);
}

function loadAttendanceData(attendanceFile, wb = null) {
  const df2Raw = attendanceFile ? rowsFor(attendanceFile, wb) : [];
  return df2Raw.map((r) => {
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
}

function applyAttendanceNames(df, df2) {
  const nameMap = new Map(df2.map((r) => [r.employee_no, r.real_name]));
  return df.map((row) => ({
    ...row,
    display_name: row.employee_no
      ? (nameMap.has(row.employee_no) ? `${nameMap.get(row.employee_no)} (${row.employee_no})` : row.employee_no)
      : String(row.operator ?? "Unknown")
  }));
}

function loadTaskData(taskFile, df2 = [], wb = null) {
  const df1 = rowsFor(taskFile, wb);
  const nameMap = new Map(df2.map((r) => [r.employee_no, r.real_name]));
  return df1.map((r) => {
    let start = parseDate(val(r, "任务单开始时间"));
    let end = parseDate(val(r, "任务单结束时间"));
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
}

function loadData(taskFile, attendanceFile) {
  const df2 = loadAttendanceData(attendanceFile);
  const groupMap = new Map(df2.map((r) => [r.employee_no, r.attendance_group]));
  const df = loadTaskData(taskFile, df2);
  return { df, df2, groupMap };
}

function isWorkType(type) {
  return ["work", "overnight", "overtime"].includes(type);
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

function loadVolumeData(file, wb = null) {
  return summarizeCompletedVolume(file, wb).unitsByDate;
}

function loadBreakData(file, wb = null) {
  if (!file) return [];
  const rows = rowsFor(file, wb)
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

function parseIndirectRows(rows) {
  return rows
    .map((r) => {
      const start = parseDate(val(r, "开始时间"));
      const end = parseDate(val(r, "结束时间"));
      return {
        employee_no: String(val(r, "employee_no") ?? "").trim(),
        "操作分类": String(val(r, "操作分类") ?? "未分类").trim() || "未分类",
        "开始时间": start,
        "结束时间": end,
        "间接工时(分)": num(val(r, "间接工时(分)"), 0),
        work_date: adjustWorkDate(start),
        duration: hoursBetween(start, end)
      };
    })
    .filter((r) => r["开始时间"] && r["结束时间"] && r["结束时间"] > r["开始时间"]);
}

function loadIndirectData(file, wb = null) {
  try {
    if (wb) {
      const sheetName = wb.SheetNames.find((name) => String(name).includes("间接")) || "间接工时明细-间接工时";
      return parseIndirectRows(rowsFor(file, wb, sheetName));
    }
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
  const dfByDay = groupBy(df.filter((r) => r.work_date), (r) => r.work_date);
  const tasksByEmpDay = groupBy(df.filter((r) => r.work_date), (r) => `${String(r.employee_no)}|${r.work_date}`);
  const breaksByEmpDay = groupBy(breaks.filter((r) => r.work_date), (r) => `${String(r.emp_no).trim()}|${r.work_date}`);

  function clipToShift(item, workStart, workEnd) {
    const start = parseDate(item.start);
    const end = parseDate(item.end);
    if (!start || !end || end <= workStart || start >= workEnd) return null;
    const clippedStart = start < workStart ? workStart : start;
    const clippedEnd = end > workEnd ? workEnd : end;
    if (clippedStart >= clippedEnd) return null;
    return { ...item, start: clippedStart, end: clippedEnd, actualStart: start, actualEnd: end };
  }

  function getIscTime(empNo, day) {
    const rows = tasksByEmpDay.get(`${String(empNo)}|${day}`) || [];
    return { start: minDate(rows, (r) => r.start), end: maxDate(rows, (r) => r.end) };
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
    let tasks = (tasksByEmpDay.get(`${empNo}|${day}`) || [])
      .map((r) => clipToShift({ ...r, type: "work", employee_no: empNo }, workStart, workEnd))
      .filter(Boolean);
    tasks.push(...(breaksByEmpDay.get(`${empNo}|${day}`) || [])
      .map((b) => clipToShift({ ...b, employee_no: empNo }, workStart, workEnd))
      .filter(Boolean));
    tasks = tasks.sort((a, b) => a.start - b.start);
    let cursor = workStart;
    for (const t of tasks) {
      const s = t.start > workStart ? t.start : workStart;
      const e = t.end < workEnd ? t.end : workEnd;
      if (s >= e) continue;
      if (s > cursor) records.push(makeTimelineRow(display, day, cursor, s, "idle", shift, false, groupName));
      const type = t.type === "work" && t.actualEnd && t.actualEnd > workEnd ? "overtime" : t.type;
      records.push(makeTimelineRow(display, day, s, e, type, shift, false, groupName, { actualEnd: t.actualEnd }));
      if (e > cursor) cursor = e;
    }
    if (cursor < workEnd) records.push(makeTimelineRow(display, day, cursor, workEnd, "idle", shift, false, groupName));
  }

  const days = [...new Set(df.map((r) => r.work_date).filter(Boolean))].sort();
  for (const day of days) {
    for (const [empNo, items] of groupBy(dfByDay.get(day) || [], (r) => String(r.employee_no))) {
      if (processed.has(`${empNo}|${day}`)) continue;
      const display = items[0].display_name;
      const groupName = groupMap.get(empNo) || "Unknown";
      let workStart = parseDate(minDate(items, (r) => r.上班时间) || minDate(items, (r) => r.start));
      let workEnd = parseDate(maxDate(items, (r) => r.下班时间) || maxDate(items, (r) => r.end));
      if (!workStart || !workEnd) continue;
      const shiftName = String(items[0].班次名称 || "");
      const shift = shiftName.includes("早") || shiftName.includes("白") ? "morning" : (shiftName.includes("晚") ? "evening" : (workStart.getHours() < 15 ? "morning" : "evening"));
      let tasks = items
        .map((r) => clipToShift({ ...r, type: "work" }, workStart, workEnd))
        .filter(Boolean);
      tasks.push(...(breaksByEmpDay.get(`${empNo}|${day}`) || [])
        .filter((b) => String(b.group).trim() === groupName)
        .map((b) => clipToShift(b, workStart, workEnd))
        .filter(Boolean));
      tasks = tasks.sort((a, b) => a.start - b.start);
      let cursor = workStart;
      for (const t of tasks) {
        const s = t.start > workStart ? t.start : workStart;
        const e = t.end < workEnd ? t.end : workEnd;
        if (s >= e) continue;
        if (s > cursor) records.push(makeTimelineRow(display, day, cursor, s, "idle", shift, true, groupName));
        const type = t.type === "work" && t.actualEnd && t.actualEnd > workEnd ? "overtime" : t.type;
        records.push(makeTimelineRow(display, day, s, e, type, shift, true, groupName, { actualEnd: t.actualEnd }));
        if (e > cursor) cursor = e;
      }
      if (cursor < workEnd) records.push(makeTimelineRow(display, day, cursor, workEnd, "idle", shift, true, groupName));
    }
  }
  return removeBreakOverlap(records);
}

function makeTimelineRow(name, date, start, end, type, shift, is_absent, group, extra = {}) {
  return { name, date, start, end, type, shift, is_absent, group, ...extra, duration: hoursBetween(start, end) };
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
      const workDuration = sum(rows.filter((r) => isWorkType(r.type)), (r) => r.duration);
      const totalDuration = sum(rows, (r) => r.duration);
      ratios[date][name] = totalDuration > 0 ? (workDuration / totalDuration) * 100 : 0;
      if (rows.some((r) => r.is_absent)) continue;
      const work = rows.filter((r) => isWorkType(r.type)).sort((a, b) => a.start - b.start);
      if (work.length) wait[date][empFromName(name)] = (work[0].start - minDate(rows, (r) => r.start)) / 1000;
    }
  }
  return { history_ratios: ratios, history_wait: wait };
}

function empFromName(name) {
  const m = String(name).match(/\(([^)]+)\)\s*$/);
  return m ? m[1].trim() : String(name).trim();
}

function firstActionFromTaskEnd(df, emp, selectedDate) {
  const rows = (df || []).filter((r) => String(r.employee_no) === String(emp));
  const starts = rows.map((r) => parseDate(r.start)).filter((d) => d && dayKey(d) === selectedDate).sort((a, b) => a - b);
  const ends = rows.map((r) => parseDate(r.end)).filter((d) => d && dayKey(d) === selectedDate).sort((a, b) => a - b);
  const firstStart = starts[0];
  const firstEnd = ends[0];
  if (firstStart && firstEnd && firstEnd < firstStart) return firstEnd;
  return null;
}

function buildSummary(dfDay, df, df2, currentDay) {
  const summary = {};
  for (const [name, rows] of groupBy(dfDay, (r) => r.name)) {
    const work = sum(rows.filter((r) => isWorkType(r.type)), (r) => r.duration);
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
    const work = group.filter((r) => isWorkType(r.type)).sort((a, b) => a.start - b.start);
    if (!work.length) continue;
    const iams = df2.find((r) => String(r.employee_no) === emp && adjustWorkDate(r.上班时间) === selectedDate);
    const iscRows = df.filter((r) => String(r.employee_no) === emp && r.work_date === selectedDate);
    const punchIn = iams?.上班时间 || minDate(iscRows, (r) => r.上班时间) || minDate(group, (r) => r.start);
    const taskEndFirstAction = firstActionFromTaskEnd(df, emp, selectedDate);
    const firstTask = taskEndFirstAction || work[0].start;
    const waitSeconds = (firstTask - punchIn) / 1000;
    const prev = historyWait?.[prevDay]?.[emp];
    let trend = "-";
    if (prev != null) {
      const diff = waitSeconds - prev;
      const m = Math.floor(Math.abs(diff) / 60);
      const s = Math.floor(Math.abs(diff) % 60);
      trend = diff > 0 ? `↑ ${m}m ${s}s` : diff < 0 ? `↓ ${m}m ${s}s` : "=";
    }
    rows.push({ Name: name, Group: group[0].group, Shift: group[0].shift, "Punch-in": timeText(punchIn), "First Task": timeText(firstTask), Wait: `${Math.floor(waitSeconds / 60)}m ${Math.floor(waitSeconds % 60)}s`, Trend: trend, _first_task_from_end: Boolean(taskEndFirstAction), _wait_seconds: waitSeconds });
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

const efficiencySessions = new Map();

function getEfficiencySession(sessionId) {
  const key = sessionId || "default";
  if (!efficiencySessions.has(key)) efficiencySessions.set(key, {});
  return efficiencySessions.get(key);
}

function pruneEfficiencySession(cache, activeIds = []) {
  if (!activeIds.length) return;
  for (const key of ["task", "attendance", "volume", "punch"]) {
    if (cache[key]?.identity && !activeIds.includes(cache[key].identity)) delete cache[key];
  }
}

function efficiencyTimelineKey(cache) {
  return [
    cache.task?.identity || "",
    cache.attendance?.identity || "",
    cache.punch?.identity || ""
  ].join("::");
}

function buildEfficiencyResponse(cache, { delta = false } = {}) {
  const volumeData = cache.volume?.volumeData || {};
  const base = cache.analysis;
  const completeness = {
    attendance: Boolean(cache.attendance),
    volume: Boolean(cache.volume),
    punch: Boolean(cache.punch)
  };
  const days = {};

  for (const date of base.dates) {
    const baseDay = base.days[date];
    days[date] = {
      ...baseDay,
      kpi: {
        ...baseDay.kpi,
        volume: cache.volume ? Math.round(volumeData[date] || 0) : null
      }
    };
  }

  if (delta) {
    return {
      partial: "volume",
      days: Object.fromEntries(Object.entries(days).map(([date, day]) => [date, { kpi: { volume: day.kpi.volume } }])),
      completeness
    };
  }

  return {
    dates: base.dates,
    currentWarehouse: base.currentWarehouse,
    timeline: base.timeline,
    days,
    history_ratios: base.history_ratios,
    completeness,
    groups: base.groups
  };
}

export function analyzeEfficiency(files, { sessionId = "default", activeFiles = [], fileKeys = [], partial = "" } = {}) {
  const cache = getEfficiencySession(sessionId);
  pruneEfficiencySession(cache, activeFiles);

  files.forEach((f, index) => {
    const identity = fileKeys[index] || `${f.originalname}|${f.size || 0}`;
    const wb = workbook(f);
    const type = classifyWorkbook(wb);
    if (type === "task" && cache.task?.identity !== identity) {
      const df = loadTaskData(f, cache.attendance?.df2 || [], wb);
      cache.task = { identity, df, indirectDf: loadIndirectData(f, wb) };
    }
    if (type === "attendance" && cache.attendance?.identity !== identity) {
      const df2 = loadAttendanceData(f, wb);
      cache.attendance = {
        identity,
        df2,
        groupMap: new Map(df2.map((r) => [r.employee_no, r.attendance_group]))
      };
    }
    if (type === "volume" && cache.volume?.identity !== identity) {
      cache.volume = { identity, volumeData: loadVolumeData(f, wb) };
    }
    if (type === "punch" && cache.punch?.identity !== identity) {
      cache.punch = { identity, breakDf: loadBreakData(f, wb) };
    }
  });

  if (!cache.task) throw new Error("缺少任务单明细（df1）");

  const df2 = cache.attendance?.df2 || [];
  const groupMap = cache.attendance?.groupMap || new Map();
  const df = cache.attendance ? applyAttendanceNames(cache.task.df, df2) : cache.task.df;
  const volumeData = cache.volume?.volumeData || {};
  const breakDf = cache.punch?.breakDf || [];
  const indirectDf = cache.task.indirectDf || [];
  const timelineKey = efficiencyTimelineKey(cache);
  if (cache.analysis?.timelineKey === timelineKey) {
    return buildEfficiencyResponse(cache, { delta: partial === "volume" });
  }

  const timeline = cache.attendance ? buildTimeline(df, df2, groupMap, breakDf) : buildIscOnlyTimeline(df);
  const { history_ratios, history_wait } = buildHistory(timeline);
  const currentWarehouse = mostCommon(df.map((r) => extractWarehouse(r.warehouse_name)));
  const dates = [...new Set(timeline.map((r) => r.date))].sort();
  const timelineByDate = groupBy(timeline, (r) => r.date);
  const dfByDate = groupBy(df.filter((r) => r.work_date), (r) => r.work_date);
  const attendanceByDate = groupBy(df2.filter((r) => adjustWorkDate(colAt(r, 13))), (r) => adjustWorkDate(colAt(r, 13)));
  const indirectByDate = groupBy(indirectDf.filter((r) => r.work_date), (r) => r.work_date);
  const completeness = {
    attendance: Boolean(cache.attendance),
    volume: Boolean(cache.volume),
    punch: Boolean(cache.punch)
  };
  const days = {};
  for (const date of dates) {
    const dfDay = timelineByDate.get(date) || [];
    const taskDay = dfByDate.get(date) || [];
    const attendanceDay = attendanceByDate.get(date) || [];
    const summary = buildSummary(dfDay, df, df2, date);
    const totalWork = sum(dfDay.filter((r) => isWorkType(r.type)), (r) => r.duration);
    const totalAttendance = sum(dfDay, (r) => r.duration);
    days[date] = {
      kpi: {
        warehouse: currentWarehouse,
        totalWorkers: new Set(dfDay.map((r) => r.name)).size,
        totalAttendance,
        totalWork,
        efficiencyRatio: totalAttendance > 0 ? (totalWork / totalAttendance) * 100 : 0,
        volume: cache.volume ? Math.round(volumeData[date] || 0) : null
      },
      summary,
      firstAction: buildFirstActionTable(dfDay, df, df2, date, history_wait),
      noOperation: findNoOperationPeople(df.filter((r) => r.work_date === date), df2.filter((r) => adjustWorkDate(r.上班时间) === date)),
      indirect: indirectDf.filter((r) => r.work_date === date).map(serialize)
    };
  }
  cache.analysis = {
    timelineKey,
    dates,
    currentWarehouse,
    timeline: timeline.map(serialize),
    days,
    history_ratios,
    completeness,
    groups: [...new Set(timeline.map((r) => r.group).filter(Boolean))].sort()
  };
  return buildEfficiencyResponse(cache);
}
