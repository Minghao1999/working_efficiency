import {
  TARGET_MAP,
  dayKey,
  extractWarehouse,
  firstPresentColumn,
  groupBy,
  num,
  parseDate,
  sheetRowsFromWorkbook,
  sum,
  val,
  validateRows,
  workbook
} from "../utils/helpers.js";
import { summarizeCompletedVolume } from "./volumeService.js";

export function analyzeWeekly({ volumeFile, iscFile, pickFile, existingDaily = [] }) {
  const daily = [];
  const volumeDates = [];
  if (volumeFile) {
    const { unitsByDate, ordersByDate } = summarizeCompletedVolume(volumeFile);
    for (const [date, units] of Object.entries(unitsByDate)) {
      daily.push({ 业务日期: date, 单量: ordersByDate[date] || 0, 件量: units });
      volumeDates.push(date);
    }
  } else if (Array.isArray(existingDaily)) {
    for (const row of existingDaily) {
      const date = row?.业务日期;
      if (!date) continue;
      daily.push({
        业务日期: date,
        单量: num(row.单量, 0),
        件量: num(row.件量, 0)
      });
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
  let pickingGantt = [];
  if (pickFile) {
    const pickAnalysis = analyzePickEfficiency(pickFile, iscFile, targetUpph, volumeDates);
    personEfficiency = pickAnalysis.personEfficiency;
    pickingGantt = pickAnalysis.pickingGantt;
  }
  return {
    kpi: { totalOrders: sum(daily, (r) => r.单量), totalUnits: sum(daily, (r) => r.件量), warehouseName, targetUpph },
    daily,
    personEfficiency,
    pickingGantt
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
  return analyzePickRows(pickRows, iscFile, targetUpph, fallbackDates);
}

export function analyzePickRows(pickRows, iscFile = null, targetUpph = "", fallbackDates = []) {
  const pickStartColumn = firstPresentColumn(pickRows, ["拣货开始时间", "任务领取时间"]);
  const pickReceiveColumn = firstPresentColumn(pickRows, ["任务领取时间", "拣货开始时间"]);
  const pickQtyColumn = firstPresentColumn(pickRows, ["实际拣货量", "拣货数量", "拣货件数", "件数"]);
  const pickEmployeeColumn = firstPresentColumn(pickRows, ["工号", "员工号"]);
  const pickTaskColumn = firstPresentColumn(pickRows, ["集合单号", "任务单号"]);
  const pick = pickRows
    .map((r, idx) => {
      const name = String(val(r, "姓名") ?? "").trim();
      const employeeNo = String(pickEmployeeColumn ? val(r, pickEmployeeColumn) ?? "" : "").trim();
      return {
        ...r,
        拣货完成时间: parseDate(val(r, "拣货完成时间")),
        拣货开始时间: parseDate(val(r, pickStartColumn)),
        _receive_time: parseDate(val(r, pickReceiveColumn)),
        实际拣货量: num(val(r, pickQtyColumn), 0),
        工号: employeeNo || name,
        姓名: name,
        _task_no: pickTaskColumn ? val(r, pickTaskColumn) : `row-${idx}`,
        _row_idx: idx
      };
    })
    .filter((r) => r.拣货完成时间);
  const fallbackDate = fallbackDates[0] || "Unknown Date";
  for (const r of pick) r.日期 = dayKey(r.拣货完成时间) || fallbackDate;
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
  const totalSpanMap = new Map();
  const taskEvents = [];
  for (const [taskKey, items] of groupBy(pick, (r) => `${r.日期}|${r.工号}|${r._task_no}`)) {
    const [日期, 工号] = taskKey.split("|");
    const starts = items.map((r) => parseDate(r.拣货开始时间)).filter(Boolean);
    const ends = items.map((r) => parseDate(r.拣货完成时间)).filter(Boolean);
    const end = ends.sort((a, b) => b - a)[0];
    if (!end) continue;
    const start = clampPickingStartToCompletionDate(starts.sort((a, b) => a - b)[0] || end, end);
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
    const firstStart = sorted.map((task) => task.start).filter(Boolean).sort((a, b) => a - b)[0];
    const lastEnd = sorted.map((task) => task.end).filter(Boolean).sort((a, b) => b - a)[0];
    totalSpanMap.set(key, totalDurationHours(firstStart, lastEnd));
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
  const personEfficiency = [...qtyMap.keys()].map((key) => {
    const [日期, 工号] = key.split("|");
    const 件数 = qtyMap.get(key) || 0;
    const 总件数 = rawMap.get(key) || 0;
    const hasAttendance = attMap.has(key);
    const 考勤时长 = hasAttendance ? attMap.get(key) || 0 : "";
    const rawEffectiveHours = timeMap.get(key) || 0;
    const effectiveCap = hasAttendance && 考勤时长 ? Math.min(考勤时长, 8) : rawEffectiveHours;
    const 有效工时 = Math.min(rawEffectiveHours, effectiveCap);
    const 拣非爆品效率 = 有效工时 ? 件数 / 有效工时 : 0;
    const totalSpanHours = totalSpanMap.get(key) || 0;
    const 总效率 = totalSpanHours ? 总件数 / totalSpanHours : 0;
    return {
      日期,
      工号,
      姓名: nameMap.get(工号) || "",
      总件数,
      件数,
      有效工时,
      总时长: totalSpanHours,
      考勤时长,
      有效工时占比: hasAttendance && 考勤时长 ? 有效工时 / 考勤时长 : "",
      拣非爆品效率,
      总效率,
      低于目标: typeof targetUpph === "number" ? 拣非爆品效率 < targetUpph : ""
    };
  }).sort((a, b) => String(a.日期).localeCompare(String(b.日期)) || b.拣非爆品效率 - a.拣非爆品效率);
  return { personEfficiency, pickingGantt: buildPickingGantt(pick) };
}

function clampPickingStartToCompletionDate(start, end) {
  if (!start || !end) return start;
  const completionDayStart = new Date(end);
  completionDayStart.setHours(0, 0, 0, 0);
  return start < completionDayStart ? completionDayStart : start;
}

function lunchWindowFor(date) {
  const start = new Date(date);
  start.setHours(12, 0, 0, 0);
  const end = new Date(date);
  end.setHours(13, 30, 0, 0);
  return { start, end };
}

function lunchBreakForRange(start, end) {
  if (!start || !end || end <= start) return null;
  const lunch = lunchWindowFor(start);
  const validStart = new Date(Math.max(start.getTime(), lunch.start.getTime()));
  const validEnd = new Date(Math.min(end.getTime(), lunch.end.getTime()));
  if ((validEnd - validStart) / 1000 < 1800) return null;
  return { start: validStart, end: new Date(validStart.getTime() + 30 * 60 * 1000) };
}

function totalDurationHours(start, end) {
  if (!start || !end || end <= start) return 0;
  const lunch = lunchBreakForRange(start, end);
  return Math.max(0, (end - start) / 3600000 - (lunch ? 0.5 : 0));
}

function buildPickingGantt(pick) {
  const rows = [];
  for (const [personKey, personRows] of groupBy(pick, (r) => `${r.日期}|${r.工号}`)) {
    const [date, employeeNo] = personKey.split("|");
    const name = personRows.find((r) => r.姓名)?.姓名 || employeeNo;
    const taskGroups = [];

    for (const [, group] of groupBy(personRows, (r) => r._task_no || r._row_idx)) {
      const sorted = [...group].sort((a, b) => (a.拣货完成时间 || 0) - (b.拣货完成时间 || 0));
      const rawReceive = sorted.map((r) => r._receive_time || r.拣货开始时间).filter(Boolean).sort((a, b) => a - b)[0];
      const completions = sorted.map((r) => r.拣货完成时间).filter(Boolean);
      if (!rawReceive || !completions.length) continue;

      let startPick;
      let endPick;
      let isBurst = false;

      if (sorted.length === 1) {
        endPick = completions[0];
        startPick = new Date(Math.max(rawReceive.getTime(), endPick.getTime() - 2 * 60 * 1000));
      } else {
        const counts = new Map();
        for (const end of completions) {
          const key = end.getTime();
          counts.set(key, (counts.get(key) || 0) + 1);
        }
        const [mostCommonEndMs, maxCount] = [...counts.entries()].sort((a, b) => b[1] - a[1])[0];
        if (maxCount / sorted.length >= 0.8) {
          endPick = new Date(Number(mostCommonEndMs));
          startPick = new Date(endPick.getTime() - sorted.length * 60 * 1000);
          isBurst = true;
        } else {
          startPick = completions.sort((a, b) => a - b)[0];
          endPick = completions.sort((a, b) => b - a)[0];
        }
      }

      if (startPick && endPick && endPick > startPick) {
        const receive = clampPickingStartToCompletionDate(rawReceive, endPick);
        startPick = clampPickingStartToCompletionDate(startPick, endPick);
        taskGroups.push({ receive, startPick, endPick, isBurst });
      }
    }

    const sortedTasks = taskGroups.sort((a, b) => a.startPick - b.startPick);
    let lunchUsed = false;

    for (let i = 0; i < sortedTasks.length; i++) {
      const task = sortedTasks[i];
      const start = i === 0 && task.receive < task.startPick ? task.receive : task.startPick;
      if (addPickingSegmentWithLunch(rows, { date, employeeNo, name, start, end: task.endPick, type: task.isBurst ? "burst" : "work" }, lunchUsed)) {
        lunchUsed = true;
      }

      const nextTask = sortedTasks[i + 1];
      if (!nextTask || nextTask.startPick <= task.endPick) continue;

      const idleStart = task.endPick;
      const idleEnd = nextTask.startPick;
      if (addPickingSegmentWithLunch(rows, { date, employeeNo, name, start: idleStart, end: idleEnd, type: "idle" }, lunchUsed)) {
        lunchUsed = true;
      }
    }
  }

  return rows.sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.name).localeCompare(String(b.name)) || new Date(a.start) - new Date(b.start));
}

function addPickingSegmentWithLunch(rows, segment, lunchUsed) {
  if (lunchUsed || !["work", "idle"].includes(segment.type)) {
    addPickingGanttRow(rows, segment);
    return false;
  }

  const lunch = lunchBreakForRange(segment.start, segment.end);
  if (!lunch) {
    addPickingGanttRow(rows, segment);
    return false;
  }

  if (lunch.start > segment.start) addPickingGanttRow(rows, { ...segment, end: lunch.start });
  addPickingGanttRow(rows, { ...segment, start: lunch.start, end: lunch.end, type: "lunch" });
  if (segment.end > lunch.end) addPickingGanttRow(rows, { ...segment, start: lunch.end });
  return true;
}

function addPickingGanttRow(rows, { date, employeeNo, name, start, end, type }) {
  if (!start || !end || end <= start) return;
  rows.push({
    date,
    employeeNo,
    name,
    start: start.toISOString(),
    end: end.toISOString(),
    type,
    duration: (end - start) / 3600000
  });
}
