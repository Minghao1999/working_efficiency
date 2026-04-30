import XLSX from "xlsx";

export const TARGET_MAP = { "1号仓": 34.57, "2号仓": 37.11, "5号仓": 11.3 };

export function sheetRows(file, sheetName = null) {
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

export function workbook(file) {
  return XLSX.read(file.buffer, { type: "buffer", cellDates: true });
}

export function sheetRowsFromWorkbook(wb, sheetName = null) {
  const name = sheetName || wb.SheetNames[0];
  const ws = wb.Sheets[name];
  if (!ws) throw new Error(`找不到 sheet：${name}`);
  return XLSX.utils.sheet_to_json(ws, { defval: null, raw: false }).map((row) => {
    const out = {};
    for (const [key, value] of Object.entries(row)) out[String(key).trim()] = value;
    return out;
  });
}

export function firstSheetRows(file, n = 5) {
  return sheetRows(file).slice(0, n);
}

export function cols(rows) {
  return rows[0] ? Object.keys(rows[0]).map((c) => String(c).trim()) : [];
}

export function val(row, key) {
  return row?.[key] ?? null;
}

export function colAt(row, idx) {
  const keys = Object.keys(row || {});
  return row?.[keys[idx]] ?? null;
}

export function headerIndex(headers, names) {
  const normalized = headers.map((header) => String(header ?? "").trim());
  for (const name of names) {
    const index = normalized.indexOf(name);
    if (index >= 0) return index;
  }
  return -1;
}

export function firstPresentColumn(rows, candidates) {
  const available = new Set(cols(rows));
  return candidates.find((column) => available.has(column)) || null;
}

export function uploadValidationError(detail) {
  const error = new Error("上传表格校验失败");
  error.status = 400;
  error.details = [detail];
  return error;
}

export function validateRows({ file, wb, sheetName = null, label, requiredColumns = [], columnGroups = [] }) {
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

export function parseDate(v) {
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

export function num(v, fallback = 0) {
  if (v == null || v === "") return fallback;
  const n = Number(String(v).replace(/,/g, "").replace("%", ""));
  return Number.isFinite(n) ? n : fallback;
}

export function dayKey(d) {
  d = parseDate(d);
  if (!d) return null;
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

export function timeText(d) {
  d = parseDate(d);
  if (!d) return "缺失";
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}:${String(d.getSeconds()).padStart(2, "0")}`;
}

export function addDays(d, days) {
  d = parseDate(d);
  if (!d) return null;
  const x = new Date(d);
  x.setDate(x.getDate() + days);
  return x;
}

export function hoursBetween(a, b) {
  a = parseDate(a);
  b = parseDate(b);
  if (!a || !b || b <= a) return 0;
  return (b - a) / 3600000;
}

export function adjustWorkDate(dt) {
  dt = parseDate(dt);
  if (!dt) return null;
  return dt.getHours() < 6 ? dayKey(addDays(dt, -1)) : dayKey(dt);
}

export function businessDate(dt) {
  dt = parseDate(dt);
  if (!dt) return null;
  return dt.getHours() < 5 ? dayKey(addDays(dt, -1)) : dayKey(dt);
}

export function extractWarehouse(name) {
  if (!name) return "Unknown";
  const match = String(name).match(/\d+号仓/);
  return match ? match[0] : String(name);
}

export function warehouseValue(row) {
  return (
    val(row, "warehouse_name") ??
    val(row, "仓库名称") ??
    val(row, "仓库") ??
    val(row, "仓库名") ??
    val(row, "库房名称") ??
    "Unknown"
  );
}

export function mostCommon(values) {
  const counts = new Map();
  for (const value of values.filter((v) => v && v !== "Unknown" && v !== "未知仓库")) {
    counts.set(value, (counts.get(value) || 0) + 1);
  }
  return [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || "Unknown";
}

export function groupBy(rows, keyFn) {
  const map = new Map();
  for (const row of rows) {
    const key = keyFn(row);
    if (key == null) continue;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  }
  return map;
}

export function sum(rows, fn) {
  return rows.reduce((acc, row) => acc + num(fn(row), 0), 0);
}

export function minDate(rows, fn) {
  const dates = rows.map((row) => parseDate(fn(row))).filter(Boolean).sort((a, b) => a - b);
  return dates[0] || null;
}

export function maxDate(rows, fn) {
  const dates = rows.map((row) => parseDate(fn(row))).filter(Boolean).sort((a, b) => b - a);
  return dates[0] || null;
}

export function classifyFile(file) {
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
