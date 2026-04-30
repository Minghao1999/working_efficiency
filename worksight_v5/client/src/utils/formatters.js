import { COLUMN_LABELS } from "../constants";

export function sortRows(rows, sort) {
  return [...rows].sort((a, b) => {
    if (!sort) {
      return Number(a._deleted) - Number(b._deleted) || Number(b.拣非爆品效率 || 0) - Number(a.拣非爆品效率 || 0);
    }

    const result = compareValues(a[sort.key], b[sort.key]);
    return sort.direction === "asc" ? result : -result;
  });
}

export function compareValues(a, b) {
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

export function personDeleteKey(row) {
  return row?.姓名 || row?.工号 || "";
}

export function formatCell(v) {
  if (v == null) return "";
  if (typeof v === "number") return Number.isInteger(v) ? v.toLocaleString() : v.toFixed(2);
  if (typeof v === "boolean") return v ? "Yes" : "No";
  if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}T/.test(v)) return v.replace("T", " ").slice(0, 19);
  return String(v);
}

export function columnLabel(key) {
  return COLUMN_LABELS[key] || key;
}

export function translateError(message) {
  return String(message || "")
    .replace("缺少任务单明细（df1）", "Missing task detail file")
    .replace("缺少班次数据（df2）", "Missing attendance shift file")
    .replace("缺少件量数据（df3）", "Missing volume file")
    .replace("缺少打卡数据（df4）", "Missing punch log file")
    .replace("请上传销售单综合数据", "Please upload the sales order file")
    .replace("缺少字段", "Missing fields")
    .replace("找不到 sheet", "Sheet not found");
}

export function formatUploadError(error) {
  if (error?.details?.length) {
    return error.details.map((item) => {
      const file = item.file ? ` (${item.file})` : "";
      const sheet = item.sheet ? `Sheet: ${item.sheet}` : "";
      const missingSheets = item.missingSheets?.length ? `Missing sheets: ${item.missingSheets.join(", ")}` : "";
      const missingColumns = item.missingColumns?.length ? `Missing columns: ${item.missingColumns.join(", ")}` : "";
      const missingColumnGroups = item.missingColumnGroups?.length ? `Missing one of: ${item.missingColumnGroups.map((group) => group.join(" / ")).join("; ")}` : "";
      return [item.label + file, sheet, missingSheets, missingColumns, missingColumnGroups].filter(Boolean).join(" - ");
    }).join("\n");
  }
  return translateError(error?.error || error?.message || "Analysis failed");
}

export function timeOf(v) {
  const d = new Date(v);
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}:${String(d.getSeconds()).padStart(2, "0")}`;
}

export function previousDate(date) {
  const d = new Date(`${date}T00:00:00`);
  d.setDate(d.getDate() - 1);
  return d.toISOString().slice(0, 10);
}

export function palette(i) {
  return ["#2563EB", "#16A34A", "#F59E0B", "#DC2626", "#7C3AED", "#0891B2", "#DB2777"][i % 7];
}
