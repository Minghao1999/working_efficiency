import XLSX from "xlsx";
import { dayKey, headerIndex, num, parseDate } from "../utils/helpers.js";

export function summarizeCompletedVolume(file) {
  if (!file) return { unitsByDate: {}, ordersByDate: {} };
  const wb = XLSX.read(file.buffer, { type: "buffer", cellDates: true, dense: true, sheets: 0 });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) return { unitsByDate: {}, ordersByDate: {} };

  const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
  const cellValue = (rowIndex, colIndex) => {
    const cell = Array.isArray(ws) ? ws[rowIndex]?.[colIndex] : ws[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })];
    return cell && typeof cell === "object" && "v" in cell ? cell.v : cell;
  };
  const rowValues = (rowIndex) => {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) row.push(cellValue(rowIndex, c));
    return row;
  };

  let headerRow = null;
  let headerRowIndex = -1;
  for (let r = range.s.r; r <= range.e.r; r++) {
    const values = rowValues(r);
    if (values.some((cell) => String(cell ?? "").trim() === "打包完成时间")) {
      headerRow = values;
      headerRowIndex = r;
      break;
    }
  }
  if (!headerRow) return { unitsByDate: {}, ordersByDate: {} };

  const dateIdx = headerIndex(headerRow, ["打包完成时间"]);
  const unitIdx = headerIndex(headerRow, ["件数", "实际件数"]);
  const cancelIdx = headerIndex(headerRow, ["是否取消"]);
  const statusIdx = headerIndex(headerRow, ["状态"]);
  const orderIdx = headerIndex(headerRow, ["京东订单号"]);
  if (dateIdx < 0 || unitIdx < 0 || statusIdx < 0) return { unitsByDate: {}, ordersByDate: {} };

  const dateCol = range.s.c + dateIdx;
  const unitCol = range.s.c + unitIdx;
  const cancelCol = cancelIdx >= 0 ? range.s.c + cancelIdx : -1;
  const statusCol = range.s.c + statusIdx;
  const orderCol = orderIdx >= 0 ? range.s.c + orderIdx : -1;
  const unitsByDate = {};
  const orderSets = {};

  for (let i = headerRowIndex + 1; i <= range.e.r; i++) {
    if (String(cellValue(i, statusCol) ?? "").trim() !== "交接完成") continue;
    if (cancelCol >= 0 && String(cellValue(i, cancelCol) ?? "").trim() === "是") continue;
    const date = parseDate(cellValue(i, dateCol));
    if (!date) continue;
    const key = dayKey(date);
    unitsByDate[key] = (unitsByDate[key] || 0) + num(cellValue(i, unitCol), 0);
    if (orderCol >= 0) {
      const order = String(cellValue(i, orderCol) ?? "").trim();
      if (order) {
        if (!orderSets[key]) orderSets[key] = new Set();
        orderSets[key].add(order);
      }
    }
  }

  const ordersByDate = {};
  for (const [date, orders] of Object.entries(orderSets)) ordersByDate[date] = orders.size;
  return { unitsByDate, ordersByDate };
}
