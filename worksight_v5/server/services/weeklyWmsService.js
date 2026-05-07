import { dayKey, dayKeyFromText, num, parseDate, sum } from "../utils/helpers.js";
import { getDb } from "./mongo.js";
import { analyzePickRows } from "./weeklyService.js";

const WMS_URL = "https://iwms.us.jdlglobal.com/reportApi/services/smartQueryWS?wsdl";

const DEFAULT_WMS_USER = "minghao.sun@jd.com";
const DEFAULT_WMS_WK_NO = "jdhk_ulXbWlUMYeET";
const DEFAULT_WAREHOUSE_NO = "C0000009943";
const WAREHOUSE_NO_BY_KEY = {
  "1": "C0000000389",
  "2": "C0000002427",
  "5": "C0000009943"
};
const TARGET_UPPH_BY_WAREHOUSE = {
  "1": 34.57,
  "2": 37.11,
  "5": 11.3
};

function pad(value) {
  return String(value).padStart(2, "0");
}

function normalizeDateRange(from, to) {
  const start = parseDate(from);
  const end = parseDate(to);
  if (!start || !end) {
    const error = new Error("Please select both From and To dates.");
    error.status = 400;
    throw error;
  }
  if (end < start) {
    const error = new Error("To date must be after From date.");
    error.status = 400;
    throw error;
  }
  const formatDate = (date) => `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
  return {
    startDate: formatDate(start),
    endDate: formatDate(end),
    startTime: `${formatDate(start)} 00:00:00`,
    endTime: `${formatDate(end)} 23:59:59`
  };
}

function dateKeysBetween(startDate, endDate) {
  const dates = [];
  const current = parseDate(startDate);
  const end = parseDate(endDate);
  if (!current || !end) return dates;
  current.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);
  while (current <= end) {
    dates.push(dayKey(current));
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

function todayKey() {
  return dayKey(new Date());
}

function cacheStatusForDate(date) {
  if (date < todayKey()) return "final";
  if (date === todayKey()) return "provisional";
  return "skip";
}

function buildSoapEnvelope(arg0, arg1) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <soap:Body>
    <wms3:queryWs xmlns:wms3="http://wms3.360buy.com">
      <arg0>${JSON.stringify(arg0)}</arg0>
      <arg1>${JSON.stringify(arg1)}</arg1>
    </wms3:queryWs>
  </soap:Body>
</soap:Envelope>`;
}

function wmsHeaders(warehouseNo) {
  const token = process.env.WMS_BEARER_TOKEN || process.env.WMS_AUTHORIZATION;
  if (!token) {
    const error = new Error("Missing WMS_BEARER_TOKEN in .env.");
    error.status = 400;
    throw error;
  }
  return {
    "Content-Type": "text/xml; charset=UTF-8",
    Accept: "application/xml, text/xml, */*; q=0.01",
    Authorization: token.startsWith("Bearer ") ? token : `Bearer ${token}`,
    routerule: `1,1,${warehouseNo}`
  };
}

function decodeXmlText(value) {
  return String(value || "")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

function parseResponseText(xmlText) {
  const decoded = decodeXmlText(xmlText);
  const start = decoded.indexOf("{");
  const end = decoded.lastIndexOf("}");
  if (start === -1 || end === -1 || end <= start) return {};
  try {
    return JSON.parse(decoded.slice(start, end + 1));
  } catch {
    return {};
  }
}

function getRows(result) {
  let rows = result.data || result.rows || result.list || result.result || [];
  if (typeof rows === "string") {
    try {
      rows = JSON.parse(rows);
    } catch {
      rows = [];
    }
  }
  return Array.isArray(rows) ? rows : [];
}

function readAny(row, keys) {
  for (const key of keys) {
    if (row?.[key] != null && row[key] !== "") return row[key];
  }
  return "";
}

function resolveWarehouseNo(warehouse, warehouseNo) {
  const key = String(warehouse || "").trim();
  if (WAREHOUSE_NO_BY_KEY[key]) return WAREHOUSE_NO_BY_KEY[key];
  return String(warehouseNo || DEFAULT_WAREHOUSE_NO).trim();
}

async function unitCacheCollection() {
  const db = await getDb();
  const collection = db.collection("weekly_unit_daily_cache");
  await collection.createIndex({ warehouseKey: 1, businessDate: 1 }, { unique: true });
  return collection;
}

async function pickingCacheCollection() {
  const db = await getDb();
  const collection = db.collection("weekly_picking_daily_cache");
  await collection.createIndex({ warehouseKey: 1, businessDate: 1 }, { unique: true });
  return collection;
}

async function readCachedDaily({ warehouseKey, dates }) {
  const collection = await unitCacheCollection();
  const docs = await collection
    .find({ warehouseKey, businessDate: { $in: dates } })
    .project({ _id: 0, businessDate: 1, orders: 1, units: 1, status: 1 })
    .toArray();
  const byDate = new Map(
    docs
      .filter((doc) => doc.status === "final" || (!doc.status && cacheStatusForDate(doc.businessDate) === "final"))
      .map((doc) => [doc.businessDate, doc])
  );
  return {
    byDate,
    missingDates: dates.filter((date) => !byDate.has(date))
  };
}

async function writeCachedDaily({ warehouseKey, warehouseNo, dates, daily }) {
  const collection = await unitCacheCollection();
  const byDate = new Map(daily.map((row) => [row.业务日期, row]));
  const now = new Date();
  const writableDates = dates.filter((date) => cacheStatusForDate(date) !== "skip");
  if (!writableDates.length) return;

  await collection.bulkWrite(
    writableDates.map((date) => {
      const row = byDate.get(date);
      const status = cacheStatusForDate(date);
      return {
        updateOne: {
          filter: { warehouseKey, businessDate: date },
          update: {
            $set: {
              warehouseKey,
              warehouseNo,
              businessDate: date,
              orders: Number(row?.单量 || 0),
              units: Number(row?.件量 || 0),
              status,
              updatedAt: now
            },
            $setOnInsert: { createdAt: now }
          },
          upsert: true
        }
      };
    })
  );
}

function dailyFromCachedDocs(dates, byDate) {
  return dates
    .map((date) => {
      const doc = byDate.get(date);
      return { 业务日期: date, 单量: Number(doc?.orders || 0), 件量: Number(doc?.units || 0) };
    })
    .filter((row) => row.单量 || row.件量);
}

function filterDailyByRequestedDates(daily, requestedDates) {
  const requested = new Set(requestedDates);
  return (daily || []).filter((row) => requested.has(row.业务日期));
}

function filterPickingRowsByRequestedDates(rows, requestedDates) {
  const requested = new Set(requestedDates);
  return (rows || []).filter((row) => requested.has(dayKeyFromText(row.拣货完成时间)));
}

async function readCachedPickingRows({ warehouseKey, dates }) {
  const collection = await pickingCacheCollection();
  const docs = await collection
    .find({ warehouseKey, businessDate: { $in: dates } })
    .project({ _id: 0, businessDate: 1, rows: 1, status: 1 })
    .toArray();
  const byDate = new Map(
    docs
      .filter((doc) => doc.status === "final" || (!doc.status && cacheStatusForDate(doc.businessDate) === "final"))
      .map((doc) => [doc.businessDate, doc])
  );
  return {
    rows: dates.flatMap((date) => byDate.get(date)?.rows || []),
    byDate,
    missingDates: dates.filter((date) => !byDate.has(date))
  };
}

function groupPickingRowsByDate(rows) {
  const byDate = new Map();
  for (const row of rows || []) {
    const date = dayKeyFromText(row.拣货完成时间);
    if (!date) continue;
    if (!byDate.has(date)) byDate.set(date, []);
    byDate.get(date).push(row);
  }
  return byDate;
}

async function writeCachedPickingRows({ warehouseKey, warehouseNo, dates, rows }) {
  const collection = await pickingCacheCollection();
  const rowsByDate = groupPickingRowsByDate(rows);
  const now = new Date();
  const writableDates = dates.filter((date) => cacheStatusForDate(date) !== "skip");
  if (!writableDates.length) return;

  await collection.bulkWrite(
    writableDates.map((date) => {
      const status = cacheStatusForDate(date);
      return {
        updateOne: {
          filter: { warehouseKey, businessDate: date },
          update: {
            $set: {
              warehouseKey,
              warehouseNo,
              businessDate: date,
              rows: rowsByDate.get(date) || [],
              status,
              updatedAt: now
            },
            $setOnInsert: { createdAt: now }
          },
          upsert: true
        }
      };
    })
  );
}

function buildUnitQueryBody({ from, to, warehouseNo, currentPage = 1, pageSize = 1000 }) {
  const { startTime, endTime } = normalizeDateRange(from, to);
  const arg0 = {
    bizType: "queryReportByCondition",
    uuid: "1",
    callCode: "360BUY.WMS3.WS.CALLCODE.10401"
  };
  const arg1 = {
    Id: "wms_order_complex_query",
    Name: "customerOrderIntegratedQuery",
    WkNo: process.env.WMS_WK_NO || DEFAULT_WMS_WK_NO,
    UserName: process.env.WMS_USER_NAME || DEFAULT_WMS_USER,
    ReportModelId: "",
    SqlLimit: "5000",
    ListSqlOrder: [],
    ListSqlWhere: [
      {
        FirstValue: startTime,
        SecondValue: endTime,
        Compare: 9,
        FieldId: "RTM.S10_TIME",
        FieldName: "packingCompletionTime",
        Location: ""
      }
    ],
    PageSize: pageSize,
    CurrentPage: currentPage,
    orgNo: process.env.WMS_ORG_NO || "1",
    distributeNo: process.env.WMS_DISTRIBUTE_NO || "1",
    warehouseNo
  };
  return buildSoapEnvelope(arg0, arg1);
}

function buildPickingQueryBody({ from, to, warehouseNo, currentPage = 1, pageSize = 1000, bigWave = false }) {
  const { startTime, endTime } = normalizeDateRange(from, to);
  const arg0 = {
    bizType: "queryReportByCondition",
    uuid: "1",
    callCode: "360BUY.WMS3.WS.CALLCODE.10401"
  };
  const arg1 = {
    Id: bigWave ? "wms_bigWave_picking_result" : "wms_picking_data_v2",
    Name: bigWave ? "bigWavePickingResultsQuery" : "pickingResultsOfQuery",
    WkNo: process.env.WMS_WK_NO || DEFAULT_WMS_WK_NO,
    UserName: process.env.WMS_USER_NAME || DEFAULT_WMS_USER,
    ReportModelId: "",
    SqlLimit: "5000",
    ListSqlOrder: [],
    ListSqlWhere: [
      {
        FirstValue: startTime,
        SecondValue: endTime,
        Compare: 9,
        FieldId: bigWave ? "d.update_time" : "UPDATE_TIME",
        FieldName: "pickingDate",
        Location: ""
      }
    ],
    PageSize: pageSize,
    CurrentPage: currentPage,
    orgNo: process.env.WMS_ORG_NO || "1",
    distributeNo: process.env.WMS_DISTRIBUTE_NO || "1",
    warehouseNo
  };
  return buildSoapEnvelope(arg0, arg1);
}

async function postWms(body, warehouseNo) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 60000);
  try {
    const response = await fetch(WMS_URL, {
      method: "POST",
      headers: wmsHeaders(warehouseNo),
      body,
      signal: controller.signal
    });
    if (!response.ok) throw new Error(`WMS request failed with ${response.status}`);
    return parseResponseText(await response.text());
  } catch (error) {
    if (error?.name === "AbortError" || String(error?.message || "").toLowerCase().includes("aborted")) {
      const timeoutError = new Error("WMS query timed out. Please try a smaller date range or query again.");
      timeoutError.status = 504;
      throw timeoutError;
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

export function summarizeUnitRows(rows) {
  const unitsByDate = {};
  const orderSets = {};
  for (const row of rows || []) {
    const canceled = String(readAny(row, ["是否取消", "isCancel", "cancelFlag"]) || "").trim();
    if (["是", "yes", "true", "1"].includes(canceled.toLowerCase())) continue;

    const date = dayKey(parseDate(readAny(row, ["打包完成时间", "packingCompletionTime", "pcEndTime", "RTM.S10_TIME", "s10Time", "endTime"])));
    if (!date) continue;
    const units = num(readAny(row, ["件数", "实际件数", "piecesQty", "actualQty", "qty", "quantity", "skuQty"]), 0);
    unitsByDate[date] = (unitsByDate[date] || 0) + units;

    const orderNo = String(readAny(row, ["JDOrderNO", "京东订单号", "jdOrderNo", "outboundNo", "salesPlatformOrder", "channelsOutboundNo", "orderNo", "customerOrderNo", "obmOrderNo"]) || "").trim();
    if (orderNo) {
      if (!orderSets[date]) orderSets[date] = new Set();
      orderSets[date].add(orderNo);
    }
  }
  const ordersByDate = {};
  for (const [date, orders] of Object.entries(orderSets)) ordersByDate[date] = orders.size;
  return { unitsByDate, ordersByDate };
}

export function weeklyDailyFromUnitRows(rows) {
  const { unitsByDate, ordersByDate } = summarizeUnitRows(rows);
  return Object.entries(unitsByDate)
    .map(([date, units]) => ({ 业务日期: date, 单量: ordersByDate[date] || 0, 件量: units }))
    .sort((a, b) => String(a.业务日期).localeCompare(String(b.业务日期)));
}

export async function queryUnitData({ from, to, warehouse, warehouseNo = DEFAULT_WAREHOUSE_NO } = {}) {
  const warehouseKey = String(warehouse || "5").trim();
  const cleanWarehouseNo = resolveWarehouseNo(warehouse, warehouseNo);
  const { startDate, endDate } = normalizeDateRange(from, to);
  const requestedDates = dateKeysBetween(startDate, endDate);
  let cacheStatus = { enabled: true };
  let cached = { byDate: new Map(), missingDates: requestedDates };

  try {
    cached = await readCachedDaily({ warehouseKey, dates: requestedDates });
  } catch (error) {
    cacheStatus = { enabled: false, error: error.message };
  }

  if (cacheStatus.enabled && !cached.missingDates.length) {
    const daily = dailyFromCachedDocs(requestedDates, cached.byDate);
    return {
      kpi: {
        totalOrders: sum(daily, (row) => row.单量),
        totalUnits: sum(daily, (row) => row.件量),
        warehouseName: "",
        targetUpph: TARGET_UPPH_BY_WAREHOUSE[warehouseKey] || ""
      },
      daily,
      rawCount: 0,
      source: "cache",
      cacheStatus
    };
  }

  const pageSize = 1000;
  const rows = [];
  const maxPages = 50;

  for (let currentPage = 1; currentPage <= maxPages; currentPage++) {
    const result = await postWms(buildUnitQueryBody({ from, to, warehouseNo: cleanWarehouseNo, currentPage, pageSize }), cleanWarehouseNo);
    const pageRows = getRows(result);
    rows.push(...pageRows);
    if (pageRows.length < pageSize) break;
  }

  const fetchedDaily = filterDailyByRequestedDates(weeklyDailyFromUnitRows(rows), requestedDates);
  let daily = fetchedDaily;

  if (cacheStatus.enabled) {
    try {
      await writeCachedDaily({ warehouseKey, warehouseNo: cleanWarehouseNo, dates: requestedDates, daily: fetchedDaily });
    } catch (error) {
      cacheStatus = { enabled: false, error: error.message };
    }
  }

  return {
    kpi: {
      totalOrders: sum(daily, (row) => row.单量),
      totalUnits: sum(daily, (row) => row.件量),
      warehouseName: "",
      targetUpph: TARGET_UPPH_BY_WAREHOUSE[warehouseKey] || ""
    },
    daily,
    rawCount: rows.length,
    source: cacheStatus.enabled && cached.byDate.size ? "mixed" : "wms",
    cacheStatus
  };
}

function normalizePickingRows(rows) {
  return (rows || []).map((row) => ({
    拣货完成时间: readAny(row, ["拣货完成时间", "updateTime", "update_time", "pickingDate", "pickingTime"]),
    拣货开始时间: readAny(row, ["拣货开始时间", "fetchTime", "fetch_time", "startTime", "start_time", "createTime"]),
    任务领取时间: readAny(row, ["任务领取时间", "fetchTime", "fetch_time", "receiveTime", "receive_time", "startTime", "createTime"]),
    实际拣货量: readAny(row, ["实际拣货量", "拣货数量", "拣货件数", "pickingQty", "GOODS_SUM", "goodsSum", "locateQty", "qty", "quantity"]) || 0,
    工号: readAny(row, ["工号", "员工号", "employeeNo", "employee_no", "email", "workerName", "updateUser"]),
    员工号: readAny(row, ["员工号", "工号", "employeeNo", "employee_no", "email", "workerName", "updateUser"]),
    姓名: readAny(row, ["姓名", "workerName", "worker_name", "updateUser", "update_user", "employeeName", "employeeNo"]) || "",
    集合单号: readAny(row, ["集合单号", "任务单号", "taskPageNo", "task_page_no", "outboundNo", "waveNo", "wave_no", "batchNo", "bigWaveNo"]),
    任务单号: readAny(row, ["任务单号", "集合单号", "taskPageNo", "task_page_no", "outboundNo", "waveNo", "wave_no", "batchNo", "bigWaveNo"]),
    储位: readAny(row, ["储位", "库位", "cellNo", "cell_no", "containerNo", "goodsNo", "outboundNo"])
  }));
}

async function fetchPickingRows({ from, to, warehouseNo, bigWave = false }) {
  const pageSize = bigWave ? 100 : 1000;
  const rows = [];
  const maxPages = 50;

  for (let currentPage = 1; currentPage <= maxPages; currentPage++) {
    const result = await postWms(buildPickingQueryBody({ from, to, warehouseNo, currentPage, pageSize, bigWave }), warehouseNo);
    const pageRows = getRows(result);
    rows.push(...pageRows);
    if (pageRows.length < pageSize) break;
  }

  return rows;
}

export async function queryPickingData({ from, to, warehouse, warehouseNo = DEFAULT_WAREHOUSE_NO, targetUpph = "", includeBigWavePick = false } = {}) {
  const warehouseKey = String(warehouse || "5").trim();
  const cleanWarehouseNo = resolveWarehouseNo(warehouse, warehouseNo);
  const { startDate, endDate } = normalizeDateRange(from, to);
  const requestedDates = dateKeysBetween(startDate, endDate);
  const useCache = !includeBigWavePick;
  let cacheStatus = { enabled: useCache };
  let cached = { rows: [], byDate: new Map(), missingDates: requestedDates };

  if (useCache) {
    try {
      cached = await readCachedPickingRows({ warehouseKey, dates: requestedDates });
    } catch (error) {
      cacheStatus = { enabled: false, error: error.message };
    }
  }

  if (useCache && cacheStatus.enabled && !cached.missingDates.length) {
    const analysis = analyzePickRows(cached.rows, null, Number(targetUpph) || "", []);
    return {
      ...analysis,
      rawCount: cached.rows.length,
      source: "cache",
      cacheStatus
    };
  }

  const regularRows = await fetchPickingRows({ from, to, warehouseNo: cleanWarehouseNo });
  const bigWaveRows = includeBigWavePick
    ? await fetchPickingRows({ from, to, warehouseNo: cleanWarehouseNo, bigWave: true })
    : [];
  const rows = [...regularRows, ...bigWaveRows];

  const normalizedRows = filterPickingRowsByRequestedDates(normalizePickingRows(rows), requestedDates);

  if (useCache && cacheStatus.enabled) {
    try {
      await writeCachedPickingRows({ warehouseKey, warehouseNo: cleanWarehouseNo, dates: requestedDates, rows: normalizedRows });
    } catch (error) {
      cacheStatus = { enabled: false, error: error.message };
    }
  }

  const analysis = analyzePickRows(normalizedRows, null, Number(targetUpph) || "", []);
  return {
    ...analysis,
    rawCount: rows.length,
    source: includeBigWavePick ? "wms+big-wave" : (cacheStatus.enabled && cached.byDate.size ? "mixed" : "wms"),
    regularRawCount: regularRows.length,
    bigWaveRawCount: bigWaveRows.length,
    cacheStatus
  };
}
