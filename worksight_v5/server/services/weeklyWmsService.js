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
const DEFAULT_RANKING_WAREHOUSES = ["1", "2", "5"];
const RANKING_SCHEMA_VERSION = 18;
const FLOOR_WEIGHT = { L1: 1, L2: 1.1, L3: 1.2, L4: 1.3 };
const HOUR_MS = 60 * 60 * 1000;
let rankingSchedulerTimer;

function pad(value) {
  return String(value).padStart(2, "0");
}

function formatDate(date) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
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

function dateRangesFromKeys(dates) {
  const sortedDates = [...new Set(dates || [])].sort();
  const ranges = [];
  let rangeStart = "";
  let previous = "";

  for (const date of sortedDates) {
    if (!rangeStart) {
      rangeStart = date;
      previous = date;
      continue;
    }

    const expectedNext = dateKeysBetween(previous, date)[1];
    if (expectedNext === date) {
      previous = date;
      continue;
    }

    ranges.push({ from: rangeStart, to: previous });
    rangeStart = date;
    previous = date;
  }

  if (rangeStart) ranges.push({ from: rangeStart, to: previous });
  return ranges;
}

function todayKey() {
  return dayKey(new Date());
}

function cacheStatusForDate(date) {
  if (date < todayKey()) return "final";
  if (date === todayKey()) return "provisional";
  return "skip";
}

function logDailySources(label, { requestedDates, cachedByDate, datesToFetch, cacheStatus }) {
  const apiDates = new Set(datesToFetch || []);
  const cacheEnabled = Boolean(cacheStatus?.enabled);
  console.log(`[${label}] data source by date:`);
  for (const date of requestedDates || []) {
    const source = cacheEnabled && cachedByDate?.has(date) && !apiDates.has(date) ? "database" : "api";
    console.log(`[${label}] ${date} -> ${source}`);
  }
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
  const collection = db.collection("weekly_picking_daily_cache_v3");
  await collection.createIndex({ warehouseKey: 1, businessDate: 1 }, { unique: true });
  return collection;
}

async function pickingRankingCacheCollection() {
  const db = await getDb();
  const collection = db.collection("weekly_picking_ranking_cache");
  await collection.createIndex({ warehouseKey: 1, period: 1, startDate: 1, endDate: 1 }, { unique: true });
  return collection;
}

async function pickingRankingJobCollection() {
  const db = await getDb();
  const collection = db.collection("weekly_picking_ranking_jobs");
  await collection.createIndex({ jobKey: 1 }, { unique: true });
  return collection;
}

async function pickingRankingExclusionCollection() {
  const db = await getDb();
  const collection = db.collection("weekly_picking_ranking_exclusions");
  await collection.createIndex({ employeeNo: 1 }, { unique: true, sparse: true });
  await collection.createIndex({ name: 1 }, { sparse: true });
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

function buildPackingQueryBody({ from, to, warehouseNo, currentPage = 1, pageSize = 100 }) {
  const { startTime, endTime } = normalizeDateRange(from, to);
  const arg0 = {
    bizType: "queryReportByCondition",
    uuid: "1",
    callCode: "360BUY.WMS3.WS.CALLCODE.10401"
  };
  const arg1 = {
    Id: "aos_check_platform",
    Name: "checkPackageNumber",
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
        FieldId: "M.CREATE_TIME",
        FieldName: "packingTime",
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

async function fetchUnitRowsForDates({ dates, warehouseNo }) {
  const rows = [];
  const pageSize = 1000;
  const maxPages = 50;

  for (const range of dateRangesFromKeys(dates)) {
    for (let currentPage = 1; currentPage <= maxPages; currentPage++) {
      const result = await postWms(buildUnitQueryBody({ ...range, warehouseNo, currentPage, pageSize }), warehouseNo);
      const pageRows = getRows(result);
      rows.push(...pageRows);
      if (pageRows.length < pageSize) break;
    }
  }

  return rows;
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
    logDailySources("weekly-unit", {
      requestedDates,
      cachedByDate: cached.byDate,
      datesToFetch: [],
      cacheStatus
    });
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

  const datesToFetch = cacheStatus.enabled ? cached.missingDates : requestedDates;
  logDailySources("weekly-unit", {
    requestedDates,
    cachedByDate: cached.byDate,
    datesToFetch,
    cacheStatus
  });
  const rows = await fetchUnitRowsForDates({ dates: datesToFetch, warehouseNo: cleanWarehouseNo });
  const fetchedDaily = filterDailyByRequestedDates(weeklyDailyFromUnitRows(rows), datesToFetch);
  let daily = cacheStatus.enabled
    ? [...dailyFromCachedDocs(requestedDates, cached.byDate), ...fetchedDaily]
      .sort((a, b) => String(a["\u4e1a\u52a1\u65e5\u671f"]).localeCompare(String(b["\u4e1a\u52a1\u65e5\u671f"])))
    : fetchedDaily;

  if (cacheStatus.enabled) {
    try {
      await writeCachedDaily({ warehouseKey, warehouseNo: cleanWarehouseNo, dates: datesToFetch, daily: fetchedDaily });
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
    ...row,
    _source_key: row._source_key || "wms",
    拣货完成时间: readAny(row, ["拣货完成时间", "updateTime", "update_time", "pickingDate", "pickingTime"]),
    拣货开始时间: readAny(row, ["拣货开始时间", "fetchTime", "fetch_time", "startTime", "start_time", "createTime"]),
    任务领取时间: readAny(row, ["任务领取时间", "fetchTime", "fetch_time", "receiveTime", "receive_time", "startTime", "createTime"]),
    实际拣货量: readAny(row, ["实际拣货量", "拣货数量", "拣货件数", "pickingQty", "GOODS_SUM", "goodsSum", "locateQty", "qty", "quantity"]) || 0,
    工号: readAny(row, ["工号", "员工号", "employeeNo", "employee_no", "email", "workerName", "updateUser"]),
    员工号: readAny(row, ["员工号", "工号", "employeeNo", "employee_no", "email", "workerName", "updateUser"]),
    姓名: readAny(row, ["姓名", "workerName", "worker_name", "updateUser", "update_user", "employeeName", "employeeNo"]) || "",
    集合单号: readAny(row, ["集合单号", "任务单号", "taskPageNo", "task_page_no", "outboundNo", "waveNo", "wave_no", "batchNo", "bigWaveNo"]),
    任务单号: readAny(row, ["任务单号", "集合单号", "taskPageNo", "task_page_no", "outboundNo", "waveNo", "wave_no", "batchNo", "bigWaveNo"]),
    储位: readAny(row, ["储位", "库位", "cellNo", "cell_no", "containerNo", "goodsNo", "outboundNo"]),
    "\u8d27\u578b": readAny(row, ["\u8d27\u578b", "??????", "goodsType", "goods_type", "sizeDefinition", "cargoType", "skuType", "productType", "goodSizeType"])
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

async function fetchPickingRowsForDates({ dates, warehouseNo, bigWave = false }) {
  const rows = [];
  for (const range of dateRangesFromKeys(dates)) {
    rows.push(...await fetchPickingRows({ ...range, warehouseNo, bigWave }));
  }
  return rows;
}

async function fetchPackingRows({ from, to, warehouseNo }) {
  const pageSize = 100;
  const rows = [];
  const maxPages = 50;

  for (let currentPage = 1; currentPage <= maxPages; currentPage++) {
    const result = await postWms(buildPackingQueryBody({ from, to, warehouseNo, currentPage, pageSize }), warehouseNo);
    const pageRows = getRows(result);
    rows.push(...pageRows);
    if (pageRows.length < pageSize) break;
  }

  return rows;
}

async function fetchPackingRowsForDates({ dates, warehouseNo }) {
  const rows = [];
  for (const date of dates || []) {
    const rangeRows = await fetchPackingRows({ from: date, to: date, warehouseNo });
    console.log(`[packing-ranking] ${date} -> ${date}: ${rangeRows.length} rows`);
    if (rangeRows[0]) {
      console.log(`[packing-ranking] sample keys: ${Object.keys(rangeRows[0]).join(", ")}`);
      console.log(`[packing-ranking] sample row: ${JSON.stringify(rangeRows[0]).slice(0, 1000)}`);
    }
    rows.push(...rangeRows);
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
    logDailySources("picking-efficiency", {
      requestedDates,
      cachedByDate: cached.byDate,
      datesToFetch: [],
      cacheStatus
    });
    return {
      ...analysis,
      rawCount: cached.rows.length,
      source: "cache",
      cacheStatus
    };
  }

  const datesToFetch = useCache && cacheStatus.enabled ? cached.missingDates : requestedDates;
  logDailySources("picking-efficiency", {
    requestedDates,
    cachedByDate: cached.byDate,
    datesToFetch,
    cacheStatus
  });
  const regularRows = (await fetchPickingRowsForDates({ dates: datesToFetch, warehouseNo: cleanWarehouseNo }))
    .map((row) => ({ ...row, _source_key: "wms-regular" }));
  const bigWaveRows = includeBigWavePick
    ? (await fetchPickingRowsForDates({ dates: datesToFetch, warehouseNo: cleanWarehouseNo, bigWave: true }))
      .map((row) => ({ ...row, _source_key: "wms-big-wave" }))
    : [];
  const rows = [...regularRows, ...bigWaveRows];

  const fetchedRows = filterPickingRowsByRequestedDates(normalizePickingRows(rows), datesToFetch);
  const normalizedRows = useCache && cacheStatus.enabled
    ? [...normalizePickingRows(cached.rows || []), ...fetchedRows]
    : fetchedRows;

  if (useCache && cacheStatus.enabled) {
    try {
      await writeCachedPickingRows({ warehouseKey, warehouseNo: cleanWarehouseNo, dates: datesToFetch, rows: fetchedRows });
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

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function previousWeekRange(now = new Date()) {
  const current = new Date(now);
  current.setHours(0, 0, 0, 0);
  const day = current.getDay() || 7;
  const thisWeekStart = addDays(current, 1 - day);
  return {
    startDate: formatDate(addDays(thisWeekStart, -7)),
    endDate: formatDate(addDays(thisWeekStart, -1))
  };
}

function previousMonthRange(now = new Date()) {
  const current = new Date(now);
  return {
    startDate: formatDate(new Date(current.getFullYear(), current.getMonth() - 1, 1)),
    endDate: formatDate(new Date(current.getFullYear(), current.getMonth(), 0))
  };
}

function firstPickValue(row, candidates) {
  return readAny(row, candidates);
}

function personEfficiencyBasics(row) {
  const keys = Object.keys(row || {});
  return {
    employeeNo: row?.["\u5de5\u53f7"] ?? row?.[keys[1]] ?? "",
    name: row?.["\u59d3\u540d"] ?? row?.[keys[2]] ?? "",
    hours: num(row?.hours ?? row?.effectiveHours ?? row?.["\u6709\u6548\u5de5\u65f6"] ?? row?.[keys[5]], 0),
    weightedUnits: num(row?.weightedUnits, 0)
  };
}

function floorFromLocation(value) {
  const match = String(value || "").toUpperCase().match(/\bL\s*([1-4])\b|L\s*([1-4])/);
  const floor = match?.[1] || match?.[2];
  return floor ? `L${floor}` : "";
}

function floorWeightFromLocation(value) {
  return FLOOR_WEIGHT[floorFromLocation(value)] || 1;
}

function goodsKind(value) {
  const type = String(value || "").trim().toUpperCase();
  return type === "T25" || type === "T50" ? "small" : "large";
}

function percent(part, total) {
  return total ? (part / total) * 100 : 0;
}

function packingPersonKey(row) {
  const scannedEmployeeNo = employeeIdFromRowValues(row);
  const scannedJdhkNo = jdhkIdFromRowValues(row);
  const emailEmployeeNo = employeeIdFromEmail(readAny(row, [
    "Packer Name",
    "packerName",
    "packer_name",
    "USER_NAME",
    "PACKER_NAME",
    "PICKER_NAME",
    "userName"
  ]));
  const fallbackEmployeeNo = employeeIdFromEmail(readAny(row, [
    "\u5de5\u53f7",
    "\u5458\u5de5\u53f7",
    "\u64cd\u4f5c\u4eba\u7f16\u7801",
    "\u6253\u5305\u4eba\u7f16\u7801",
    "\u590d\u6838\u4eba\u7f16\u7801",
    "employeeNo",
    "employee_no",
    "CREATE_USER",
    "USER_NAME",
    "PICKER_NAME",
    "PACKER_NAME",
    "operatorNo",
    "operatorCode",
    "createUser",
    "createUserCode",
    "userCode"
  ]));
  const employeeNo = scannedEmployeeNo || emailEmployeeNo || scannedJdhkNo || fallbackEmployeeNo;
  const name = identityText(readAny(row, [
    "\u59d3\u540d",
    "\u64cd\u4f5c\u4eba",
    "\u6253\u5305\u4eba",
    "\u590d\u6838\u4eba",
    "operatorName",
    "createUserName",
    "USER_NAME",
    "CREATE_USER",
    "userName",
    "realName"
  ]));
  return {
    employeeNo,
    name,
    key: employeeNo ? employeeKey(employeeNo) : identityText(name)
  };
}

function packingUnitsValue(row) {
  const value = readAny(row, [
    "Actual Pack Quantity",
    "actualPackQuantity",
    "ACTUAL_PACK_QUANTITY",
    "\u4ef6\u6570",
    "\u5305\u88f9\u6570",
    "\u6253\u5305\u4ef6\u6570",
    "\u590d\u6838\u4ef6\u6570",
    "\u6570\u91cf",
    "packingUnits",
    "packageQty",
    "PACKAGE_QTY",
    "packageNum",
    "packageCount",
    "checkPackageNumber",
    "checkPackageNum",
    "qty",
    "quantity",
    "num",
    "count"
  ]);
  const units = num(value, NaN);
  return Number.isFinite(units) ? units : 1;
}

function packingTimeValue(row) {
  return parseDate(readAny(row, [
    "packingTime",
    "Packing Time",
    "PACKING_TIME",
    "\u6253\u5305\u65f6\u95f4",
    "\u590d\u6838\u65f6\u95f4",
    "\u64cd\u4f5c\u65f6\u95f4",
    "\u521b\u5efa\u65f6\u95f4",
    "CREATE_TIME",
    "UPDATE_TIME",
    "createTime",
    "create_time",
    "M.CREATE_TIME",
    "operateTime",
    "operationTime",
    "updateTime"
  ]));
}

function summarizePackingByPerson(packingRows = []) {
  const byPerson = new Map();
  const spanByPersonDay = new Map();
  const debugPackingKeys = new Map();

  for (const row of packingRows || []) {
    const person = packingPersonKey(row);
    if (!person.key) continue;
    const rowText = JSON.stringify(row);
    if (rowText.includes("OSJlpSVwbftE")) {
      const debug = debugPackingKeys.get(person.key) || { rows: 0, units: 0, sample: row };
      debug.rows += 1;
      debug.units += packingUnitsValue(row);
      debugPackingKeys.set(person.key, debug);
    }

    const current = byPerson.get(person.key) || {
      employeeNo: person.employeeNo,
      name: person.name,
      packingUnits: 0,
      packingHours: 0,
      packingRows: 0
    };
    if (!current.employeeNo && person.employeeNo) current.employeeNo = person.employeeNo;
    if (!current.name && person.name) current.name = person.name;
    current.packingUnits += packingUnitsValue(row);
    current.packingRows += 1;
    byPerson.set(person.key, current);

    const time = packingTimeValue(row);
    if (!time) continue;
    const spanKey = `${person.key}|${dayKey(time) || ""}`;
    const span = spanByPersonDay.get(spanKey) || { personKey: person.key, first: time, last: time };
    if (time < span.first) span.first = time;
    if (time > span.last) span.last = time;
    spanByPersonDay.set(spanKey, span);
  }

  for (const span of spanByPersonDay.values()) {
    const current = byPerson.get(span.personKey);
    if (current) current.packingHours += Math.max(0, (span.last - span.first) / 3600000);
  }

  const totals = [...byPerson.values()].reduce((acc, row) => {
    acc.people += 1;
    acc.rows += row.packingRows || 0;
    acc.units += row.packingUnits || 0;
    acc.hours += row.packingHours || 0;
    return acc;
  }, { people: 0, rows: 0, units: 0, hours: 0 });
  console.log(`[packing-ranking] summarized ${totals.people} people, rows=${totals.rows}, units=${totals.units}, hours=${totals.hours.toFixed(2)}`);
  for (const row of [...byPerson.values()].sort((a, b) => (b.packingUnits || 0) - (a.packingUnits || 0)).slice(0, 15)) {
    console.log(`[packing-ranking] person ${row.employeeNo || row.name}: rows=${row.packingRows || 0}, units=${row.packingUnits || 0}, hours=${(row.packingHours || 0).toFixed(2)}`);
  }
  for (const [key, debug] of debugPackingKeys.entries()) {
    console.log(`[packing-ranking] debug OSJlpSVwbftE -> ${key}: rows=${debug.rows}, units=${debug.units}`);
  }

  return byPerson;
}

function buildPickingRankingRows(personEfficiency, pickRows = [], packingRows = []) {
  const byPerson = new Map();
  const packingByPerson = summarizePackingByPerson(packingRows);
  for (const row of personEfficiency || []) {
    const item = personEfficiencyBasics(row);
    const key = item.employeeNo ? employeeKey(item.employeeNo) : String(item.name || "").trim();
    if (!key) continue;
    const current = byPerson.get(key) || {
      employeeNo: item.employeeNo,
      name: item.name,
      weightedUnits: 0,
      hours: 0,
      pickingWeightedUnits: 0,
      pickingHours: 0,
      packingUnits: 0,
      packingHours: 0
    };
    current.weightedUnits += item.weightedUnits;
    current.hours += item.hours;
    current.pickingWeightedUnits += item.weightedUnits;
    current.pickingHours += item.hours;
    byPerson.set(key, current);
  }

  const spanByPersonDay = new Map();
  const mixByPerson = new Map();
  for (const row of pickRows || []) {
    const completion = parseDate(firstPickValue(row, ["\u62e3\u8d27\u5b8c\u6210\u65f6\u95f4", "updateTime", "update_time", "pickingDate", "pickingTime"]));
    if (!completion) continue;
    const employeeNo = identityText(firstPickValue(row, ["\u5de5\u53f7", "\u5458\u5de5\u53f7", "employeeNo", "employee_no", "email", "workerName", "updateUser"]));
    const name = identityText(firstPickValue(row, ["\u59d3\u540d", "workerName", "worker_name", "updateUser", "update_user", "employeeName", "employeeNo"]));
    const key = employeeNo ? employeeKey(employeeNo) : identityText(name);
    if (!key || !byPerson.has(key)) continue;
    const spanKey = `${key}|${dayKey(completion) || ""}`;
    const current = spanByPersonDay.get(spanKey) || { personKey: key, first: completion, last: completion };
    if (completion < current.first) current.first = completion;
    if (completion > current.last) current.last = completion;
    spanByPersonDay.set(spanKey, current);

    const qty = num(firstPickValue(row, ["\u5b9e\u9645\u62e3\u8d27\u91cf", "\u62e3\u8d27\u6570\u91cf", "\u62e3\u8d27\u4ef6\u6570", "pickingQty", "GOODS_SUM", "goodsSum", "locateQty", "qty", "quantity"]), 0);
    const location = firstPickValue(row, ["\u50a8\u4f4d", "\u5e93\u4f4d", "cellNo", "cell_no", "containerNo", "goodsNo", "outboundNo"]);
    const cargoKind = goodsKind(firstPickValue(row, ["\u8d27\u578b", "goodsType", "goods_type", "sizeDefinition", "cargoType", "skuType", "productType", "goodSizeType"]));
    const weightedQty = qty * (cargoKind === "small" ? 1 : 2.5) * floorWeightFromLocation(location);
    const mix = mixByPerson.get(key) || {
      totalQty: 0,
      weightedUnits: 0,
      floorQty: { L1: 0, L2: 0, L3: 0, L4: 0 },
      smallQty: 0,
      largeQty: 0
    };
    const floor = floorFromLocation(location);
    if (floor) mix.floorQty[floor] += qty;
    if (cargoKind === "small") mix.smallQty += qty;
    else mix.largeQty += qty;
    mix.totalQty += qty;
    mix.weightedUnits += weightedQty;
    mixByPerson.set(key, mix);
  }

  const hoursByPerson = new Map();
  for (const span of spanByPersonDay.values()) {
    hoursByPerson.set(
      span.personKey,
      (hoursByPerson.get(span.personKey) || 0) + Math.max(0, (span.last - span.first) / 3600000)
    );
  }

  for (const [key, hours] of hoursByPerson.entries()) {
    const current = byPerson.get(key);
    if (current) {
      current.hours = hours;
      current.pickingHours = hours;
    }
  }

  for (const [key, packing] of packingByPerson.entries()) {
    if (byPerson.has(key)) continue;
    byPerson.set(key, {
      employeeNo: packing.employeeNo,
      name: packing.name,
      weightedUnits: 0,
      hours: 0,
      pickingWeightedUnits: 0,
      pickingHours: 0,
      packingUnits: packing.packingUnits,
      packingHours: packing.packingHours
    });
  }

  return [...byPerson.values()]
    .map((row) => {
      const rowKey = row.employeeNo ? employeeKey(row.employeeNo) : String(row.name || "").trim();
      const mix = mixByPerson.get(rowKey) || {
        totalQty: 0,
        weightedUnits: row.weightedUnits,
        floorQty: { L1: 0, L2: 0, L3: 0, L4: 0 },
        smallQty: 0,
        largeQty: 0
      };
      const weightedUnits = mix.totalQty ? mix.weightedUnits : row.weightedUnits;
      const packing = packingByPerson.get(row.employeeNo ? employeeKey(row.employeeNo) : "")
        || packingByPerson.get(String(row.name || "").trim())
        || {};
      const pickingWeightedUnits = weightedUnits;
      const pickingHours = row.pickingHours ?? row.hours;
      const packingUnits = num(packing.packingUnits, 0);
      const packingHours = num(packing.packingHours, 0);
      const totalUnits = pickingWeightedUnits + packingUnits;
      const totalHours = pickingHours + packingHours;
      return {
        rank: 0,
        employeeNo: row.employeeNo,
        name: row.name,
        weightedUnits,
        pickingWeightedUnits,
        packingUnits,
        hours: pickingHours,
        pickingHours,
        packingHours,
        effectiveHours: pickingHours,
        efficiency: totalHours ? totalUnits / totalHours : 0,
        l1Percent: percent(mix.floorQty.L1, mix.totalQty),
        l2Percent: percent(mix.floorQty.L2, mix.totalQty),
        l3Percent: percent(mix.floorQty.L3, mix.totalQty),
        l4Percent: percent(mix.floorQty.L4, mix.totalQty),
        smallPercent: percent(mix.smallQty, mix.totalQty),
        largePercent: percent(mix.largeQty, mix.totalQty),
        mainCargo: mix.largeQty > mix.smallQty ? "Large" : mix.smallQty > mix.largeQty ? "Small" : "Mixed"
      };
    })
    .filter((row) => row.pickingWeightedUnits > 0 || row.pickingHours > 0 || row.packingUnits > 0 || row.packingHours > 0)
    .sort((a, b) => b.efficiency - a.efficiency || b.pickingWeightedUnits - a.pickingWeightedUnits)
    .map((row, index) => ({ ...row, rank: index + 1 }));
}
function identityText(value) {
  return String(value || "").trim();
}

function employeeIdFromEmail(value) {
  const text = identityText(value);
  if (!text.includes("@")) return text;
  return text.split("@")[0].trim();
}

function employeeIdFromRowValues(row) {
  for (const value of Object.values(row || {})) {
    const match = String(value || "").match(/\b(US\d{4,})\b/i);
    if (match) return match[1].toUpperCase();
  }
  return "";
}

function jdhkIdFromRowValues(row) {
  for (const value of Object.values(row || {})) {
    const match = String(value || "").match(/\b(jdhk_[A-Za-z0-9]+)\b/);
    if (match) return match[1];
  }
  return "";
}

function employeeKey(value) {
  return employeeIdFromEmail(value).toUpperCase();
}

function exclusionKeys(row) {
  return [
    identityText(row?.employeeNo).toLowerCase(),
    identityText(row?.name).toLowerCase()
  ].filter(Boolean);
}

function isExcludedRankingRow(row, excluded) {
  return exclusionKeys(row).some((key) => excluded.has(key));
}

function rerankRows(rows) {
  return [...(rows || [])]
    .sort((a, b) => b.efficiency - a.efficiency || (b.pickingWeightedUnits ?? b.weightedUnits ?? 0) - (a.pickingWeightedUnits ?? a.weightedUnits ?? 0))
    .map((row, index) => ({ ...row, rank: index + 1 }));
}

function recalculateRankingRow(row) {
  const pickingWeightedUnits = num(row.pickingWeightedUnits ?? row.weightedUnits, 0);
  const pickingHours = num(row.pickingHours ?? row.hours ?? row.effectiveHours, 0);
  const packingUnits = num(row.packingUnits, 0);
  const packingHours = num(row.packingHours, 0);
  const totalUnits = pickingWeightedUnits + packingUnits;
  const totalHours = pickingHours + packingHours;
  return {
    ...row,
    weightedUnits: pickingWeightedUnits,
    pickingWeightedUnits,
    hours: pickingHours,
    pickingHours,
    effectiveHours: pickingHours,
    packingUnits,
    packingHours,
    efficiency: totalHours ? totalUnits / totalHours : 0
  };
}

function mergeRankingRows(target, source) {
  const targetPickingUnits = num(target.pickingWeightedUnits ?? target.weightedUnits, 0);
  const sourcePickingUnits = num(source.pickingWeightedUnits ?? source.weightedUnits, 0);
  const merged = {
    ...target,
    pickingWeightedUnits: targetPickingUnits + sourcePickingUnits,
    weightedUnits: targetPickingUnits + sourcePickingUnits,
    pickingHours: num(target.pickingHours ?? target.hours ?? target.effectiveHours, 0) + num(source.pickingHours ?? source.hours ?? source.effectiveHours, 0),
    hours: num(target.pickingHours ?? target.hours ?? target.effectiveHours, 0) + num(source.pickingHours ?? source.hours ?? source.effectiveHours, 0),
    effectiveHours: num(target.pickingHours ?? target.hours ?? target.effectiveHours, 0) + num(source.pickingHours ?? source.hours ?? source.effectiveHours, 0),
    packingUnits: num(target.packingUnits, 0) + num(source.packingUnits, 0),
    packingHours: num(target.packingHours, 0) + num(source.packingHours, 0)
  };

  if (!targetPickingUnits && sourcePickingUnits) {
    merged.l1Percent = source.l1Percent;
    merged.l2Percent = source.l2Percent;
    merged.l3Percent = source.l3Percent;
    merged.l4Percent = source.l4Percent;
    merged.smallPercent = source.smallPercent;
    merged.largePercent = source.largePercent;
    merged.mainCargo = source.mainCargo;
  }

  return recalculateRankingRow(merged);
}

async function readRankingExclusions() {
  const collection = await pickingRankingExclusionCollection();
  const docs = await collection.find({}).project({ _id: 0, employeeNo: 1, name: 1 }).toArray();
  return new Set(docs.flatMap(exclusionKeys));
}

async function filterExcludedRankingRows(rows) {
  const excluded = await readRankingExclusions();
  if (!excluded.size) return rows || [];
  return rerankRows((rows || []).filter((row) => !isExcludedRankingRow(row, excluded)));
}

async function filterRankingDoc(doc) {
  if (!doc) return doc;
  return { ...doc, rows: await filterExcludedRankingRows(doc.rows || []) };
}

function rankingDocPayload({ period, warehouseKey, warehouseNo, startDate, endDate, rows, source, rawCount }) {
  return {
    schemaVersion: RANKING_SCHEMA_VERSION,
    period,
    warehouseKey,
    warehouseNo,
    startDate,
    endDate,
    rows,
    source,
    rawCount,
    updatedAt: new Date()
  };
}

export async function getOrCreatePickingRanking({ period, from, to, warehouse, warehouseNo = DEFAULT_WAREHOUSE_NO, forceRefresh = false } = {}) {
  const warehouseKey = String(warehouse || "5").trim();
  const cleanWarehouseNo = resolveWarehouseNo(warehouseKey, warehouseNo);
  const startDate = from;
  const endDate = to;
  const collection = await pickingRankingCacheCollection();
  const filter = { warehouseKey, period, startDate, endDate };

  if (!forceRefresh) {
    const cached = await collection.findOne(filter, { projection: { _id: 0 } });
    if (cached?.schemaVersion === RANKING_SCHEMA_VERSION) return { ...(await filterRankingDoc(cached)), source: "database" };
  }

  const analysis = await queryPickingData({
    from: startDate,
    to: endDate,
    warehouse: warehouseKey,
    warehouseNo: cleanWarehouseNo,
    targetUpph: TARGET_UPPH_BY_WAREHOUSE[warehouseKey] || "",
    includeBigWavePick: false
  });
  const cachedPickRows = await readCachedPickingRows({ warehouseKey, dates: dateKeysBetween(startDate, endDate) });
  const sourceRows = filterPickingRowsByRequestedDates(normalizePickingRows(cachedPickRows.rows || []), dateKeysBetween(startDate, endDate));
  const packingRows = await fetchPackingRowsForDates({
    dates: dateKeysBetween(startDate, endDate),
    warehouseNo: cleanWarehouseNo
  });
  const rows = await filterExcludedRankingRows(buildPickingRankingRows(analysis.personEfficiency || [], sourceRows, packingRows));
  const payload = rankingDocPayload({
    period,
    warehouseKey,
    warehouseNo: cleanWarehouseNo,
    startDate,
    endDate,
    rows,
    source: analysis.source,
    rawCount: analysis.rawCount || 0
  });

  await collection.updateOne(
    filter,
    { $set: payload, $setOnInsert: { createdAt: new Date() } },
    { upsert: true }
  );

  return { ...payload, source: "api" };
}

export async function queryPickingRankings({ warehouse, warehouseNo = DEFAULT_WAREHOUSE_NO, now = new Date(), forceRefresh = false, period = "" } = {}) {
  const weekly = previousWeekRange(now);
  const monthly = previousMonthRange(now);
  if (period === "week") {
    return {
      week: await getOrCreatePickingRanking({ period: "week", from: weekly.startDate, to: weekly.endDate, warehouse, warehouseNo, forceRefresh })
    };
  }
  if (period === "month") {
    return {
      month: await getOrCreatePickingRanking({ period: "month", from: monthly.startDate, to: monthly.endDate, warehouse, warehouseNo, forceRefresh })
    };
  }
  const [week, month] = await Promise.all([
    getOrCreatePickingRanking({ period: "week", from: weekly.startDate, to: weekly.endDate, warehouse, warehouseNo, forceRefresh }),
    getOrCreatePickingRanking({ period: "month", from: monthly.startDate, to: monthly.endDate, warehouse, warehouseNo, forceRefresh })
  ]);
  return { week, month };
}

export async function excludePickingRankingPerson({ employeeNo, name } = {}) {
  const cleanEmployeeNo = identityText(employeeNo);
  const cleanName = identityText(name);
  if (!cleanEmployeeNo && !cleanName) {
    const error = new Error("Missing employee identifier.");
    error.status = 400;
    throw error;
  }

  const exclusions = await pickingRankingExclusionCollection();
  const now = new Date();
  const filter = cleanEmployeeNo ? { employeeNo: cleanEmployeeNo } : { name: cleanName };
  await exclusions.updateOne(
    filter,
    {
      $set: {
        employeeNo: cleanEmployeeNo,
        name: cleanName,
        updatedAt: now
      },
      $setOnInsert: { createdAt: now }
    },
    { upsert: true }
  );

  const rankingCache = await pickingRankingCacheCollection();
  const docs = await rankingCache.find({
    $or: [
      cleanEmployeeNo ? { "rows.employeeNo": cleanEmployeeNo } : null,
      cleanName ? { "rows.name": cleanName } : null
    ].filter(Boolean)
  }).toArray();
  const excluded = await readRankingExclusions();

  for (const doc of docs) {
    const rows = rerankRows((doc.rows || []).filter((row) => !isExcludedRankingRow(row, excluded)));
    await rankingCache.updateOne(
      { _id: doc._id },
      { $set: { rows, updatedAt: now } }
    );
  }

  return { ok: true, removedFromCachedRankings: docs.length };
}

export async function mergePickingRankingEmployee({
  warehouse,
  period,
  startDate,
  endDate,
  sourceEmployeeNo,
  sourceName,
  targetEmployeeNo
} = {}) {
  const warehouseKey = String(warehouse || "5").trim();
  const cleanPeriod = identityText(period);
  const targetKey = employeeKey(targetEmployeeNo);
  const sourceKey = employeeKey(sourceEmployeeNo);
  const sourceNameKey = identityText(sourceName).toLowerCase();

  if (!cleanPeriod || !startDate || !endDate || !sourceKey || !targetKey) {
    const error = new Error("Missing ranking merge fields.");
    error.status = 400;
    throw error;
  }

  const collection = await pickingRankingCacheCollection();
  const filter = { warehouseKey, period: cleanPeriod, startDate, endDate };
  const cached = await collection.findOne(filter, { projection: { _id: 0 } });
  if (!cached?.rows?.length) return { changed: false, reason: "ranking-not-found", ranking: null };

  const rows = [...cached.rows];
  const sourceIndex = rows.findIndex((row) => (
    employeeKey(row.employeeNo) === sourceKey
    || (!row.employeeNo && sourceNameKey && identityText(row.name).toLowerCase() === sourceNameKey)
  ));
  const targetIndex = rows.findIndex((row) => employeeKey(row.employeeNo) === targetKey);

  if (sourceIndex < 0) return { changed: false, reason: "source-not-found", ranking: await filterRankingDoc(cached) };
  if (targetIndex < 0 || sourceIndex === targetIndex) {
    return { changed: false, reason: targetIndex < 0 ? "target-not-found" : "same-row", ranking: await filterRankingDoc(cached) };
  }

  const targetRow = rows[targetIndex];
  const sourceRow = rows[sourceIndex];
  const nextRows = rows.filter((_, index) => index !== sourceIndex);
  const adjustedTargetIndex = sourceIndex < targetIndex ? targetIndex - 1 : targetIndex;
  nextRows[adjustedTargetIndex] = mergeRankingRows(targetRow, sourceRow);
  const rankedRows = rerankRows(nextRows);
  const nextDoc = { ...cached, rows: rankedRows, updatedAt: new Date() };

  await collection.updateOne(filter, { $set: { rows: rankedRows, updatedAt: nextDoc.updatedAt } });
  return { changed: true, ranking: { ...(await filterRankingDoc(nextDoc)), source: "database" } };
}

export async function updatePickingRankingEmployeeName({ employeeNo, oldName, name } = {}) {
  const cleanEmployeeNo = identityText(employeeNo);
  const employeeNoKey = employeeKey(cleanEmployeeNo);
  const cleanOldName = identityText(oldName);
  const cleanName = identityText(name);

  if (!cleanName || (!employeeNoKey && !cleanOldName)) {
    const error = new Error("Missing employee identifier or name.");
    error.status = 400;
    throw error;
  }

  const collection = await pickingRankingCacheCollection();
  const docs = await collection.find({
    $or: [
      cleanEmployeeNo ? { "rows.employeeNo": cleanEmployeeNo } : null,
      cleanOldName ? { "rows.name": cleanOldName } : null
    ].filter(Boolean)
  }).toArray();

  const now = new Date();
  let updatedDocs = 0;
  let updatedRows = 0;

  for (const doc of docs) {
    let changed = false;
    const rows = (doc.rows || []).map((row) => {
      const matchesEmployee = employeeNoKey && employeeKey(row.employeeNo) === employeeNoKey;
      const matchesName = cleanOldName && identityText(row.name) === cleanOldName;
      if (!matchesEmployee && !matchesName) return row;
      changed = true;
      updatedRows += 1;
      return { ...row, name: cleanName };
    });

    if (!changed) continue;
    updatedDocs += 1;
    await collection.updateOne(
      { _id: doc._id },
      { $set: { rows, updatedAt: now } }
    );
  }

  return { ok: true, updatedDocs, updatedRows, name: cleanName };
}

async function claimRankingJob(jobKey) {
  const collection = await pickingRankingJobCollection();
  try {
    await collection.insertOne({ jobKey, createdAt: new Date() });
    return true;
  } catch (error) {
    if (error?.code === 11000) return false;
    throw error;
  }
}

async function runScheduledRankingJob(now = new Date()) {
  const day = now.getDay();
  const date = now.getDate();
  const tasks = [];

  if (day === 1) {
    const range = previousWeekRange(now);
    for (const warehouse of DEFAULT_RANKING_WAREHOUSES) {
      tasks.push({ period: "week", warehouse, ...range });
    }
  }

  if (date === 1) {
    const range = previousMonthRange(now);
    for (const warehouse of DEFAULT_RANKING_WAREHOUSES) {
      tasks.push({ period: "month", warehouse, ...range });
    }
  }

  for (const task of tasks) {
    const jobKey = `${task.period}|${task.warehouse}|${task.startDate}|${task.endDate}`;
    if (!(await claimRankingJob(jobKey))) continue;
    await getOrCreatePickingRanking({
      period: task.period,
      from: task.startDate,
      to: task.endDate,
      warehouse: task.warehouse,
      forceRefresh: true
    });
    console.log("Picking ranking cached", jobKey);
  }
}

export function startPickingRankingScheduler() {
  if (rankingSchedulerTimer) return;

  const run = async () => {
    try {
      await runScheduledRankingJob();
    } catch (error) {
      console.warn("Picking ranking scheduler skipped:", error.message);
    }
  };

  setTimeout(run, 15_000).unref?.();
  rankingSchedulerTimer = setInterval(run, HOUR_MS);
  rankingSchedulerTimer.unref?.();
}
