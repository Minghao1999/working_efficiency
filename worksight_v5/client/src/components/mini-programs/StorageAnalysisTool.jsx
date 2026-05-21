import React, { useState } from "react";
import { ArrowLeft, Boxes, Download, Loader2, TableProperties, TriangleAlert, Warehouse } from "lucide-react";
import * as XLSX from "xlsx";
import { FilePicker } from "../controls";

const SLOT_CAPACITY = 120;
const ITEM_LENGTH = 40;
const LEVELS = [1, 2, 3, 4, 5];
const BIN_NUMBERS = [1, 2, 3];

const AISLE_RACK_RANGES = new Map([
  [1, [13, 70]],
  [2, [7, 65]],
  [3, [7, 69]],
  [4, [3, 70]],
  [5, [3, 70]],
  [6, [4, 70]],
  [7, [3, 70]],
  [8, [3, 70]],
  [9, [3, 70]],
  [10, [3, 70]],
  [11, [3, 70]],
  [12, [4, 70]],
  [13, [3, 70]],
  [14, [3, 70]],
  [15, [3, 70]],
  [16, [3, 70]],
  [17, [3, 70]],
  [18, [3, 69]],
  [19, [3, 70]],
  [20, [3, 70]],
  [21, [4, 70]],
  [22, [3, 70]],
  [23, [3, 70]],
  [24, [4, 70]]
]);

function normalizeHeader(value) {
  return String(value ?? "").replace(/\s+/g, "").toLowerCase();
}

function findColumn(headers, aliases, fallbackIndex = -1) {
  const normalizedAliases = aliases.map(normalizeHeader);
  const index = headers.findIndex((header) => normalizedAliases.includes(normalizeHeader(header)));
  if (index >= 0) return headers[index];
  if (fallbackIndex >= 0 && fallbackIndex < headers.length) return headers[fallbackIndex];
  return "";
}

function extractNumber(value, prefix) {
  const match = String(value ?? "").match(new RegExp(`${prefix}(\\d+)`, "i"));
  return match ? Number(match[1]) : NaN;
}

function removeBin(location) {
  return String(location ?? "").replace(/-B\d+/i, "");
}

function isInRackRange(aNum, rNum) {
  const range = AISLE_RACK_RANGES.get(aNum);
  if (!range || !Number.isFinite(rNum)) return false;
  return rNum >= range[0] && rNum <= range[1];
}

function isValidLevel(levelNum) {
  return LEVELS.includes(levelNum);
}

function isValidBin(binNum) {
  return BIN_NUMBERS.includes(binNum);
}

async function readFirstSheet(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) return [];
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
}

function normalizeStorageRows(rows, locationCol, lengthCol = "") {
  return rows.map((row) => {
    const location = String(row[locationCol] ?? "");
    const aNum = extractNumber(location, "A");
    const rNum = extractNumber(location, "R");
    const levelNum = extractNumber(location, "L");
    const binNum = extractNumber(location, "B");
    return {
      row,
      location,
      baseSlot: removeBin(location),
      aNum,
      rNum,
      levelNum,
      binNum,
      length: lengthCol ? Number(row[lengthCol]) : 0,
      inScope: isInRackRange(aNum, rNum) && isValidLevel(levelNum)
    };
  });
}

function generateBaseSlots() {
  const slots = [];
  for (const [aNum, [rackStart, rackEnd]] of AISLE_RACK_RANGES.entries()) {
    for (let rNum = rackStart; rNum <= rackEnd; rNum += 1) {
      for (const levelNum of LEVELS) {
        slots.push({
          slot: `A${aNum}-R${rNum}-L${levelNum}`,
          aNum,
          rNum,
          levelNum,
          level: `L${levelNum}`
        });
      }
    }
  }
  return slots;
}

function summarizeLevel(slots, levelNum) {
  const rows = slots.filter((slot) => slot.levelNum === levelNum);
  return {
    level: `L${levelNum}`,
    levelSort: levelNum,
    totalSlots: rows.reduce((sum, slot) => sum + slot.totalBinSlots, 0),
    usedSlots: rows.reduce((sum, slot) => sum + slot.usedBinSlots, 0),
    emptySlots: rows.reduce((sum, slot) => sum + slot.emptyBinSlots, 0),
    totalItems: rows.reduce((sum, slot) => sum + slot.canPutItems, 0),
    oneItem: rows.filter((slot) => slot.canPutItems === 1).length,
    twoItems: rows.filter((slot) => slot.canPutItems === 2).length,
    consecutiveTwoEmpty: rows.filter((slot) => slot.hasConsecutiveTwoEmptyBinsAndCanPutTwo).length,
    threeItems: rows.filter((slot) => slot.canPutItems === 3).length,
    baseSlots: rows.length
  };
}

function analyzeStorage(usedRows, emptyRows) {
  const usedHeaders = Object.keys(usedRows[0] || {});
  const emptyHeaders = Object.keys(emptyRows[0] || {});
  const usedLocationCol = findColumn(usedHeaders, ["储位编码", "storage cell", "location", "cell"], 15);
  const emptyLocationCol = findColumn(emptyHeaders, ["储位编码", "storage cell", "location", "cell"], 0);
  const lengthCol = findColumn(usedHeaders, ["长", "length", "长度"], -1);

  if (!usedLocationCol) throw new Error("非空储位表找不到储位编码列。");
  if (!emptyLocationCol) throw new Error("空储位表找不到储位编码列。");
  if (!lengthCol) throw new Error("非空储位表找不到长度列（长 / length）。");

  const normalizedUsed = normalizeStorageRows(usedRows, usedLocationCol, lengthCol);
  const normalizedEmpty = normalizeStorageRows(emptyRows, emptyLocationCol);
  const usedFiltered = normalizedUsed.filter((entry) => entry.inScope);
  const emptyFiltered = normalizedEmpty.filter((entry) => entry.inScope);
  const usedByBase = new Map();
  const emptyBinKeys = new Set();

  for (const entry of usedFiltered) {
    if (!entry.baseSlot) continue;
    const current = usedByBase.get(entry.baseSlot) || {
      length: 0,
      occupiedBins: new Set()
    };
    current.length += Number.isFinite(entry.length) ? entry.length : 0;
    if (isValidBin(entry.binNum)) current.occupiedBins.add(entry.binNum);
    usedByBase.set(entry.baseSlot, current);
  }

  for (const entry of emptyFiltered) {
    if (!entry.baseSlot || !isValidBin(entry.binNum)) continue;
    emptyBinKeys.add(`${entry.baseSlot}-B${entry.binNum}`);
  }

  const slots = generateBaseSlots()
    .map((entry) => {
      const used = usedByBase.get(entry.slot);
      const usedBinSlots = used?.occupiedBins.size || 0;
      const occupiedBins = used?.occupiedBins || new Set();
      const emptyBins = BIN_NUMBERS.filter((binNum) => !occupiedBins.has(binNum));
      const uploadedEmptyBins = BIN_NUMBERS.filter((binNum) => emptyBinKeys.has(`${entry.slot}-B${binNum}`));
      const emptyBinSlots = Math.max(0, BIN_NUMBERS.length - usedBinSlots);
      const usedLength = used?.length || 0;
      const remainSpace = SLOT_CAPACITY - usedLength;
      const canPutItems = Math.max(0, Math.min(BIN_NUMBERS.length, Math.floor(remainSpace / ITEM_LENGTH)));
      const hasUploadedThreeConnected = [1, 2, 3].every((binNum) => uploadedEmptyBins.includes(binNum));
      const hasUploadedTwoConnected = !hasUploadedThreeConnected && (
        (uploadedEmptyBins.includes(1) && uploadedEmptyBins.includes(2)) ||
        (uploadedEmptyBins.includes(2) && uploadedEmptyBins.includes(3))
      );
      return {
        ...entry,
        totalBinSlots: BIN_NUMBERS.length,
        usedBinSlots,
        emptyBinSlots,
        hasConsecutiveTwoEmptyBinsAndCanPutTwo: hasUploadedTwoConnected && canPutItems === 2,
        uploadedEmptyBinSlots: uploadedEmptyBins.length,
        usedLength,
        remainSpace,
        canPutItems
      };
    })
    .sort((a, b) => (
      a.levelNum - b.levelNum ||
      a.aNum - b.aNum ||
      a.rNum - b.rNum ||
      a.slot.localeCompare(b.slot, undefined, { numeric: true, sensitivity: "base" })
    ));

  const levelStats = LEVELS.map((levelNum) => summarizeLevel(slots, levelNum));
  const summaryRows = slots.filter((slot) => slot.levelNum >= 1 && slot.levelNum <= 4);

  return {
    columns: { usedLocationCol, emptyLocationCol, lengthCol },
    counts: {
      usedRows: usedRows.length,
      emptyRows: emptyRows.length,
      usedFiltered: usedFiltered.length,
      emptyFiltered: emptyFiltered.length,
      usedExcluded: normalizedUsed.length - usedFiltered.length,
      emptyExcluded: normalizedEmpty.length - emptyFiltered.length,
      totalSlots: slots.reduce((sum, slot) => sum + slot.totalBinSlots, 0),
      totalItems: slots.reduce((sum, slot) => sum + slot.canPutItems, 0),
      summaryTotalSlots: summaryRows.reduce((sum, slot) => sum + slot.totalBinSlots, 0),
      summaryUsedSlots: summaryRows.reduce((sum, slot) => sum + slot.usedBinSlots, 0),
      summaryEmptySlots: summaryRows.reduce((sum, slot) => sum + slot.emptyBinSlots, 0),
      summaryTotalItems: summaryRows.reduce((sum, slot) => sum + slot.canPutItems, 0)
    },
    levelStats,
    slots
  };
}

function downloadStorageDetails(slots) {
  const rows = slots.map((row) => ({
    储位: row.slot,
    Aisle: row.aNum,
    Rack: row.rNum,
    层级: row.level,
    总B储位: row.totalBinSlots,
    已用B: row.usedBinSlots,
    空B: row.emptyBinSlots,
    已用长度: Number(row.usedLength.toFixed(2)),
    剩余空间: Number(row.remainSpace.toFixed(2)),
    还能放40inch件数: row.canPutItems
  }));
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "储位明细");
  XLSX.writeFile(workbook, "storage-analysis-details.xlsx");
}

export function StorageAnalysisTool({ onBack }) {
  const [usedFile, setUsedFile] = useState(null);
  const [emptyFile, setEmptyFile] = useState(null);
  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  async function runAnalysis() {
    setError("");
    setResult(null);
    if (!usedFile || !emptyFile) {
      setError("请先上传非空储位和空储位表。");
      return;
    }

    setLoading(true);
    try {
      const [usedRows, emptyRows] = await Promise.all([readFirstSheet(usedFile), readFirstSheet(emptyFile)]);
      if (!usedRows.length) throw new Error("非空储位表没有读取到数据。");
      if (!emptyRows.length) throw new Error("空储位表没有读取到数据。");
      setResult(analyzeStorage(usedRows, emptyRows));
    } catch (err) {
      setError(err.message || "储位分析失败。");
    } finally {
      setLoading(false);
    }
  }

  function clearAll() {
    setUsedFile(null);
    setEmptyFile(null);
    setResult(null);
    setError("");
  }

  return (
    <div className="panel storage-analysis-tool">
      <div className="tool-head">
        <button className="ghost-btn" onClick={onBack}>
          <ArrowLeft size={16} /> Back
        </button>
        <div className="panel-title">
          <Warehouse size={20} />
          <span>储位分析</span>
        </div>
      </div>

      <div className="storage-analysis-layout">
        <div className="automation-card storage-upload-card">
          <h2>上传原始表</h2>
          <p className="hint-line">按 A/R 有效范围过滤库位，B1-B3 按 3 个储位统计，空间按每个 A-R-L 的 120inch 计算。</p>

          <div className="storage-upload-grid">
            <div className="upload-box">
              <strong>非空储位表</strong>
              <span>来自iWMS库存查询。</span>
              <FilePicker
                accept=".xlsx,.xls"
                files={usedFile ? [usedFile] : []}
                onChange={(files) => setUsedFile(files[0] || null)}
              />
            </div>

            <div className="upload-box">
              <strong>空储位表</strong>
              <span>来自iWMS储位状态查询，储位状态选择空闲。</span>
              <FilePicker
                accept=".xlsx,.xls"
                files={emptyFile ? [emptyFile] : []}
                onChange={(files) => setEmptyFile(files[0] || null)}
              />
            </div>
          </div>

          <div className="button-row storage-actions">
            <button className="primary-btn" onClick={runAnalysis} disabled={loading || !usedFile || !emptyFile}>
              {loading ? <Loader2 className="spin-icon" size={18} /> : <TableProperties size={18} />}
              {loading ? "分析中" : "开始分析"}
            </button>
            <button className="ghost-btn" onClick={clearAll} disabled={loading}>
              Clear
            </button>
          </div>

          {error && (
            <div className="error">
              <TriangleAlert size={16} /> {error}
            </div>
          )}
        </div>

        <div className="storage-summary-card">
          {!result && (
            <div className="empty-state">
              <Boxes size={34} />
              <strong>等待分析</strong>
              <span>选择两张 Excel 后，储位空间统计会显示在这里。</span>
            </div>
          )}

          {result && (
            <div className="storage-summary-grid">
              <div className="metric">
                <span>总储位</span>
                <strong>{result.counts.summaryTotalSlots.toLocaleString()}</strong>
                <small>B1-B3 计数，排除 L5</small>
              </div>
              <div className="metric">
                <span>已用储位</span>
                <strong>{result.counts.summaryUsedSlots.toLocaleString()}</strong>
                <small>被非空记录占用的 B，排除 L5</small>
              </div>
              <div className="metric">
                <span>空储位</span>
                <strong>{result.counts.summaryEmptySlots.toLocaleString()}</strong>
                <small>未被占用的 B，排除 L5</small>
              </div>
              <div className="metric">
                <span>还能放货物</span>
                <strong>{result.counts.summaryTotalItems.toLocaleString()}</strong>
                <small>按剩余 40inch 计算，排除 L5</small>
              </div>
            </div>
          )}
        </div>
      </div>

      {result && (
        <>
          <div className="automation-card wide">
            <div className="table-head storage-table-head">
              <h2>每层空间统计</h2>
              <span>识别列：非空储位 {result.columns.usedLocationCol}，空储位 {result.columns.emptyLocationCol}，长度 {result.columns.lengthCol}</span>
            </div>
            <div className="table-wrap">
              <table className="compact-table">
                <thead>
                  <tr>
                    <th>层级</th>
                    <th>总储位</th>
                    <th>已用储位</th>
                    <th>空储位</th>
                    <th>可放 1 件的 A-R-L</th>
                    <th>可放 2 件的 A-R-L</th>
                    <th>连续 2 个空 B 且可放 2 件的 A-R-L</th>
                    <th>可放 3 件的 A-R-L</th>
                    <th>合计还能放货物</th>
                  </tr>
                </thead>
                <tbody>
                  {result.levelStats.map((row) => (
                    <tr key={row.level}>
                      <td>{row.level}</td>
                      <td>{row.totalSlots}</td>
                      <td>{row.usedSlots}</td>
                      <td>{row.emptySlots}</td>
                      <td>{row.oneItem}</td>
                      <td>{row.twoItems}</td>
                      <td>{row.consecutiveTwoEmpty}</td>
                      <td>{row.threeItems}</td>
                      <td>{row.totalItems}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="automation-card wide">
            <div className="table-head storage-table-head">
              <div>
                <h2>储位明细</h2>
                <span>共 {result.slots.length.toLocaleString()} 个 A-R-L</span>
              </div>
              <button className="ghost-btn" onClick={() => downloadStorageDetails(result.slots)}>
                <Download size={16} /> 下载明细
              </button>
            </div>
            <div className="table-wrap storage-detail-table">
              <table className="compact-table">
                <thead>
                  <tr>
                    <th>储位</th>
                    <th>Aisle</th>
                    <th>Rack</th>
                    <th>层级</th>
                    <th>总 B 储位</th>
                    <th>已用 B</th>
                    <th>空 B</th>
                    <th>已用长度</th>
                    <th>剩余空间</th>
                    <th>还能放 40inch 件数</th>
                  </tr>
                </thead>
                <tbody>
                  {result.slots.map((row) => (
                    <tr key={row.slot}>
                      <td>{row.slot}</td>
                      <td>{row.aNum}</td>
                      <td>{row.rNum}</td>
                      <td>{row.level}</td>
                      <td>{row.totalBinSlots}</td>
                      <td>{row.usedBinSlots}</td>
                      <td>{row.emptyBinSlots}</td>
                      <td>{Number(row.usedLength.toFixed(2))}</td>
                      <td>{Number(row.remainSpace.toFixed(2))}</td>
                      <td>{row.canPutItems}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  );
}
