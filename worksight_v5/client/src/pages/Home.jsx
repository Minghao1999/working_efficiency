import React, { useEffect, useRef, useState } from "react";
import { AlertTriangle, CalendarDays, FileSpreadsheet, Medal, MessageSquarePlus, Package, Trophy, Users, X } from "lucide-react";
import { API } from "../constants";
import { ConfirmDialog, GlassSelect, ProgressBar } from "../components/controls";
import { formatCell, formatUploadError } from "../utils/formatters";

const WAREHOUSE_OPTIONS = [
  { value: "1", label: "Warehouse 1" },
  { value: "2", label: "Warehouse 2" },
  { value: "5", label: "Warehouse 5" }
];

const HIDE_UNFINISHED_RANKINGS = import.meta.env.VITE_HIDE_UNFINISHED_RANKINGS
  ? import.meta.env.VITE_HIDE_UNFINISHED_RANKINGS === "true"
  : import.meta.env.PROD;
const UNFINISHED_RANKING_WAREHOUSES = new Set(["2"]);

async function readApiJson(response) {
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) return response.json();
  throw new Error(`API returned ${response.status || "non-JSON"} instead of JSON.`);
}

function cloneRankingData(data) {
  return data ? JSON.parse(JSON.stringify(data)) : data;
}

function normalizeEmployeeId(value) {
  const text = String(value || "").trim();
  return (text.includes("@") ? text.split("@")[0] : text).toUpperCase();
}

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function recalculateRankingRow(row) {
  const pickingWeightedUnits = toNumber(row.pickingWeightedUnits ?? row.weightedUnits);
  const pickingHours = toNumber(row.pickingHours ?? row.hours ?? row.effectiveHours);
  const packingUnits = toNumber(row.packingUnits);
  const packingHours = toNumber(row.packingHours);
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
    efficiency: totalHours ? (pickingWeightedUnits + packingUnits) / totalHours : 0
  };
}

function mergeRankingRows(target, source) {
  const targetPickingUnits = toNumber(target.pickingWeightedUnits ?? target.weightedUnits);
  const sourcePickingUnits = toNumber(source.pickingWeightedUnits ?? source.weightedUnits);
  const merged = {
    ...target,
    pickingWeightedUnits: targetPickingUnits + sourcePickingUnits,
    weightedUnits: targetPickingUnits + sourcePickingUnits,
    pickingHours: toNumber(target.pickingHours ?? target.hours ?? target.effectiveHours) + toNumber(source.pickingHours ?? source.hours ?? source.effectiveHours),
    hours: toNumber(target.pickingHours ?? target.hours ?? target.effectiveHours) + toNumber(source.pickingHours ?? source.hours ?? source.effectiveHours),
    effectiveHours: toNumber(target.pickingHours ?? target.hours ?? target.effectiveHours) + toNumber(source.pickingHours ?? source.hours ?? source.effectiveHours),
    packingUnits: toNumber(target.packingUnits) + toNumber(source.packingUnits),
    packingHours: toNumber(target.packingHours) + toNumber(source.packingHours)
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

function rerankRows(rows) {
  return [...rows]
    .sort((a, b) => toNumber(b.efficiency) - toNumber(a.efficiency) || toNumber(b.pickingWeightedUnits ?? b.weightedUnits) - toNumber(a.pickingWeightedUnits ?? a.weightedUnits))
    .map((row, index) => ({ ...row, rank: index + 1 }));
}

function mergeEmployeeInRanking(ranking, sourceRow, targetEmployeeNo) {
  if (!ranking?.rows?.length) return ranking;
  const rows = [...ranking.rows];
  const sourceKey = normalizeEmployeeId(sourceRow.employeeNo);
  const sourceName = String(sourceRow.name || "").trim().toLowerCase();
  const targetKey = normalizeEmployeeId(targetEmployeeNo);
  const sourceIndex = rows.findIndex((row) => (
    normalizeEmployeeId(row.employeeNo) === sourceKey
    || (!row.employeeNo && sourceName && String(row.name || "").trim().toLowerCase() === sourceName)
  ));
  const targetIndex = rows.findIndex((row) => normalizeEmployeeId(row.employeeNo) === targetKey);
  if (sourceIndex < 0 || targetIndex < 0 || sourceIndex === targetIndex) return ranking;
  const nextRows = rows.filter((_, index) => index !== sourceIndex);
  const adjustedTargetIndex = sourceIndex < targetIndex ? targetIndex - 1 : targetIndex;
  nextRows[adjustedTargetIndex] = mergeRankingRows(rows[targetIndex], rows[sourceIndex]);
  return { ...ranking, rows: rerankRows(nextRows) };
}

function removePersonFromRankingData(data, person) {
  if (!data || !person) return data;
  const employeeKey = normalizeEmployeeId(person.employeeNo);
  const nameKey = String(person.name || "").trim().toLowerCase();
  const removeFromRanking = (ranking) => {
    if (!ranking?.rows) return ranking;
    return {
      ...ranking,
      rows: rerankRows(ranking.rows.filter((row) => {
        const sameEmployee = employeeKey && normalizeEmployeeId(row.employeeNo) === employeeKey;
        const sameName = nameKey && String(row.name || "").trim().toLowerCase() === nameKey;
        return !sameEmployee && !sameName;
      }))
    };
  };
  return {
    ...data,
    week: removeFromRanking(data.week),
    month: removeFromRanking(data.month)
  };
}

function updatePersonNameInRankingData(data, person, name) {
  if (!data || !person) return data;
  const employeeKey = normalizeEmployeeId(person.employeeNo);
  const oldNameKey = String(person.name || "").trim().toLowerCase();
  const updateRanking = (ranking) => {
    if (!ranking?.rows) return ranking;
    return {
      ...ranking,
      rows: ranking.rows.map((row) => {
        const sameEmployee = employeeKey && normalizeEmployeeId(row.employeeNo) === employeeKey;
        const sameName = oldNameKey && String(row.name || "").trim().toLowerCase() === oldNameKey;
        return sameEmployee || sameName ? { ...row, name } : row;
      })
    };
  };
  return {
    ...data,
    week: updateRanking(data.week),
    month: updateRanking(data.month)
  };
}

export function Home({ onNavigate }) {
  const [warehouse, setWarehouse] = useState("5");
  const [rankingPeriod, setRankingPeriod] = useState("week");
  const [rankingData, setRankingData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [pendingDelete, setPendingDelete] = useState(null);
  const [pendingNameUpdate, setPendingNameUpdate] = useState(null);
  const [pendingEmployeeMerge, setPendingEmployeeMerge] = useState(null);
  const [error, setError] = useState("");
  const originalRankingDataRef = useRef(null);
  const rankingInProgress = HIDE_UNFINISHED_RANKINGS && UNFINISHED_RANKING_WAREHOUSES.has(warehouse);

  useEffect(() => {
    const controller = new AbortController();
    async function loadRanking() {
      setRankingData(null);
      originalRankingDataRef.current = null;
      setError("");
      if (rankingInProgress) {
        setLoading(false);
        return;
      }
      setLoading(true);
      try {
        const weekResponse = await fetch(`${API}/api/weekly/picking-rankings`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ warehouse, period: "week" }),
          signal: controller.signal
        });
        const weekJson = await readApiJson(weekResponse);
        if (!weekResponse.ok) throw weekJson;
        setRankingData(weekJson);
        originalRankingDataRef.current = cloneRankingData(weekJson);
        if (!controller.signal.aborted) setLoading(false);

        const monthResponse = await fetch(`${API}/api/weekly/picking-rankings`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ warehouse, period: "month" }),
          signal: controller.signal
        });
        const monthJson = await readApiJson(monthResponse);
        if (!monthResponse.ok) throw monthJson;
        setRankingData((current) => {
          const next = { ...(current || {}), ...monthJson };
          originalRankingDataRef.current = cloneRankingData(next);
          return next;
        });
      } catch (e) {
        if (e.name === "AbortError") return;
        const message = String(e?.message || "");
        setError(message === "Failed to fetch"
          ? "API is not reachable. Check that the backend is running on http://127.0.0.1:3001."
          : formatUploadError(e));
      } finally {
        if (!controller.signal.aborted) setLoading(false);
      }
    }
    loadRanking();
    return () => controller.abort();
  }, [warehouse, rankingInProgress]);

  function refreshRanking() {
    setError("");
    if (rankingInProgress) {
      setRankingData(null);
      setLoading(false);
      return;
    }
    setRankingData(cloneRankingData(originalRankingDataRef.current));
    setLoading(false);
  }

  async function confirmDeletePerson() {
    if (!pendingDelete) return;
    setDeleting(true);
    setError("");
    try {
      const response = await fetch(`${API}/api/weekly/picking-rankings/exclude`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          employeeNo: pendingDelete.employeeNo,
          name: pendingDelete.name
        })
      });
      const json = await readApiJson(response);
      if (!response.ok) throw json;
      setRankingData((current) => removePersonFromRankingData(current, pendingDelete));
      originalRankingDataRef.current = removePersonFromRankingData(originalRankingDataRef.current, pendingDelete);
      setPendingDelete(null);
    } catch (e) {
      setError(formatUploadError(e));
    } finally {
      setDeleting(false);
    }
  }

  function mergeRankingEmployee({ row, targetEmployeeNo, period, ranking }) {
    const cleanTarget = String(targetEmployeeNo || "").trim();
    if (!cleanTarget || cleanTarget === row.employeeNo) return false;
    setError("");
    if (mergeEmployeeInRanking(ranking, row, cleanTarget) === ranking) return false;
    setPendingEmployeeMerge({ row, targetEmployeeNo: cleanTarget, period, ranking });
    return true;
  }

  function confirmEmployeeMerge() {
    if (!pendingEmployeeMerge) return;
    const { row, targetEmployeeNo, period, ranking } = pendingEmployeeMerge;
    const mergedRanking = mergeEmployeeInRanking(ranking, row, targetEmployeeNo);
    if (mergedRanking === ranking) {
      setPendingEmployeeMerge(null);
      return;
    }
    setRankingData((current) => ({
      ...(current || {}),
      [period]: mergedRanking
    }));
    setPendingEmployeeMerge(null);
  }

  function updateRankingEmployeeName({ row, name }) {
    const cleanName = String(name || "").trim();
    if (!cleanName || cleanName === row.name) return false;
    setPendingNameUpdate({ row, name: cleanName });
    return true;
  }

  async function confirmNameUpdate() {
    if (!pendingNameUpdate) return;
    const { row, name } = pendingNameUpdate;
    setError("");
    try {
      const response = await fetch(`${API}/api/weekly/picking-rankings/update-name`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          employeeNo: row.employeeNo,
          oldName: row.name,
          name
        })
      });
      const json = await readApiJson(response);
      if (!response.ok) throw json;
      setRankingData((current) => updatePersonNameInRankingData(current, row, name));
      originalRankingDataRef.current = updatePersonNameInRankingData(originalRankingDataRef.current, row, name);
      setPendingNameUpdate(null);
    } catch (e) {
      setError(formatUploadError(e));
    }
  }

  return (
    <section className="page">
      <header className="page-head">
        <h1>WorkSight Pro</h1>
        <p>Warehouse operations analytics for productivity, orders, units, and trends</p>
      </header>
      <section className="panel ranking-panel">
        <div className="table-head">
          <div>
            <div className="panel-title">
              <Trophy size={20} />
              <span>Picking Efficiency Ranking</span>
            </div>
            <p className="hint-line">Goods type factor: T25 / T50 = 1 small unit, all other types = 2.5 small units. Floor factor: L1 = 1, L2 = 1.1, L3 = 1.2, L4 = 1.3.</p>
            <p className="hint-line">Picking Weighted Units = Non-Burst Units * Goods type factor* Floor factor</p>
            <div className="ranking-period-toggle" role="group" aria-label="Ranking period">
              <button
                type="button"
                className={rankingPeriod === "week" ? "active" : ""}
                onClick={() => setRankingPeriod("week")}
              >
                Weekly
              </button>
              <button
                type="button"
                className={rankingPeriod === "month" ? "active" : ""}
                onClick={() => setRankingPeriod("month")}
              >
                Monthly
              </button>
            </div>
          </div>
          <div className="overview-ranking-tools">
            <label className="date-query-warehouse">
              <span>Warehouse</span>
              <GlassSelect value={warehouse} options={WAREHOUSE_OPTIONS} onChange={setWarehouse} className="warehouse-query-select" />
            </label>
            <button type="button" className="ghost-btn" onClick={refreshRanking} disabled={loading || rankingInProgress}>
              {loading ? "Loading..." : "Refresh"}
            </button>
          </div>
        </div>
        <div className="ranking-range">
          <CalendarDays size={16} />
          <span>Week: {rankingData?.week?.startDate || "..."} to {rankingData?.week?.endDate || "..."}</span>
          <span>Month: {rankingData?.month?.startDate || "..."} to {rankingData?.month?.endDate || "..."}</span>
        </div>
        {loading && <ProgressBar value={68} label="Loading cached picking rankings..." />}
        {error && <div className="error">{error}</div>}
        {rankingInProgress ? (
          <RankingComingSoon title={rankingPeriod === "week" ? "Weekly Ranking" : "Monthly Ranking"} />
        ) : (
          <RankingTable
            title={rankingPeriod === "week" ? "Weekly Ranking" : "Monthly Ranking"}
            ranking={rankingPeriod === "week" ? rankingData?.week : rankingData?.month}
            period={rankingPeriod === "week" ? "week" : "month"}
            onDelete={setPendingDelete}
            onMergeEmployee={mergeRankingEmployee}
            onUpdateName={updateRankingEmployeeName}
            disabled={deleting || loading}
          />
        )}
      </section>
      <div className="landing-grid">
        <button className="feature-tile" onClick={() => onNavigate("efficiency")}>
          <Users size={32} />
          <h2>Efficiency Dashboard</h2>
          <p>Review employee productivity, Gantt timelines, and Idle / Work distribution</p>
        </button>
        <button className="feature-tile" onClick={() => onNavigate("weekly")}>
          <Package size={32} />
          <h2>Order / Unit Analysis</h2>
          <p>Summarize daily order and unit trends with Excel export</p>
        </button>
        <button className="feature-tile" onClick={() => onNavigate("miniApps")}>
          <FileSpreadsheet size={32} />
          <h2>Mini Programs</h2>
          <p>Open small tools for warehouse operations</p>
        </button>
        <button className="feature-tile" onClick={() => onNavigate("feedback")}>
          <MessageSquarePlus size={32} />
          <h2>Feedback</h2>
          <p>Share feature ideas and change requests for WorkSight</p>
        </button>
      </div>
      {pendingDelete && (
        <ConfirmDialog
          title="Delete from ranking?"
          message={`Delete ${pendingDelete.name || pendingDelete.employeeNo} from picking rankings? This person will not appear in future weekly or monthly rankings.`}
          onCancel={() => setPendingDelete(null)}
          onConfirm={confirmDeletePerson}
        />
      )}
      {pendingNameUpdate && (
        <ConfirmDialog
          title="Update name?"
          message={`Change ${pendingNameUpdate.row.employeeNo || pendingNameUpdate.row.name} from "${pendingNameUpdate.row.name || ""}" to "${pendingNameUpdate.name}"?`}
          confirmLabel="Confirm"
          confirmClassName="primary-btn"
          onCancel={() => setPendingNameUpdate(null)}
          onConfirm={confirmNameUpdate}
        />
      )}
      {pendingEmployeeMerge && (
        <ConfirmDialog
          title="Merge employee ID?"
          message={`Merge ${pendingEmployeeMerge.row.employeeNo || pendingEmployeeMerge.row.name} into ${pendingEmployeeMerge.targetEmployeeNo}?`}
          confirmLabel="Confirm"
          confirmClassName="primary-btn"
          onCancel={() => setPendingEmployeeMerge(null)}
          onConfirm={confirmEmployeeMerge}
        />
      )}
    </section>
  );
}

function RankingComingSoon({ title }) {
  return (
    <div className="ranking-card ranking-coming-soon">
      <div className="ranking-card-head">
        <h2>{title}</h2>
      </div>
      <div className="ranking-coming-soon-body">
        <Trophy size={26} />
        <strong>还在制作中，马上</strong>
        <span>Warehouse 5 ranking is available now.</span>
      </div>
    </div>
  );
}

function RankingTable({ title, ranking, period, onDelete, onMergeEmployee, onUpdateName, disabled }) {
  const rows = ranking?.rows || [];
  const first = rows[0];
  const last = rows.length > 1 ? rows[rows.length - 1] : null;

  return (
    <div className="ranking-card">
      <div className="ranking-card-head">
        <h2>{title}</h2>
        {!!rows.length && <span>{rows.length} people - {ranking?.source || "database"}</span>}
      </div>
      {first && (
        <div className="ranking-callouts">
          <RankingCallout icon={<Medal size={16} />} label="First" row={first} />
          {last && <RankingCallout icon={<AlertTriangle size={16} />} label="Last" row={last} muted />}
        </div>
      )}
      {rows.length ? (
        <div className="table-wrap compact-table">
          <table>
            <thead>
              <tr>
                <th className="ranking-delete-col"></th>
                <th>Rank</th>
                <th>Employee ID</th>
                <th>Name</th>
                <th>Picking Weighted Units</th>
                <th>Packing Units</th>
                <th>Picking Hours</th>
                <th>Packing Hours</th>
                <th>Efficiency</th>
                <th>L1%</th>
                <th>L2%</th>
                <th>L3%</th>
                <th>L4%</th>
                <th>Main Cargo</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row, index) => (
                <tr key={`${row.employeeNo || row.name}-${index}`} className={index === 0 ? "rank-first" : index === rows.length - 1 ? "rank-last" : ""}>
                  <td className="ranking-delete-col">
                    <button
                      type="button"
                      className="ranking-delete-btn"
                      aria-label={`Delete ${row.name || row.employeeNo} from rankings`}
                      title="Delete from rankings"
                      disabled={disabled}
                      onClick={() => onDelete(row)}
                    >
                      <X size={14} />
                    </button>
                  </td>
                  <td>{row.rank}</td>
                  <td>
                    <EmployeeIdInput
                      row={row}
                      ranking={ranking}
                      period={period}
                      disabled={disabled}
                      onMergeEmployee={onMergeEmployee}
                    />
                  </td>
                  <td>
                    <NameInput
                      row={row}
                      disabled={disabled}
                      onUpdateName={onUpdateName}
                    />
                  </td>
                  <td>{formatCell(row.pickingWeightedUnits ?? row.weightedUnits)}</td>
                  <td>{formatCell(row.packingUnits || 0)}</td>
                  <td>{formatCell(row.pickingHours ?? row.hours ?? row.effectiveHours)}</td>
                  <td>{formatCell(row.packingHours || 0)}</td>
                  <td>{formatCell(row.efficiency)}</td>
                  <td>{formatPercent(row.l1Percent)}</td>
                  <td>{formatPercent(row.l2Percent)}</td>
                  <td>{formatPercent(row.l3Percent)}</td>
                  <td>{formatPercent(row.l4Percent)}</td>
                  <td>{formatCargoMix(row)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <div className="empty">No cached ranking for this period yet.</div>
      )}
    </div>
  );
}

function formatPercent(value) {
  const n = Number(value || 0);
  return `${n.toFixed(1)}%`;
}

function formatCargoMix(row) {
  return `Large ${formatPercent(row.largePercent)} / Small ${formatPercent(row.smallPercent)}`;
}

function EmployeeIdInput({ row, ranking, period, disabled, onMergeEmployee }) {
  const [value, setValue] = useState(row.employeeNo || "");

  useEffect(() => {
    setValue(row.employeeNo || "");
  }, [row.employeeNo]);

  function commit() {
    const cleanValue = value.trim();
    if (!cleanValue || cleanValue === row.employeeNo) {
      setValue(row.employeeNo || "");
      return;
    }
    const pending = onMergeEmployee?.({ row, targetEmployeeNo: cleanValue, period, ranking });
    if (pending) setValue(row.employeeNo || "");
  }

  return (
    <input
      className="ranking-employee-input"
      value={value}
      disabled={disabled}
      onChange={(event) => setValue(event.target.value)}
      onBlur={commit}
      onKeyDown={(event) => {
        if (event.key === "Enter") event.currentTarget.blur();
        if (event.key === "Escape") {
          setValue(row.employeeNo || "");
          event.currentTarget.blur();
        }
      }}
      aria-label={`Employee ID for ${row.name || row.employeeNo}`}
    />
  );
}

function NameInput({ row, disabled, onUpdateName }) {
  const [value, setValue] = useState(row.name || "");

  useEffect(() => {
    setValue(row.name || "");
  }, [row.name]);

  function commit() {
    const cleanValue = value.trim();
    if (!cleanValue || cleanValue === row.name) {
      setValue(row.name || "");
      return;
    }
    const pending = onUpdateName?.({ row, name: cleanValue });
    if (pending) setValue(row.name || "");
  }

  return (
    <input
      className="ranking-name-input"
      value={value}
      disabled={disabled}
      onChange={(event) => setValue(event.target.value)}
      onBlur={commit}
      onKeyDown={(event) => {
        if (event.key === "Enter") event.currentTarget.blur();
        if (event.key === "Escape") {
          setValue(row.name || "");
          event.currentTarget.blur();
        }
      }}
      aria-label={`Name for ${row.employeeNo || row.name}`}
    />
  );
}

function RankingCallout({ icon, label, row, muted = false }) {
  return (
    <div className={muted ? "ranking-callout muted" : "ranking-callout"}>
      <span>{icon}{label}</span>
      <strong>{row.name || row.employeeNo}</strong>
      <small>{formatCell(row.efficiency)} units/hour</small>
    </div>
  );
}
