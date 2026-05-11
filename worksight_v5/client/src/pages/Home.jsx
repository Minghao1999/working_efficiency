import React, { useEffect, useState } from "react";
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
const UNFINISHED_RANKING_WAREHOUSES = new Set(["1", "2"]);

async function readApiJson(response) {
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) return response.json();
  throw new Error(`API returned ${response.status || "non-JSON"} instead of JSON.`);
}

export function Home({ onNavigate }) {
  const [warehouse, setWarehouse] = useState("5");
  const [rankingPeriod, setRankingPeriod] = useState("week");
  const [rankingData, setRankingData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [pendingDelete, setPendingDelete] = useState(null);
  const [error, setError] = useState("");
  const rankingInProgress = HIDE_UNFINISHED_RANKINGS && UNFINISHED_RANKING_WAREHOUSES.has(warehouse);

  useEffect(() => {
    const controller = new AbortController();
    async function loadRanking() {
      setRankingData(null);
      setError("");
      if (rankingInProgress) {
        setLoading(false);
        return;
      }
      setLoading(true);
      try {
        const response = await fetch(`${API}/api/weekly/picking-rankings`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ warehouse }),
          signal: controller.signal
        });
        const json = await readApiJson(response);
        if (!response.ok) throw json;
        setRankingData(json);
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

  async function refreshRanking() {
    setError("");
    if (rankingInProgress) {
      setRankingData(null);
      setLoading(false);
      return;
    }
    setLoading(true);
    try {
      const response = await fetch(`${API}/api/weekly/picking-rankings`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ warehouse })
      });
      const json = await readApiJson(response);
      if (!response.ok) throw json;
      setRankingData(json);
    } catch (e) {
      const message = String(e?.message || "");
      setError(message === "Failed to fetch"
        ? "API is not reachable. Check that the backend is running on http://127.0.0.1:3001."
        : formatUploadError(e));
    } finally {
      setLoading(false);
    }
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
      setPendingDelete(null);
      await refreshRanking();
    } catch (e) {
      setError(formatUploadError(e));
    } finally {
      setDeleting(false);
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
            <p className="hint-line">Weighted Units = Non-Burst Units * Goods type factor* Floor factor</p>
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
            onDelete={setPendingDelete}
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

function RankingTable({ title, ranking, onDelete, disabled }) {
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
                <th>Weighted Units</th>
                <th>Hours</th>
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
                  <td>{row.employeeNo}</td>
                  <td>{row.name}</td>
                  <td>{formatCell(row.weightedUnits)}</td>
                  <td>{formatCell(row.hours ?? row.effectiveHours)}</td>
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

function RankingCallout({ icon, label, row, muted = false }) {
  return (
    <div className={muted ? "ranking-callout muted" : "ranking-callout"}>
      <span>{icon}{label}</span>
      <strong>{row.name || row.employeeNo}</strong>
      <small>{formatCell(row.efficiency)} weighted units/hour</small>
    </div>
  );
}
