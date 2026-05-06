import React, { useState } from "react";
import { columnLabel, formatCell, sortRows } from "../utils/formatters";

export function DataTable({ rows, lowUpph, highlightWait }) {
  if (!rows?.length) return <div className="empty">No data</div>;
  const keys = Object.keys(rows[0]).filter((k) => !k.startsWith("_"));
  return (
    <div className="table-wrap">
      <table>
        <thead><tr>{keys.map((k) => <th key={k}>{columnLabel(k)}</th>)}</tr></thead>
        <tbody>
          {rows.map((row, idx) => {
            const upph = Number(row.UPPH);
            const low = lowUpph && Number.isFinite(upph) && upph < lowUpph;
            return <tr key={idx} className={low ? "low-row" : ""}>{keys.map((k) => <td key={k} className={highlightWait && k === "Wait" && parseInt(row[k]) > 5 ? "danger-text" : ""}>{formatCell(row[k])}</td>)}</tr>;
          })}
        </tbody>
      </table>
    </div>
  );
}

export function DailyDetailTable({ rows, lowUpph, manual, manualHours, setManualHours }) {
  if (!rows?.length) return <div className="empty">No data</div>;
  const keys = Object.keys(rows[0]).filter((k) => !k.startsWith("_"));
  return (
    <div className="table-wrap">
      <table>
        <thead><tr>{keys.map((k) => <th key={k}>{columnLabel(k)}</th>)}</tr></thead>
        <tbody>
          {rows.map((row, idx) => {
            const upph = Number(row.UPPH);
            const low = lowUpph && Number.isFinite(upph) && upph < lowUpph;
            return (
              <tr key={row.业务日期 || idx} className={low ? "low-row" : ""}>
                {keys.map((k) => (
                  <td key={k}>
                    {manual && k === "总工时" ? (
                      <input
                        className="table-number-input"
                        type="number"
                        min="0"
                        step="0.01"
                        inputMode="decimal"
                        placeholder="Enter hours"
                        value={manualHours[row.业务日期] ?? ""}
                        onChange={(e) => {
                          const value = e.target.value;
                          setManualHours((current) => ({ ...current, [row.业务日期]: value }));
                        }}
                      />
                    ) : (
                      formatCell(row[k])
                    )}
                  </td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

export function EditablePersonTable({ rows, markedForDelete = new Set(), setMarkedForDelete = () => {}, hiddenColumns = [] }) {
  const [sort, setSort] = useState(null);

  if (!rows.length) return <div className="empty">No data</div>;
  const sorted = sortRows(rows, sort);
  const hidden = new Set(hiddenColumns);
  const hiddenLabels = new Set(["Attendance Hours", "Effective Hours Ratio", ...hiddenColumns.map((key) => columnLabel(key))]);
  const keys = Object.keys(sorted[0]).filter((key) => !key.startsWith("_") && !hidden.has(key) && !hiddenLabels.has(columnLabel(key)));
  const onSort = (key) => {
    setSort((current) => {
      if (!current || current.key !== key) return { key, direction: "asc" };
      if (current.direction === "asc") return { key, direction: "desc" };
      return null;
    });
  };

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {keys.map((k) => (
              <th key={k}>
                <button className="sort-header" onClick={() => onSort(k)}>
                  <span>{columnLabel(k)}</span>
                  <span className="sort-mark">{sort?.key === k ? (sort.direction === "asc" ? "↑" : "↓") : "↕"}</span>
                </button>
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {sorted.map((row) => {
            const id = `${row.日期}|${row.工号}`;
            return <tr key={id} className={row._deleted ? "deleted-row" : ""}>{keys.map((k) => <td key={k}>{k === "删除" ? <input type="checkbox" checked={markedForDelete.has(id)} onChange={(e) => {
              const next = new Set(markedForDelete);
              e.target.checked ? next.add(id) : next.delete(id);
              setMarkedForDelete(next);
            }} /> : formatCell(row[k])}</td>)}</tr>;
          })}
        </tbody>
      </table>
    </div>
  );
}
