import React, { useEffect, useState } from "react";
import { Trash2 } from "lucide-react";
import { fileIdentity } from "../utils/files";

export function UploadBox({ title, caption, disabled = false, onChange }) {
  const [fileName, setFileName] = useState("");
  return (
    <label className={disabled ? "upload-box disabled" : "upload-box"}>
      <strong>{title}</strong>
      <span>{caption}</span>
      <FilePicker
        accept=".xlsx,.xls"
        disabled={disabled}
        files={fileName ? [{ name: fileName }] : []}
        onChange={(files) => {
          const file = files[0] || null;
          setFileName(file?.name || "");
          onChange(file);
        }}
      />
    </label>
  );
}

export function LaborHoursInput({ mode, setMode, onFileChange }) {
  const [fileName, setFileName] = useState("");

  useEffect(() => {
    if (mode === "manual") setFileName("");
  }, [mode]);

  return (
    <div className="upload-box">
      <strong>Labor Hours</strong>
      <span>ISC 出勤工时(选择Global Export).</span>
      <div className="segmented">
        <button type="button" className={mode === "excel" ? "active" : ""} onClick={() => setMode("excel")}>Excel Upload</button>
        <button type="button" className={mode === "manual" ? "active" : ""} onClick={() => setMode("manual")}>Manual Entry</button>
      </div>
      {mode === "excel" ? (
        <FilePicker
          accept=".xlsx,.xls"
          files={fileName ? [{ name: fileName }] : []}
          onChange={(files) => {
            const file = files[0] || null;
            setFileName(file?.name || "");
            onFileChange(file);
          }}
        />
      ) : (
        <div className="hint-line">Enter Total Hours directly in the Daily Detail table.</div>
      )}
    </div>
  );
}

export function FilePicker({ multiple = false, accept, disabled = false, files = [], onChange }) {
  const label = files.length ? files.map((file) => file.name).join(", ") : "No file selected";
  return (
    <label className={disabled ? "file-picker disabled" : "file-picker"}>
      <span>Choose File{multiple ? "s" : ""}</span>
      <input
        type="file"
        multiple={multiple}
        accept={accept}
        disabled={disabled}
        onChange={(e) => onChange([...e.target.files])}
      />
      <em>{label}</em>
    </label>
  );
}

export function SelectedFileList({ files, onRemove }) {
  if (!files.length) return null;
  return (
    <div className="selected-files">
      {files.map((file) => (
        <div className="selected-file" key={fileIdentity(file)}>
          <span>{file.name}</span>
          <button type="button" onClick={() => onRemove(fileIdentity(file))} aria-label={`Remove ${file.name}`}>
            <Trash2 size={15} />
          </button>
        </div>
      ))}
    </div>
  );
}

export function Metric({ icon, label, value, note = "" }) {
  return (
    <div className="metric">
      {icon && <span className="metric-icon">{icon}</span>}
      <span>{label}</span>
      <strong>{value}</strong>
      {note && <small>{note}</small>}
    </div>
  );
}

export function ProgressBar({ value, label }) {
  return (
    <div className="progress-block" aria-live="polite">
      <div className="progress-meta">
        <span>{label}</span>
        <strong>{Math.round(value)}%</strong>
      </div>
      <div className="progress-track">
        <div className="progress-fill" style={{ width: `${Math.min(100, Math.max(0, value))}%` }} />
      </div>
    </div>
  );
}

export function Tabs({ tabs, value, onChange }) {
  const activeIndex = Math.max(0, tabs.findIndex(([id]) => id === value));
  return (
    <div className="tabs" style={{ "--active-index": activeIndex, "--tab-count": tabs.length }}>
      {tabs.map(([id, label]) => <button key={id} className={value === id ? "active" : ""} onClick={() => onChange(id)}>{label}</button>)}
    </div>
  );
}

export function GlassSelect({ value, options, onChange, className = "" }) {
  const [open, setOpen] = useState(false);
  const selected = options.includes(value) ? value : options[0] || "";

  return (
    <div className={`glass-select ${open ? "open" : ""} ${className}`} onBlur={() => setOpen(false)}>
      <button type="button" className="glass-select-trigger" onClick={() => setOpen((current) => !current)}>
        <span>{selected}</span>
        <span className="glass-select-caret">⌄</span>
      </button>
      {open && (
        <div className="glass-select-menu">
          {options.map((option) => (
            <button
              type="button"
              key={option}
              className={option === selected ? "active" : ""}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                onChange(option);
                setOpen(false);
              }}
            >
              {option}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

export function ChartPanel({ title, caption, children }) {
  return <div className="panel chart-panel"><h2>{title}</h2>{caption && <p className="hint-line">{caption}</p>}{children}</div>;
}

export function SelectLine({ label, value, options, onChange }) {
  return <label className="select-line">{label}<GlassSelect value={value} options={options} onChange={onChange} /></label>;
}

export function ConfirmDialog({ title, message, onCancel, onConfirm }) {
  return (
    <div className="modal-backdrop" role="presentation">
      <div className="modal" role="dialog" aria-modal="true" aria-labelledby="confirm-title">
        <h2 id="confirm-title">{title}</h2>
        <p>{message}</p>
        <div className="modal-actions">
          <button className="ghost-btn" onClick={onCancel}>Cancel</button>
          <button className="danger-btn" onClick={onConfirm}>Delete</button>
        </div>
      </div>
    </div>
  );
}
