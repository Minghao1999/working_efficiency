import React, { useEffect, useState } from "react";
import { Trash2, X } from "lucide-react";
import { fileIdentity } from "../utils/files";

export function UploadBox({ title, caption, disabled = false, onChange, actionSlot = null }) {
  const [fileName, setFileName] = useState("");
  const clearFile = () => {
    setFileName("");
    onChange(null);
  };
  return (
    <div className={`${disabled ? "upload-box disabled" : "upload-box"} ${actionSlot ? "upload-box-with-tools" : ""}`}>
      <div className="upload-box-main">
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
        {fileName && (
          <div className="selected-file single-upload-file">
            <span>{fileName}</span>
            <button type="button" onClick={clearFile} aria-label={`Remove ${fileName}`}>
              <Trash2 size={15} />
            </button>
          </div>
        )}
      </div>
      {actionSlot && (
        <div className="upload-box-tools">
          {actionSlot}
        </div>
      )}
    </div>
  );
}

export function LaborHoursInput({ mode, setMode, onFileChange }) {
  const [fileName, setFileName] = useState("");

  useEffect(() => {
    if (mode === "manual") setFileName("");
  }, [mode]);

  const clearFile = () => {
    setFileName("");
    onFileChange(null);
  };

  return (
    <div className="upload-box labor-hours-box">
      <div className="upload-box-head">
        <strong>Labor Hours</strong>
        <div className="segmented compact sliding" style={{ "--active-index": mode === "excel" ? 0 : 1, "--segment-count": 2 }}>
          <button type="button" className={mode === "excel" ? "active" : ""} onClick={() => setMode("excel")}>Excel Upload</button>
          <button type="button" className={mode === "manual" ? "active" : ""} onClick={() => setMode("manual")}>Manual Entry</button>
        </div>
      </div>
      <span>ISC 出勤工时(选择Global Export).</span>
      {mode === "excel" ? (
        <>
          <FilePicker
            accept=".xlsx,.xls"
            files={fileName ? [{ name: fileName }] : []}
            onChange={(files) => {
              const file = files[0] || null;
              setFileName(file?.name || "");
              onFileChange(file);
            }}
          />
          {fileName && (
            <div className="selected-file single-upload-file">
              <span>{fileName}</span>
              <button type="button" onClick={clearFile} aria-label={`Remove ${fileName}`}>
                <Trash2 size={15} />
              </button>
            </div>
          )}
        </>
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

export function GlassSelect({ value, options, onChange, onRemoveOption, className = "" }) {
  const [open, setOpen] = useState(false);
  const normalizedOptions = options.map((option) => (
    typeof option === "object" ? option : { value: option, label: option }
  ));
  const selected = normalizedOptions.find((option) => option.value === value) || normalizedOptions[0] || { value: "", label: "" };

  return (
    <div className={`glass-select ${open ? "open" : ""} ${className}`} onBlur={() => setOpen(false)}>
      <button type="button" className="glass-select-trigger" onClick={() => setOpen((current) => !current)}>
        <span>{selected.label}</span>
        <span className="glass-select-caret">⌄</span>
      </button>
      {open && (
        <div className="glass-select-menu">
          {normalizedOptions.map((option) => (
            <div key={option.value} className={option.value === selected.value ? "glass-select-option active" : "glass-select-option"}>
              <button
                type="button"
                className="glass-select-option-main"
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => {
                  onChange(option.value);
                  setOpen(false);
                }}
              >
                {option.label}
              </button>
              {onRemoveOption && option.value !== "All People" && (
                <button
                  type="button"
                  className="glass-select-option-remove"
                  aria-label={`Remove ${option.label}`}
                  title={`Remove ${option.label}`}
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={(event) => {
                    event.stopPropagation();
                    onRemoveOption(option.value);
                  }}
                >
                  <X size={15} />
                </button>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export function ChartPanel({ title, caption, children }) {
  return <div className="panel chart-panel"><h2>{title}</h2>{caption && <p className="hint-line">{caption}</p>}{children}</div>;
}

export function SelectLine({ label, value, options, onChange, onRemoveOption }) {
  return <label className="select-line">{label}<GlassSelect value={value} options={options} onChange={onChange} onRemoveOption={onRemoveOption} /></label>;
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
