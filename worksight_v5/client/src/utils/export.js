import { API } from "../constants";

export async function exportRows(rows, sheetName, fileName) {
  const res = await fetch(`${API}/api/export`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ rows, sheetName, fileName }) });
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}
