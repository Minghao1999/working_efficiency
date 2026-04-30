import { Router } from "express";
import XLSX from "xlsx";

const router = Router();

router.post("/", (req, res) => {
  const rows = req.body?.rows || [];
  const sheetName = req.body?.sheetName || "????";
  const fileName = req.body?.fileName || "export.xlsx";
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), sheetName);
  const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.send(buffer);
});

export default router;
