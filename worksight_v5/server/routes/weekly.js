import { Router } from "express";
import multer from "multer";
import { analyzeWeekly } from "../services/weeklyService.js";
import { queryPickingData, queryUnitData } from "../services/weeklyWmsService.js";

const router = Router();
const upload = multer({ storage: multer.memoryStorage() });

router.post("/analyze", upload.fields([{ name: "volume", maxCount: 1 }, { name: "isc", maxCount: 1 }, { name: "pick", maxCount: 1 }]), (req, res) => {
  try {
    const volumeFile = req.files?.volume?.[0];
    const iscFile = req.files?.isc?.[0];
    const pickFile = req.files?.pick?.[0];
    const existingDaily = req.body?.existingDaily ? JSON.parse(req.body.existingDaily) : [];
    if (!volumeFile && !iscFile && !pickFile) throw new Error("?????????");
    res.json(analyzeWeekly({ volumeFile, iscFile, pickFile, existingDaily }));
  } catch (error) {
    res.status(error.status || 400).json({ error: error.message, details: error.details || undefined });
  }
});

router.post("/query-unit", async (req, res) => {
  try {
    const { from, to, warehouse, warehouseNo } = req.body || {};
    res.json(await queryUnitData({ from, to, warehouse, warehouseNo }));
  } catch (error) {
    res.status(error.status || 400).json({ error: error.message, details: error.details || undefined });
  }
});

router.post("/query-picking", async (req, res) => {
  try {
    const { from, to, warehouse, warehouseNo, targetUpph } = req.body || {};
    res.json(await queryPickingData({ from, to, warehouse, warehouseNo, targetUpph }));
  } catch (error) {
    res.status(error.status || 400).json({ error: error.message, details: error.details || undefined });
  }
});

export default router;
