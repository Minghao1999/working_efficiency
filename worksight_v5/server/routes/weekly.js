import { Router } from "express";
import multer from "multer";
import { analyzeWeekly } from "../services/weeklyService.js";
import { excludePickingRankingPerson, queryPickingData, queryPickingRankings, queryUnitData } from "../services/weeklyWmsService.js";

const router = Router();
const upload = multer({ storage: multer.memoryStorage() });
const weeklyUpload = upload.fields([{ name: "volume", maxCount: 1 }, { name: "isc", maxCount: 1 }, { name: "pick", maxCount: 2 }]);

router.post("/analyze", (req, res) => {
  weeklyUpload(req, res, (uploadError) => {
    if (uploadError) {
      return res.status(400).json({
        error: uploadError.code === "LIMIT_UNEXPECTED_FILE"
          ? "Too many picking files. Big-wave picking supports up to 2 files."
          : uploadError.message
      });
    }

  try {
    const volumeFile = req.files?.volume?.[0];
    const iscFile = req.files?.isc?.[0];
    const pickFiles = req.files?.pick || [];
    const existingDaily = req.body?.existingDaily ? JSON.parse(req.body.existingDaily) : [];
    if (!volumeFile && !iscFile && !pickFiles.length) throw new Error("?????????");
    res.json(analyzeWeekly({ volumeFile, iscFile, pickFiles, existingDaily }));
  } catch (error) {
    res.status(error.status || 400).json({ error: error.message, details: error.details || undefined });
  }
  });
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
    const { from, to, warehouse, warehouseNo, targetUpph, includeBigWavePick } = req.body || {};
    res.json(await queryPickingData({ from, to, warehouse, warehouseNo, targetUpph, includeBigWavePick }));
  } catch (error) {
    res.status(error.status || 400).json({ error: error.message, details: error.details || undefined });
  }
});

router.post("/picking-rankings", async (req, res) => {
  try {
    const { warehouse, warehouseNo } = req.body || {};
    res.json(await queryPickingRankings({ warehouse, warehouseNo }));
  } catch (error) {
    res.status(error.status || 400).json({ error: error.message, details: error.details || undefined });
  }
});

router.post("/picking-rankings/exclude", async (req, res) => {
  try {
    const { employeeNo, name } = req.body || {};
    res.json(await excludePickingRankingPerson({ employeeNo, name }));
  } catch (error) {
    res.status(error.status || 400).json({ error: error.message, details: error.details || undefined });
  }
});

export default router;
