import { Router } from "express";
import multer from "multer";
import { analyzeWeekly } from "../services/weeklyService.js";

const router = Router();
const upload = multer({ storage: multer.memoryStorage() });

router.post("/analyze", upload.fields([{ name: "volume", maxCount: 1 }, { name: "isc", maxCount: 1 }, { name: "pick", maxCount: 1 }]), (req, res) => {
  try {
    const volumeFile = req.files?.volume?.[0];
    const iscFile = req.files?.isc?.[0];
    const pickFile = req.files?.pick?.[0];
    if (!volumeFile && !iscFile && !pickFile) throw new Error("?????????");
    res.json(analyzeWeekly({ volumeFile, iscFile, pickFile }));
  } catch (error) {
    res.status(error.status || 400).json({ error: error.message, details: error.details || undefined });
  }
});

export default router;
