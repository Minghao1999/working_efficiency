import { Router } from "express";
import multer from "multer";
import { analyzeEfficiency } from "../services/efficiencyService.js";

const router = Router();
const upload = multer({ storage: multer.memoryStorage() });

router.post("/analyze", upload.array("files"), (req, res) => {
  try {
    const activeFiles = JSON.parse(req.body?.activeFiles || "[]");
    const fileKeys = JSON.parse(req.body?.fileKeys || "[]");
    res.json(analyzeEfficiency(req.files || [], {
      sessionId: req.body?.sessionId || "default",
      activeFiles,
      fileKeys,
      partial: req.body?.partial || ""
    }));
  } catch (error) {
    res.status(error.status || 400).json({ error: error.message, details: error.details || undefined });
  }
});

export default router;
