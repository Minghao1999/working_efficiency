import { Router } from "express";
import multer from "multer";
import fs from "node:fs";
import path from "node:path";
import { analyzeEfficiency } from "../services/efficiencyService.js";

const router = Router();
const uploadDir = path.join(process.cwd(), "server", ".uploads");
fs.mkdirSync(uploadDir, { recursive: true });

const upload = multer({
  storage: multer.diskStorage({
    destination: (_req, _file, cb) => cb(null, uploadDir),
    filename: (_req, file, cb) => {
      const safeName = path.basename(file.originalname || "upload.xlsx").replace(/[^\w.-]+/g, "_");
      cb(null, `${Date.now()}-${Math.random().toString(16).slice(2)}-${safeName}`);
    }
  }),
  limits: {
    fileSize: 300 * 1024 * 1024,
    files: 12
  }
});

function cleanupFiles(files = []) {
  for (const file of files) {
    if (!file.path) continue;
    fs.promises.unlink(file.path).catch(() => {});
  }
}

router.post("/analyze", (req, res) => {
  upload.array("files")(req, res, (uploadError) => {
    try {
      if (uploadError) {
        const message = uploadError.code === "LIMIT_FILE_SIZE"
          ? "Uploaded workbook is too large. Please keep each file under 300MB."
          : uploadError.message;
        const status = uploadError.code?.startsWith("LIMIT_") ? 413 : 400;
        res.status(status).json({ error: message });
        return;
      }
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
    } finally {
      cleanupFiles(req.files);
    }
  });
});

export default router;
