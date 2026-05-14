import "dotenv/config";
import express from "express";
import cors from "cors";
import efficiencyRouter from "./routes/efficiency.js";
import weeklyRouter from "./routes/weekly.js";
import exportRouter from "./routes/export.js";
import pickingExceptionRouter from "./routes/pickingException.js";
import feedbackRouter from "./routes/feedback.js";

const app = express();

const allowedOrigin = process.env.FRONTEND_URL;
app.use(cors({
  origin: allowedOrigin || true
}));
app.use(express.json({ limit: "20mb" }));

app.get("/api/health", (_req, res) => res.json({ ok: true }));
app.use("/api/efficiency", efficiencyRouter);
app.use("/api/weekly", weeklyRouter);
app.use("/api/export", exportRouter);
app.use("/api/picking-exception", pickingExceptionRouter);
app.use("/api/feedback", feedbackRouter);

app.use((error, _req, res, _next) => {
  res.status(error.status || 400).json({
    error: error.code === "LIMIT_UNEXPECTED_FILE"
      ? "Too many files were uploaded for this field."
      : error.message || "Request failed"
  });
});

export default app;
