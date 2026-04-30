import express from "express";
import cors from "cors";
import efficiencyRouter from "./routes/efficiency.js";
import weeklyRouter from "./routes/weekly.js";
import exportRouter from "./routes/export.js";

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

export default app;
