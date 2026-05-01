import express from "express";
import { searchPickingExceptionLocation } from "../services/pickingExceptionService.js";

const router = express.Router();

router.post("/search-location", async (req, res) => {
  const result = await searchPickingExceptionLocation({
    barcode: req.body?.barcode,
    containerNo: req.body?.container_no ?? req.body?.containerNo,
    warehouse: req.body?.warehouse
  });

  res.status(result.success ? 200 : 400).json(result);
});

export default router;
