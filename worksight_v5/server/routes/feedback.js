import express from "express";
import { getDb } from "../services/mongo.js";

const router = express.Router();
const COLLECTION = "feedback";

function serializeFeedback(item) {
  return {
    id: item._id.toString(),
    type: item.type || "feedback",
    title: item.title || "",
    message: item.message,
    author: item.author || "",
    status: item.status || "new",
    createdAt: item.createdAt
  };
}

router.get("/", async (_req, res) => {
  try {
    const db = await getDb();
    const items = await db
      .collection(COLLECTION)
      .find({})
      .sort({ createdAt: -1 })
      .limit(100)
      .toArray();

    res.json({ success: true, items: items.map(serializeFeedback) });
  } catch (error) {
    res.status(503).json({
      success: false,
      message: error.message
    });
  }
});

router.post("/", async (req, res) => {
  const message = String(req.body?.message || "").trim();
  const author = String(req.body?.author || "").trim();

  if (!message) {
    return res.status(400).json({ success: false, message: "Feedback content is required." });
  }

  try {
    const db = await getDb();
    const document = {
      message,
      author,
      status: "new",
      createdAt: new Date().toISOString()
    };
    const result = await db.collection(COLLECTION).insertOne(document);

    res.status(201).json({
      success: true,
      item: serializeFeedback({ ...document, _id: result.insertedId })
    });
  } catch (error) {
    res.status(503).json({
      success: false,
      message: error.message
    });
  }
});

export default router;
