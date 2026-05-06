import { getDb } from "./mongo.js";

const DAY_MS = 24 * 60 * 60 * 1000;
const RETENTION_MONTHS = 1;
let cleanupTimer;

function retentionCutoff(now = new Date()) {
  const cutoff = new Date(now);
  cutoff.setMonth(cutoff.getMonth() - RETENTION_MONTHS);
  return {
    date: cutoff,
    iso: cutoff.toISOString(),
    day: cutoff.toISOString().slice(0, 10)
  };
}

export async function cleanupOldDatabaseData(now = new Date()) {
  const db = await getDb();
  const cutoff = retentionCutoff(now);

  const [feedback, unitCache, pickingCache] = await Promise.all([
    db.collection("feedback").deleteMany({ createdAt: { $lt: cutoff.iso } }),
    db.collection("weekly_unit_daily_cache").deleteMany({
      $or: [
        { businessDate: { $lt: cutoff.day } },
        { businessDate: { $exists: false }, createdAt: { $lt: cutoff.date } }
      ]
    }),
    db.collection("weekly_picking_daily_cache").deleteMany({
      $or: [
        { businessDate: { $lt: cutoff.day } },
        { businessDate: { $exists: false }, createdAt: { $lt: cutoff.date } }
      ]
    })
  ]);

  return {
    cutoff: cutoff.iso,
    deleted: {
      feedback: feedback.deletedCount,
      weeklyUnitDailyCache: unitCache.deletedCount,
      weeklyPickingDailyCache: pickingCache.deletedCount
    }
  };
}

export function startDatabaseCleanup() {
  if (cleanupTimer) return;

  const runCleanup = async () => {
    try {
      const result = await cleanupOldDatabaseData();
      console.log("Database cleanup complete", result);
    } catch (error) {
      console.warn("Database cleanup skipped:", error.message);
    }
  };

  setTimeout(runCleanup, 10_000).unref?.();
  cleanupTimer = setInterval(runCleanup, DAY_MS);
  cleanupTimer.unref?.();
}
