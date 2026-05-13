import "dotenv/config";

process.env.TZ = process.env.APP_TIME_ZONE || process.env.TZ || "America/New_York";

const { default: app } = await import("./app.js");
const { startDatabaseCleanup } = await import("./services/databaseCleanup.js");
const { startPickingRankingScheduler } = await import("./services/weeklyWmsService.js");

const PORT = process.env.PORT || 3001;

app.listen(PORT, "0.0.0.0", () => {
  console.log(`WorkSight API running on port ${PORT} (${process.env.TZ})`);
  startDatabaseCleanup();
  startPickingRankingScheduler();
});
