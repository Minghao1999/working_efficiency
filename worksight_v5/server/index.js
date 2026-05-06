import app from "./app.js";
import { startDatabaseCleanup } from "./services/databaseCleanup.js";

const PORT = process.env.PORT || 3001;

app.listen(PORT, "0.0.0.0", () => {
  console.log(`WorkSight API running on port ${PORT}`);
});