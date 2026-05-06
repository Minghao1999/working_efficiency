import app from "./app.js";
import { startDatabaseCleanup } from "./services/databaseCleanup.js";

const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
  console.log(`WorkSight API running on http://127.0.0.1:${PORT}`);
  startDatabaseCleanup();
});
