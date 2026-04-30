import app from "./app.js";

const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
  console.log(`WorkSight API running on http://127.0.0.1:${PORT}`);
});
