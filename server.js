import express from "express";
import cors from "cors";

const app = express();
const port = 3000;

app.use(
  cors({
    origin: ["http://127.0.0.1:5173", "http://localhost:5173"],
    credentials: true,
  })
);

app.listen(port, () => {
  console.log(`Server listening at http://localhost:${port}`);
});