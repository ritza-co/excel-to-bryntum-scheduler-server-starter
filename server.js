import express from "express";
import cors from "cors";
import fs from "fs/promises";

const app = express();
const port = 3000;

app.use(
  cors({
    origin: ["http://127.0.0.1:5173", "http://localhost:5173"],
    credentials: true,
  })
);

app.get("/download", async (req, res) => {
  try {
    const data = await fs.readFile("./data/scheduler-data.json", "utf-8");
    const jsonData = JSON.parse(data);
    res.json(jsonData);
  } catch (error) {
    console.error("Error: ", error);
    res.status(500).send("An error occurred.");
  }
});

app.listen(port, () => {
  console.log(`Server listening at http://localhost:${port}`);
});