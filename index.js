import express from "express";
import cors from "cors";
import readExcelFile from "./readExcel.js";

const app = express();
const PORT = 3000;

app.use(cors());

app.get("/api/data", async (req, res) => {
  try {
    const results = await readExcelFile();
    res.json(results);
  } catch (error) {
    res.status(500).json({ error: "Error reading Excel file" });
    console.error(error);
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
