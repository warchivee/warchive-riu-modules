import express from "express";
import cors from "cors";
import parseExcelFile from "./parse.js";

const app = express();
const PORT = 3000;

app.use(cors());

app.get("/", async (req, res) => {
  try {
    const results = await parseExcelFile();
    res.json(results);
  } catch (error) {
    res.status(500).json({ error: "Error parsing Excel file" });
    console.error(error);
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
