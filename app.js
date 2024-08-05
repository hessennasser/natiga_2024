const express = require("express");
const path = require("path");
const ExcelJS = require("exceljs");

const app = express();
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Path to the local Excel file
const excelFilePath = path.join(__dirname, "data.xlsx");

// Store the sheet data in memory (use efficient structure)
let sheetData = [];

async function loadSheetData() {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);
    const sheet = workbook.worksheets[0];

    sheetData = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row

      let obj = {};
      row.values.forEach((value, index) => {
        obj[`Column${index}`] = value || "";
      });
      sheetData.push(obj);
    });
  } catch (error) {
    console.error("Error reading the Excel file", error);
    throw error;
  }
}

// Load sheet data once when the server starts
(async () => {
  try {
    await loadSheetData();
  } catch (error) {
    console.error("Error loading sheet data", error);
  }
})();

// Render search page
app.get("/", (req, res) => {
  res.render("search");
});

// Handle search requests
app.post("/search", (req, res) => {
  const { query } = req.body;

  try {
    const results = sheetData.filter((row) => {
      const values = Object.values(row);
      return values.some((value) =>
        value.toString().toLowerCase().includes(query.toLowerCase())
      );
    });

    res.render("results", { results });
  } catch (error) {
    console.error("Error querying the sheet data", error);
    res.status(500).send("Error querying the sheet data");
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
