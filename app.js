const express = require("express");
const path = require("path");
const fs = require("fs");
const ExcelJS = require("exceljs");

const app = express();
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Path to the local Excel file
const excelFilePath = path.join(__dirname, "data.xlsx");

// Load the data from the local Excel file (streaming approach)
async function loadSheetData() {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);
    const sheet = workbook.worksheets[0];

    const data = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) {
        data.push(row.values.slice(1)); // First row as header
      } else {
        data.push(row.values.slice(1));
      }
    });

    return data;
  } catch (error) {
    console.error("Error reading the Excel file", error);
    throw error;
  }
}

// Store the sheet data in a variable
let sheetData = [];
(async () => {
  try {
    const data = await loadSheetData();
    const headers = data[0];
    sheetData = data.slice(1).map((row) => {
      let obj = {};
      headers.forEach((header, i) => {
        obj[header] = row[i] || "";
      });
      return obj;
    });
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
      const keys = Object.keys(row);
      if (keys.length < 2) return false; // Ensure there are at least two columns

      const firstField = row[keys[0]];
      const secondField = row[keys[1]];

      return (
        firstField.toString().toLowerCase().includes(query.toLowerCase()) ||
        secondField.toString().toLowerCase().includes(query.toLowerCase())
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
