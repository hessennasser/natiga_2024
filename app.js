const express = require("express");
const path = require("path");
const fs = require("fs");
const xlsx = require("xlsx");

const app = express();
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Path to the local Excel file
const excelFilePath = path.join(__dirname, "data.xlsx");

// Load the data from the local Excel file
function loadExcelData() {
  try {
    const workbook = xlsx.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0]; // Get the first sheet
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 }); // Convert to array of arrays

    return data;
  } catch (error) {
    console.error("Error reading the Excel file", error);
    throw error;
  }
}

// Store the sheet data in a variable
let sheetData = [];
try {
  const data = loadExcelData();
  const headers = data[0];
  sheetData = data.slice(1).map((row) => {
    let obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
} catch (error) {
  console.error("Error loading sheet data", error);
}

// Render search page
app.get("/", (req, res) => {
  res.render("search");
});
app.get("/search", (req, res) => {
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

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
