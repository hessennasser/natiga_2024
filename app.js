const express = require("express");
const xlsx = require("xlsx");
const path = require("path");

const app = express();
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Load the Excel file
const filePath = path.join(__dirname, "data.xlsx");
const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const excelData = xlsx.utils.sheet_to_json(sheet);

// Render search page
app.get("/", (req, res) => {
  res.render("search");
});

// Handle search requests
app.post("/search", (req, res) => {
  const { query } = req.body;

  try {
    const results = excelData.filter((row) =>
      Object.values(row).some((value) =>
        value.toString().toLowerCase().includes(query.toLowerCase())
      )
    );
    res.render("results", { results });
  } catch (error) {
    console.error("Error querying the Excel file", error);
    res.status(500).send("Error querying the Excel file");
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
