const express = require("express");
const { google } = require("googleapis");
const path = require("path");
const fs = require("fs");

const app = express();
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Google Drive API configuration
const credentials = JSON.parse(
  fs.readFileSync(path.join(__dirname, "service_account.json"))
);
const auth = new google.auth.JWT(
  credentials.client_email,
  null,
  credentials.private_key,
  ["https://www.googleapis.com/auth/spreadsheets.readonly"]
);

const sheets = google.sheets({ version: "v4", auth });

// Your Google Sheets file ID and range
const spreadsheetId = "1-mICanz5G0t1CziTSEQ1Q4L-9F1ff7vn"; // Replace with your Google Sheets ID
const range = "Sheet1"; // Adjust if your sheet name is different

// Load the Excel data from Google Sheets
async function loadSheetData() {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });
    return response.data.values;
  } catch (error) {
    console.error("Error fetching the file from Google Sheets", error);
    throw error;
  }
}

// Store the sheet data in a variable
let sheetData = [];
loadSheetData()
  .then((data) => {
    // Convert sheet data to JSON format if necessary
    // For example, assuming the first row is headers
    const headers = data[0];
    sheetData = data.slice(1).map((row) => {
      let obj = {};
      headers.forEach((header, i) => {
        obj[header] = row[i];
      });
      return obj;
    });
  })
  .catch((error) => {
    console.error("Error loading sheet data", error);
  });

// Render search page
app.get("/", (req, res) => {
  res.render("search");
});

// Handle search requests
app.post("/search", (req, res) => {
  const { query } = req.body;

  try {
    const results = sheetData.filter((row) => {
      // Extract the first and second keys (columns)
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
