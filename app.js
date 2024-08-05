const express = require("express");
const { google } = require("googleapis");
const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");
const stream = require("stream");

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
  ["https://www.googleapis.com/auth/drive.readonly"]
);

const drive = google.drive({ version: "v3", auth });

// Load the Excel file from Google Drive
async function loadExcelFile(fileId) {
  try {
    const response = await drive.files.get(
      { fileId, alt: "media" },
      { responseType: "stream" }
    );

    return new Promise((resolve, reject) => {
      const chunks = [];
      response.data
        .on("data", (chunk) => chunks.push(chunk))
        .on("end", () => {
          const buffer = Buffer.concat(chunks);
          const workbook = xlsx.read(buffer);
          const sheetName = workbook.SheetNames[0];
          const sheet = workbook.Sheets[sheetName];
          const excelData = xlsx.utils.sheet_to_json(sheet);
          resolve(excelData);
        })
        .on("error", reject);
    });
  } catch (error) {
    console.error("Error fetching the file from Google Drive", error);
    throw error;
  }
}

// Store the Excel data in a variable
let excelData = [];
const fileId = "your_google_drive_file_id"; // Replace with your Google Drive file ID

// Load the data when the server starts
loadExcelFile(fileId)
  .then((data) => {
    excelData = data;
    console.log("Excel data loaded successfully.");
  })
  .catch((error) => {
    console.error("Failed to load Excel data:", error);
  });

// Render search page
app.get("/", (req, res) => {
  res.render("search");
});

// Handle search requests
app.post("/search", (req, res) => {
  const { query } = req.body;

  try {
    const results = excelData.filter((row) => {
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
    console.error("Error querying the Excel file", error);
    res.status(500).send("Error querying the Excel file");
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
