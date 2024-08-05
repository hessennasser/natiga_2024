const { google } = require("googleapis");
const path = require("path");
const fs = require("fs");
const express = require("express");
const app = express();

// Load client secrets from a local file
const credentials = JSON.parse(
  fs.readFileSync(path.join(__dirname, "service_account.json"))
);

// Configure JWT client for service account
const auth = new google.auth.JWT(
  credentials.client_email,
  null,
  credentials.private_key,
  ["https://www.googleapis.com/auth/spreadsheets.readonly"]
);

const sheets = google.sheets({ version: "v4", auth });

// Set up Express
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Search route
app.get("/search", async (req, res) => {
  const query = req.query.query || "";
  try {
    // Read data from the Google Sheets
    const spreadsheetId = "1-mICanz5G0t1CziTSEQ1Q4L-9F1ff7vn"; // Your spreadsheet ID
    const range = "Sheet1!A:D"; // Adjust the range based on your sheet structure

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });

    const rows = response.data.values;
    if (!rows.length) {
      res.send("No data found.");
      return;
    }

    // Filter rows based on query
    const results = rows.filter((row) =>
      row.some((cell) => cell.toLowerCase().includes(query.toLowerCase()))
    );

    res.json(results);
  } catch (error) {
    console.error("Error searching for data:", error);
    res.status(500).send("Error searching for data");
  }
});

// Start Express server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server started on http://localhost:${PORT}`);
});
