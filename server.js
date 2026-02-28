const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const XLSX = require("xlsx");
const fs = require("fs");

const app = express();
app.use(cors());
app.use(bodyParser.json());

const filePath = "contacts.xlsx";

app.post("/submit", (req, res) => {
  const { name, email, message } = req.body;

  const newEntry = [{
    Name: name,
    Email: email,
    Message: message,
    Date: new Date().toLocaleString()
  }];

  let workbook;
  let worksheet;

  if (fs.existsSync(filePath)) {
    workbook = XLSX.readFile(filePath);
    worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const existingData = XLSX.utils.sheet_to_json(worksheet);
    const updatedData = [...existingData, ...newEntry];
    const newSheet = XLSX.utils.json_to_sheet(updatedData);
    workbook.Sheets[workbook.SheetNames[0]] = newSheet;
  } else {
    workbook = XLSX.utils.book_new();
    worksheet = XLSX.utils.json_to_sheet(newEntry);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Contacts");
  }

  XLSX.writeFile(workbook, filePath);

  res.json({ message: "Form submitted successfully!" });
});

app.listen(5000, () => {
  console.log("Server running on port 5000");
});