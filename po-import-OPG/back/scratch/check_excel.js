const xlsx = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, '../../VENKAT_SAMPLE_Apr22-Mar23.xlsx');
try {
  const workbook = xlsx.readFile(filePath);
  console.log("Sheet names:", workbook.SheetNames);

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rawRows = xlsx.utils.sheet_to_json(sheet);
  console.log("Number of rows:", rawRows.length);
  if (rawRows.length > 0) {
    console.log("First row keys:", Object.keys(rawRows[0]));
    console.log("First row data sample:", rawRows[0]);
  }
} catch (err) {
  console.error("Error reading file:", err.message);
}
