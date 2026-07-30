const xlsx = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, '../../VENKAT_SAMPLE_Apr22-Mar23.xlsx');
try {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets['detail'];
  const rawRows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  
  console.log("Printing first 10 rows:");
  for (let i = 0; i < Math.min(10, rawRows.length); i++) {
    console.log(`Row ${i}:`, rawRows[i].slice(0, 15));
  }
} catch (err) {
  console.error("Error:", err.message);
}
