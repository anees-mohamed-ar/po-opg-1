const xlsx = require('xlsx');
const { getRowValue } = require('../tallyXMLBuilder');
const realExcelPath = "c:\\Users\\IMT290\\Desktop\\Projects\\po-OPG\\Sample_PO_Apr23_Dec23.xlsx";
const workbook = xlsx.readFile(realExcelPath);
const sheet = workbook.Sheets['detail'];
const rawRows = xlsx.utils.sheet_to_json(sheet);
const rows = rawRows.filter(r => String(r['Purchasing Document']).split('.')[0].trim() === '4600001071');
console.log(`Found ${rows.length} rows for 4600001071.`);
rows.forEach((row, i) => {
    console.log(`Row ${i + 1}:`);
    console.log(`  Deletion Indicator: ${row['Deletion Indicator']}`);
    console.log(`  Material: ${row['Material']}`);
    console.log(`  Order Qty: ${row['Order Quantity']}`);
});
