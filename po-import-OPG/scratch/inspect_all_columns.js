const xlsx = require('xlsx');
const realExcelPath = "c:\\Users\\IMT290\\Desktop\\Projects\\po-OPG\\Sample_PO_Apr23_Dec23.xlsx";
const workbook = xlsx.readFile(realExcelPath);
const sheet = workbook.Sheets['detail'];
const rawRows = xlsx.utils.sheet_to_json(sheet);
if (rawRows.length > 0) {
    const keys = Object.keys(rawRows[0]);
    console.log("All columns in sheet:");
    console.log(JSON.stringify(keys, null, 2));
}
