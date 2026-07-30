const xlsx = require('xlsx');
const realExcelPath = "c:\\Users\\IMT290\\Desktop\\Projects\\po-OPG\\Sample_PO_Apr23_Dec23.xlsx";
const workbook = xlsx.readFile(realExcelPath);
const sheet = workbook.Sheets['detail'];
const rawRows = xlsx.utils.sheet_to_json(sheet);
const row = rawRows.find(r => String(r['Purchasing Document']).split('.')[0].trim() === '4600001039');
if (row) {
    console.log("Found PO 4600001039 row keys and values where val is non-zero/non-empty:");
    for (const key of Object.keys(row)) {
        const val = row[key];
        if (val !== undefined && val !== null && String(val).trim() !== '' && val !== 0) {
            console.log(`${key}: ${val}`);
        }
    }
} else {
    console.log("PO 4600001039 not found.");
}
