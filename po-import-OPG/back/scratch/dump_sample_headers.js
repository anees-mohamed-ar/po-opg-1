const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const dir = 'c:/Users/IMT290/Desktop/Projects/po-OPG/OPG_input_excel_headers';
const files = {
    PO: 'OPg_PO_Headers.xlsx',
    GRN: 'OPG_GRN_Headers.xlsx',
    Purchase: 'OPG_Purchase_Headers.xlsx',
    Sales: 'OPG_Sales_Headers.xlsx',
    FI: 'OPG_FI_Headers.xlsx'
};

for (const [mod, file] of Object.entries(files)) {
    const wb = xlsx.readFile(path.join(dir, file));
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    const headers = rows[0].filter(h => h !== null && h !== undefined && String(h).trim() !== '');
    console.log(`\n=== ${mod} (${headers.length} headers in sample) ===`);
    console.log(headers);
}
