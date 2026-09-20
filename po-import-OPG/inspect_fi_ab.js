const xlsx = require('xlsx');
const path = require('path');

const excelPath = "c:\\Users\\IMT290\\Desktop\\Projects\\po-OPG\\imatrix_FI data _2023_3100.XLSX";
const workbook = xlsx.readFile(excelPath);

console.log("Sheets:", workbook.SheetNames);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = xlsx.utils.sheet_to_json(sheet);

console.log("Total rows:", rows.length);
if (rows.length > 0) {
    console.log("Sample headers:", Object.keys(rows[0]));
}

const abRows = rows.filter(r => {
    let docType = '';
    for (const k of Object.keys(r)) {
        if (k.toLowerCase().includes('doc') && k.toLowerCase().includes('type')) {
            docType = r[k];
            break;
        }
    }
    return String(docType).trim().toUpperCase() === 'AB';
});

console.log("\nTotal 'AB' doc type rows found:", abRows.length);
if (abRows.length > 0) {
    console.log("Sample AB rows (first 10):");
    abRows.slice(0, 10).forEach((r, idx) => {
        console.log(`--- Row ${idx+1} ---`);
        console.log(JSON.stringify(r, null, 2));
    });
} else {
    // Print all unique doc types found in file
    const docTypes = new Set();
    rows.forEach(r => {
        for (const k of Object.keys(r)) {
            if (k.toLowerCase().includes('type')) {
                docTypes.add(r[k]);
            }
        }
    });
    console.log("Doc types found in sheet:", Array.from(docTypes));
}
