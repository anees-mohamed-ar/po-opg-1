const xlsx = require('xlsx');
const path = require('path');

const excelPath = path.resolve('..', 'Sample_PO_Apr23_Dec23.xlsx');

try {
    const workbook = xlsx.readFile(excelPath);
    const sheet = workbook.Sheets['detail'];
    const rawRows = xlsx.utils.sheet_to_json(sheet);
    
    // Find a row where WOTB, WOTB_1 or FRB1, FRB1  - Vendor is non-zero/non-empty
    for (const row of rawRows) {
        const delInd = row['Deletion Indicator'];
        if (delInd && String(delInd).trim().toUpperCase() === 'L') continue;

        if (row['WOTB'] && row['WOTB_1']) {
            console.log('Row with WOTB & WOTB_1 found:');
            console.log('WOTB:', row['WOTB']);
            console.log('WOTB_1:', row['WOTB_1']);
            console.log('Vendor keys matching WOTB:', Object.keys(row).filter(k => k.startsWith('WOTB')));
            break;
        }
    }

    for (const row of rawRows) {
        const delInd = row['Deletion Indicator'];
        if (delInd && String(delInd).trim().toUpperCase() === 'L') continue;

        if (row['FRB1'] && row['FRB1  - Vendor']) {
            console.log('\nRow with FRB1 & FRB1  - Vendor found:');
            console.log('FRB1:', row['FRB1']);
            console.log('FRB1  - Vendor:', row['FRB1  - Vendor']);
            break;
        }
    }
} catch (e) {
    console.error(e);
}
