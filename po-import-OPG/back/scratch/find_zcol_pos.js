const xlsx = require('xlsx');
const { getRowValue } = require('../tallyXMLBuilder');
const realExcelPath = "c:\\Users\\IMT290\\Desktop\\Projects\\po-OPG\\Sample_PO_Apr23_Dec23.xlsx";
const workbook = xlsx.readFile(realExcelPath);
const sheet = workbook.Sheets['detail'];
const rawRows = xlsx.utils.sheet_to_json(sheet);

const ledgerColumns = [
    'FRB1', 'FRB2', 'FRC2', 'JEXS', 'NAVS', 'P001', 'P101', 'PB00', 'PBXX  -  Gross Price', 'R003', 'RA00', 'RA01', 'SKTO', 'WOTB',
    'ZBCA', 'ZBCD', 'ZBED', 'ZCEC', 'ZCEQ', 'ZCST', 'ZDSD', 'ZEQP', 'ZFRQ', 'ZFRV', 'ZHAN', 'ZINS', 'ZLAN', 'ZMFR', 'ZMIS', 'ZNE1', 'ZNE2', 'ZPAC', 'ZPF%', 'ZPNF', 'ZROY', 'ZRTQ', 'ZRUQ', 'ZSIZ', 'ZSTC', 'ZSTP', 'ZSTV', 'ZVAT', 'ZVIN', 'ZWRF'
];

const zcolRows = rawRows.filter(row => {
    const docType = String(getRowValue(row, 'PO -  Doc Type') || '').trim();
    return docType === 'ZCOL';
});

console.log(`Found ${zcolRows.length} rows for ZCOL.`);

// Group by PO
const groups = {};
zcolRows.forEach(row => {
    const poNum = String(row['Purchasing Document']).split('.')[0].trim();
    if (!groups[poNum]) groups[poNum] = [];
    groups[poNum].push(row);
});

for (const poNum of Object.keys(groups)) {
    console.log(`\nPO: ${poNum}`);
    const items = groups[poNum];
    ledgerColumns.forEach(col => {
        let sum = 0;
        items.forEach(item => {
            const val = getRowValue(item, col);
            if (val) {
                sum += parseFloat(String(val).replace(/,/g, '')) || 0;
            }
        });
        if (Math.abs(sum) > 0.01) {
            console.log(`  - ${col}: ${sum}`);
        }
    });
}
