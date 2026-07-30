const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const { generateTallyXML, getRowValue } = require('../tallyXMLBuilder');
const realExcelPath = "c:\\Users\\IMT290\\Desktop\\Projects\\po-OPG\\Sample_PO_Apr23_Dec23.xlsx";
const workbook = xlsx.readFile(realExcelPath);
const sheet = workbook.Sheets['detail'];
const rawRows = xlsx.utils.sheet_to_json(sheet);

// Build vendor lookup map
const vendorMap = {};
rawRows.forEach(row => {
    const vCode = getRowValue(row, 'Vendor');
    const vName = getRowValue(row, 'Vendor Name');
    if (vCode !== undefined && vCode !== null && vName !== undefined && vName !== null) {
        const cleanCode = String(vCode).split('.')[0].trim();
        vendorMap[cleanCode] = String(vName).trim();
    }
});

const groups = {};
rawRows.forEach(row => {
    const poNum = row['Purchasing Document'];
    if (poNum !== undefined && poNum !== null && String(poNum).trim() !== '') {
        const poKey = String(poNum).split('.')[0].trim();
        if (!groups[poKey]) {
            groups[poKey] = {
                poNumber: poKey,
                items: []
            };
        }
        groups[poKey].items.push(row);
    }
});

const po = groups['4600001071'];
if (po) {
    const xml = generateTallyXML(po, vendorMap);
    console.log("=== XML Ledgers & UDFs for 4600001071 ===");
    const lines = xml.split('\n');
    lines.forEach(line => {
        if (line.includes('<LEDGERNAME>') || line.includes('<AMOUNT>') || line.includes('<UDF:') || line.includes('DESC=')) {
            console.log(line.trim());
        }
    });
}
