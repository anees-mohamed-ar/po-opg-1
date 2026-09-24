const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');
const { generatePurchaseTallyXML } = require('../tallyXMLBuilder');

const excelPath = path.join(__dirname, '../../Imatrix_Purchase Invoice_Apr23_Dec23_Version - 2 with line item (1).XLSX');
if (fs.existsSync(excelPath)) {
    const workbook = xlsx.readFile(excelPath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rawGrid = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    
    // Find headers
    let headerRowIndex = -1;
    for (let i = 0; i < rawGrid.length; i++) {
        const row = rawGrid[i];
        if (row && row.includes('Document Number')) {
            headerRowIndex = i;
            break;
        }
    }
    
    const headers = rawGrid[headerRowIndex].map(h => String(h || '').trim());
    const rawRows = [];
    for (let i = headerRowIndex + 1; i < rawGrid.length; i++) {
        const row = rawGrid[i];
        if (!row || row.length === 0) continue;
        const obj = {};
        headers.forEach((header, colIdx) => {
            if (header) {
                const newVal = row[colIdx];
                const hasCurrent = obj[header] !== undefined && obj[header] !== null && String(obj[header]).trim() !== '';
                const hasNew = newVal !== undefined && newVal !== null && String(newVal).trim() !== '';
                if (hasCurrent && !hasNew) {
                    // Keep current populated value
                } else {
                    obj[header] = newVal;
                }
            }
        });
        rawRows.push(obj);
    }
    
    // Filter rows matching 5105646126
    const matchedRows = rawRows.filter(row => String(row['Document Number']).trim() === '5105646126');
    const group = {
        poNumber: '5105646126',
        items: matchedRows
    };
    
    // Generate Tally XML
    const xml = generatePurchaseTallyXML(group);
    console.log('XML Generated Successfully! Length:', xml.length);
    fs.writeFileSync(path.join(__dirname, 'test_purchase_output.xml'), xml, 'utf8');
    console.log('Saved test XML to scratch/test_purchase_output.xml');
} else {
    console.log('Excel file not found');
}
