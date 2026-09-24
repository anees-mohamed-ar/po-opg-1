const xlsx = require('xlsx');
const { generateTallyXML, getRowValue } = require('./tallyXMLBuilder');

const realExcelPath = "c:\\Users\\IMT290\\Desktop\\Projects\\po-OPG\\Sample_PO_Apr23_Dec23.xlsx";

try {
    const workbook = xlsx.readFile(realExcelPath);
    let detailSheetName = null;
    for (const name of workbook.SheetNames) {
        if (name.toLowerCase().includes('detail')) {
            detailSheetName = name;
            break;
        }
    }
    
    const sheet = workbook.Sheets[detailSheetName];
    const rawRows = xlsx.utils.sheet_to_json(sheet);
    
    // Group by Purchasing Document
    const groups = {};
    rawRows.forEach(row => {
        // Skip deleted rows
        const delInd = getRowValue(row, 'Deletion Indicator');
        if (delInd && String(delInd).trim().toUpperCase() === 'L') {
            return;
        }

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
    
    for (const poKey of Object.keys(groups)) {
        console.log(`\nPO #${poKey}:`);
        const xml = generateTallyXML(groups[poKey], vendorMap);
        const lines = xml.split('\n');
        
        let foundLedgerEntries = false;
        lines.forEach(line => {
            if (line.includes('<LEDGERNAME>') || line.includes('<AMOUNT>')) {
                // Only print ledger entries (they are at the end, after inventory entries)
                // Let's filter out "Purchase" ledger names
                if (!line.includes('Purchase')) {
                    console.log("  " + line.trim());
                    foundLedgerEntries = true;
                }
            }
        });
    }
} catch (err) {
    console.error(err);
}
