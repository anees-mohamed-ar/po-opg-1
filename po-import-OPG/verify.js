const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const { generateTallyXML, getRowValue } = require('./tallyXMLBuilder');

const realExcelPath = "c:\\Users\\IMT290\\Desktop\\Projects\\po-OPG\\Sample_PO_Apr23_Dec23.xlsx";

console.log("Loading Excel file from:", realExcelPath);

try {
    const workbook = xlsx.readFile(realExcelPath);
    let detailSheetName = null;
    for (const name of workbook.SheetNames) {
        if (name.toLowerCase().includes('detail')) {
            detailSheetName = name;
            break;
        }
    }
    
    console.log("Using sheet:", detailSheetName);
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

    console.log("Total unique POs found:", Object.keys(groups).length);
    
    // Verify PO 4500003402 specifically
    const po4500003402 = groups['4500003402'];
    if (po4500003402) {
        console.log("\nGenerating XML for PO 4500003402 (ZSPR)...");
        const xml = generateTallyXML(po4500003402, vendorMap);
        
        const testOutputPath = path.join(__dirname, 'test_output_4500003266.xml');
        fs.writeFileSync(testOutputPath, xml, 'utf8');
        console.log("Saved generated XML to:", testOutputPath);
        console.log("\nChecking xml structure: first 500 characters:");
        console.log(xml.substring(0, 500));
        console.log("\nChecking inventory item names and ledgers:");
        const lines = xml.split('\n');
        lines.forEach(line => {
            if (line.includes('<STOCKITEMNAME>') || line.includes('<RATE>') || line.includes('<AMOUNT>') || line.includes('<LEDGERNAME>')) {
                console.log(line.trim());
            }
        });
    } else {
        console.log("PO 4500003402 not found in the sheet!");
    }

    // Verify PO 4600001039 (ZCOL) specifically
    const po4600001039 = groups['4600001039'];
    if (po4600001039) {
        console.log("\nGenerating XML for ZCOL PO 4600001039...");
        const xml = generateTallyXML(po4600001039, vendorMap);
        
        const testOutputPathZCOL = path.join(__dirname, 'test_output_ZCOL.xml');
        fs.writeFileSync(testOutputPathZCOL, xml, 'utf8');
        console.log("Saved generated ZCOL XML to:", testOutputPathZCOL);
        console.log("\nChecking ZCOL ledger entries:");
        const lines = xml.split('\n');
        lines.forEach(line => {
            if (line.includes('<LEDGERNAME>') || line.includes('<AMOUNT>') || line.includes('<UDF:')) {
                // filter out purchase ledger to keep output concise
                if (!line.includes('Purchase')) {
                    console.log(line.trim());
                }
            }
        });
    } else {
        console.log("PO 4600001039 not found in the sheet!");
    }
} catch (err) {
    console.error("Verification failed with error:", err);
}
