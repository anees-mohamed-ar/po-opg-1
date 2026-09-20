const fs = require('fs');
const content = fs.readFileSync('c:/Users/IMT290/Desktop/Projects/po-OPG/po-import-OPG/tallyXMLBuilder.js', 'utf8');

const functions = [
    { name: 'PO', fn: 'function generateTallyXML', endFn: 'function generatePurchaseTallyXML' },
    { name: 'Purchase', fn: 'function generatePurchaseTallyXML', endFn: 'function generateGRNTallyXML' },
    { name: 'GRN', fn: 'function generateGRNTallyXML', endFn: 'function generateStockJournalTallyXML' },
    { name: 'StockJournal', fn: 'function generateStockJournalTallyXML', endFn: 'function generateSalesOrderTallyXML' },
    { name: 'SalesOrder', fn: 'function generateSalesOrderTallyXML', endFn: 'function generateFITallyXML' },
    { name: 'FI', fn: 'function generateFITallyXML', endFn: 'module.exports' },
];

for (let i = 0; i < functions.length; i++) {
    const item = functions[i];
    const startIdx = content.indexOf(item.fn);
    const endIdx = content.indexOf(item.endFn, startIdx);
    const fnCode = content.substring(startIdx, endIdx);

    const regex = /getRowValue\s*\(\s*[^,]+,\s*['"]([^'"]+)['"]\s*\)/g;
    let match;
    const found = new Set();
    while ((match = regex.exec(fnCode)) !== null) {
        found.add(match[1]);
    }
    console.log(`=== ${item.name} (${found.size}) ===`);
    console.log(JSON.stringify(Array.from(found).sort(), null, 2));
}
