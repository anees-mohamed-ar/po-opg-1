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

for (const item of functions) {
    const startIdx = content.indexOf(item.fn);
    const endIdx = content.indexOf(item.endFn, startIdx);
    const fnCode = content.substring(startIdx, endIdx);
    
    // Look for comments, variables, logic
    console.log(`\n================= ${item.name} =================`);
    // Print lines that mention getRowValue
    const lines = fnCode.split('\n').filter(l => l.includes('getRowValue'));
    lines.forEach(l => console.log(l.trim()));
}
