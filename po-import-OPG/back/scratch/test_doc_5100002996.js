const xlsx = require('xlsx');
const path = require('path');
const { generatePurchaseTallyXML, generateFITallyXML } = require('../tallyXMLBuilder');

const filePath = path.join(__dirname, '../../imatrix_FI data _2023_3100.XLSX');
console.log("Reading imatrix_FI data _2023_3100.XLSX...");
const workbook = xlsx.readFile(filePath);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = xlsx.utils.sheet_to_json(sheet);

const rows5100002996 = rows.filter(r => {
    const docNum = r['Document Number'] || r['Document No'] || '';
    return String(docNum).trim() === '5100002996';
});

console.log(`Found ${rows5100002996.length} rows for 5100002996`);

const group = {
    poNumber: '5100002996',
    items: rows5100002996
};

console.log("\n--- Testing generatePurchaseTallyXML ---");
try {
    const purchaseXml = generatePurchaseTallyXML(group);
    console.log("Purchase XML generated successfully! Length:", purchaseXml.length);
} catch (err) {
    console.error("generatePurchaseTallyXML Error:", err);
}

console.log("\n--- Testing generateFITallyXML ---");
try {
    const fiXml = generateFITallyXML(group);
    console.log("FI XML generated successfully! Length:", fiXml.length);
} catch (err) {
    console.error("generateFITallyXML Error:", err);
}
