const xlsx = require('xlsx');

const excelPath = "c:\\Users\\IMT290\\Desktop\\Projects\\po-OPG\\imatrix_FI data _2023_3100.XLSX";
console.log("Reading Excel file...");
const workbook = xlsx.readFile(excelPath, { cellDates: true, cellFormulas: false });

const sheetName = workbook.SheetNames[0];
const rows = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

console.log("Total rows in sheet:", rows.length);

// Group by Document Number
const docGroups = {};
rows.forEach(r => {
    const docNum = r['Document Number'] || r['Document No'] || r['Doc Number'] || '';
    if (docNum) {
        if (!docGroups[docNum]) docGroups[docNum] = [];
        docGroups[docNum].push(r);
    }
});

console.log(`Total documents: ${Object.keys(docGroups).length}`);

// Find AB documents where all / multiple entries have Account Type 'K'
const abBothKDocs = [];

for (const [docNum, items] of Object.entries(docGroups)) {
    const docType = items[0]['Doc Type'] || items[0]['Document Type'] || '';
    if (String(docType).trim().toUpperCase() === 'AB') {
        const kItems = items.filter(it => String(it['Account Type']).trim().toUpperCase() === 'K');
        const hasDebitK = kItems.some(it => String(it['Debit/Credit Ind.'] || it['Debit/Credit Ind']).trim().toUpperCase() === 'S');
        const hasCreditK = kItems.some(it => String(it['Debit/Credit Ind.'] || it['Debit/Credit Ind']).trim().toUpperCase() === 'H');
        
        if (hasDebitK && hasCreditK) {
            abBothKDocs.push({ docNum, items, kItems });
        }
    }
}

console.log(`\nFound ${abBothKDocs.length} 'AB' documents where BOTH Debit and Credit have Account Type 'K'`);

abBothKDocs.slice(0, 10).forEach(({ docNum, items }, idx) => {
    console.log(`\n=================== AB BOTH-K DOC #${idx + 1}: ${docNum} (${items.length} lines) ===================`);
    items.forEach((item, lIdx) => {
        console.log(`Line #${lIdx + 1}:`, {
            'Doc Type': item['Doc Type'],
            'Posting Key': item['Posting Key'],
            'Account Type': item['Account Type'],
            'Debit/Credit Ind.': item['Debit/Credit Ind.'] || item['Debit/Credit Ind'],
            'Vendor': item['Vendor'],
            'Vendor Name': item['Vendor Name'] || item['Vendor Description'],
            'G/L Account': item['G/L Account'],
            'G/L Account_1': item['G/L Account_1'],
            'G/L Account Description': item['G/L Account Description'] || item['G/L Account Name'],
            'Amount in LC': item['Amount in LC'],
            'Text': item['Text']
        });
    });
});
