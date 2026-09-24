const xlsx = require('xlsx');
const path = 'c:\\Users\\IMT290\\Desktop\\Projects\\po-OPG\\GRN_Apr23_Mar23_Posting Date_Imatrix.XLSX';

const workbook = xlsx.readFile(path);
console.log('Sheet Names:', workbook.SheetNames);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rawGrid = xlsx.utils.sheet_to_json(sheet, { header: 1 });

console.log('Total rows in grid:', rawGrid.length);

let headerRowIndex = -1;
for (let i = 0; i < Math.min(rawGrid.length, 50); i++) {
    const row = rawGrid[i];
    if (row && Array.isArray(row)) {
        const hasHeader = row.some(cell => {
            const str = String(cell || '').trim();
            return str === 'Purchasing Document' || 
                   str === 'Document Number' || 
                   str === 'Invoice No' || 
                   str === 'Invoice Number' || 
                   str === 'Material Document' ||
                   str === 'Invoicing Party' ||
                   str === 'Vendor';
        });
        if (hasHeader) {
            headerRowIndex = i;
            break;
        }
    }
}

console.log('Header Row Index:', headerRowIndex);
if (headerRowIndex !== -1) {
    const headers = rawGrid[headerRowIndex].map(h => String(h || '').trim());
    console.log('Headers:', headers);

    const rows = [];
    for (let i = headerRowIndex + 1; i < rawGrid.length; i++) {
        const r = rawGrid[i];
        if (!r || r.length === 0) continue;
        const obj = {};
        headers.forEach((h, idx) => {
            if (h) obj[h] = r[idx];
        });
        rows.push(obj);
    }

    console.log('Total data rows parsed:', rows.length);
    const eventTypes = {};
    rows.forEach(r => {
        const keys = Object.keys(r);
        const eventKey = keys.find(k => k.toLowerCase().includes('event type') || k.toLowerCase().includes('trans'));
        const val = eventKey ? String(r[eventKey]).trim() : 'NONE';
        eventTypes[val] = (eventTypes[val] || 0) + 1;
    });
    console.log('Event types distribution:', eventTypes);

    // Find WA rows
    const waRows = rows.filter(r => {
        const keys = Object.keys(r);
        const eventKey = keys.find(k => k.toLowerCase().includes('event type') || k.toLowerCase().includes('trans'));
        return eventKey && String(r[eventKey]).trim().toUpperCase() === 'WA';
    });

    console.log('Total WA rows found:', waRows.length);
    if (waRows.length > 0) {
        console.log('Sample WA row keys and values:');
        console.log('  Trans./Event Type:', waRows[0]['Trans./Event Type'] || waRows[0]['Trans./Event TypeA']);
        console.log('  Receiving Plant:', waRows[0]['Receiving Plant']);
        console.log('  Plant:', waRows[0]['Plant']);
        console.log('  Material Document:', waRows[0]['Material Document']);
        console.log('All keys of sample WA row:', Object.keys(waRows[0]));
    }
}
