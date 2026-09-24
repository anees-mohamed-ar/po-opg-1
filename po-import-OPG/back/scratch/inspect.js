const fs = require('fs');
const path = require('path');

const filePath = path.join(__dirname, '../../po_opg_front/src/App.js');
const content = fs.readFileSync(filePath, 'utf8');
const lines = content.split('\n');

lines.forEach((line, idx) => {
    if (line.includes('importType') || line.includes('setImportType')) {
        console.log(`${idx + 1}: ${line.trim()}`);
    }
});
