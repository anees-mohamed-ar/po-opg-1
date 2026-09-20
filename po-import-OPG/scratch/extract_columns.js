const fs = require('fs');
const content = fs.readFileSync('c:/Users/IMT290/Desktop/Projects/po-OPG/po-import-OPG/tallyXMLBuilder.js', 'utf8');

const regex = /getRowValue\s*\(\s*[^,]+,\s*['"]([^'"]+)['"]\s*\)/g;
let match;
const found = new Set();
while ((match = regex.exec(content)) !== null) {
  found.add(match[1]);
}
console.log('Total unique getRowValue column names in tallyXMLBuilder.js:', found.size);
console.log(JSON.stringify(Array.from(found).sort(), null, 2));
