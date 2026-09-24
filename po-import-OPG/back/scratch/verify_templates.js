const { generateTemplateWorkbook, generateTemplateCSV, TEMPLATES_CONFIG } = require('../templatesConfig');
const xlsx = require('xlsx');

for (const key of Object.keys(TEMPLATES_CONFIG)) {
    const buf = generateTemplateWorkbook(key);
    const wb = xlsx.read(buf);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    console.log(`Module [${key}]: ${data[0].length} headers, sample row items: ${data[1].length}`);
    const csv = generateTemplateCSV(key);
    if (!csv.includes('\r\n')) {
        throw new Error(`CSV failed for ${key}`);
    }
}
console.log('All templates generated and verified successfully!');
