const xlsx = require('xlsx');
const wb = xlsx.readFile('c:/Users/IMT290/Desktop/Projects/po-OPG/Sample_PO_Apr23_Dec23.xlsx');
const data = xlsx.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
const poGroups = {};
data.forEach(r => {
  const type = (r['PO -  Doc Type '] || r['Purchasing Doc. Type'] || '').toString().trim();
  const po = (r['Purchasing Document'] || '').toString().split('.')[0].trim();
  const del = (r['Deletion Indicator'] || '').toString().trim();
  if ((type === 'ZCOL' || type.includes('ZCOL')) && del !== 'L') {
    if (!poGroups[po]) poGroups[po] = [];
    poGroups[po].push(r);
  }
});

const inrMulti = [];
const usdMulti = [];
const inrSingle = [];
const usdSingle = [];

Object.keys(poGroups).forEach(po => {
  const items = poGroups[po];
  const curr = (items[0]['Currency'] || '').toString().trim().toUpperCase();
  if (items.length > 1) {
    if (curr === 'INR') inrMulti.push(po);
    else if (curr === 'USD') usdMulti.push(po);
  } else {
    if (curr === 'INR') inrSingle.push(po);
    else if (curr === 'USD') usdSingle.push(po);
  }
});

console.log('INR Multi POs (No L):', inrMulti);
console.log('USD Multi POs (No L):', usdMulti);
console.log('INR Single POs (No L):', inrSingle.slice(0, 5));
console.log('USD Single POs (No L):', usdSingle.slice(0, 5));

function dumpPO(poNum) {
  console.log('=== DETAILS FOR PO ' + poNum + ' ===');
  const rows = poGroups[poNum];
  if (!rows) return console.log('PO not found');
  rows.forEach((r, idx) => {
    console.log('--- Line Item ' + (idx + 1) + ' ---');
    for (const k of Object.keys(r)) {
      if (r[k] !== undefined && r[k] !== null && r[k] !== 0 && String(r[k]).trim() !== '') {
        console.log(k + ': ' + r[k]);
      }
    }
  });
}

if (inrMulti.length > 0) dumpPO(inrMulti[0]);
else if (inrSingle.length > 0) dumpPO(inrSingle[0]);

if (usdMulti.length > 0) dumpPO(usdMulti[0]);
else if (usdSingle.length > 0) dumpPO(usdSingle[0]);
