const xlsx = require('xlsx');
const wb = xlsx.readFile('c:/Users/IMT290/Desktop/Projects/po-OPG/Sample_PO_Apr23_Dec23.xlsx');
const data = xlsx.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
const poGroups = {};
data.forEach(r => {
  const type = (r['PO -  Doc Type '] || r['Purchasing Doc. Type'] || '').toString().trim();
  const po = (r['Purchasing Document'] || '').toString().split('.')[0].trim();
  if (type === 'ZCOL' || type.includes('ZCOL')) {
    if (!poGroups[po]) poGroups[po] = [];
    poGroups[po].push(r);
  }
});
const multi = Object.keys(poGroups).filter(po => poGroups[po].length > 1);
console.log('Multi-item ZCOL POs:', multi);
if (multi.length > 0) {
  const po = multi[0];
  console.log('PO:', po, 'Total Lines:', poGroups[po].length);
  poGroups[po].forEach((it, idx) => {
    console.log('Item ' + (idx+1) + ':', JSON.stringify({
      Item: it['Item'],
      Material: it['Material'],
      ShortText: it['Short Text'],
      Vendor: it['Vendor Name  '],
      OrderQty: it['Order Quantity'],
      OrderUnit: it['Order Unit'],
      NetOrderPrice: it['Net Order Price'],
      GrossVal: it['PBXX  -  Gross Price'] || it['Net Order Value'],
      ZCEC: it['ZCEC'],
      JEXS: it['JEXS'],
      EffectiveVal: it['Effective value']
    }, null, 2));
  });
}
