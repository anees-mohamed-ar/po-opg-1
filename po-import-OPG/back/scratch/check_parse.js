const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom'); // We can use jsdom or a simple regex to check what is returned

const filePath = path.join(__dirname, '../../Purchase Order_4600001039_utf8.xml');
const text = fs.readFileSync(filePath, 'utf8');

const { window } = new JSDOM();
const parser = new window.DOMParser();
const xmlDoc = parser.parseFromString(text, 'text/xml');

const voucher = xmlDoc.querySelector('VOUCHER');
if (!voucher) {
  console.log("No VOUCHER node found!");
  process.exit(1);
}

const voucherNum = voucher.querySelector('VOUCHERNUMBER')?.textContent || '';
console.log("Voucher Number:", voucherNum);

const ledgerNodes = xmlDoc.querySelectorAll('LEDGERENTRIES.LIST');
console.log("Found ledger nodes:", ledgerNodes.length);

const ledgerMap = {};
const ledgers = Array.from(ledgerNodes).map(node => {
  const name = node.querySelector('LEDGERNAME')?.textContent || '';
  const rawAmount = node.querySelector('AMOUNT')?.textContent || '0';
  const parsedAmount = parseFloat(rawAmount);
  if (name) {
    ledgerMap[name.trim()] = parsedAmount;
  }
  return {
    name,
    amount: parsedAmount,
    isParty: node.querySelector('ISPARTYLEDGER')?.textContent === 'Yes',
  };
});

console.log("Ledger Map:", ledgerMap);
console.log("Ledgers List:", ledgers);
console.log("Total Amount check:", ledgers.find(l => l.isParty));
