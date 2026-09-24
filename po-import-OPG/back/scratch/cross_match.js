const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const headerDir = 'c:/Users/IMT290/Desktop/Projects/po-OPG/OPG_input_excel_headers';

// Clean helper
const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');

const modules = {
  PO: {
    file: 'OPg_PO_Headers.xlsx',
    programKeys: [
      // Header level
      'Purchasing Document', 'PO -  Doc Type ', 'Document Date ', 'Vendor ', 'Vendor Name  ',
      'Street', 'City  ', 'Post Code', 'GST NO', 'Destination region', 'Region Name',
      // Item level
      'Deletion Indicator', 'Material', 'Short Text', 'Currency', 'Exchange Rate',
      'Order Quantity', 'Order Unit', 'Net Order Price', 'Net Order Value', 'Effective value',
      // Charges / condition columns (standard mm condition types)
      'FRB1', 'FRB2', 'FRC2', 'JEXS', 'NAVS', 'P001', 'P101', 'PB00', 'PBXX  -  Gross Price',
      'R003', 'RA00', 'RA01', 'SKTO', 'WOTB', 'ZBCA', 'ZBCD', 'ZBED', 'ZCEC', 'ZCEQ', 'ZCST',
      'ZDSD', 'ZEQP', 'ZFRQ', 'ZFRV', 'ZHAN', 'ZINS', 'ZLAN', 'ZMFR', 'ZMIS', 'ZNE1', 'ZNE2',
      'ZPAC', 'ZPF%', 'ZPNF', 'ZROY', 'ZRTQ', 'ZRUQ', 'ZSIZ', 'ZSTC', 'ZSTP', 'ZSTV', 'ZVAT',
      'ZVIN', 'ZWRF'
    ]
  },
  GRN: {
    file: 'OPG_GRN_Headers.xlsx',
    programKeys: [
      'Trans./Event Type', 'Trans./Event TypeA', 'Material Document', 'Purchase Order',
      'Document Date', 'Posting Date', 'Vendor', 'Vendor Description', 'Destination region',
      'Deletion Indicator', 'Material', 'Material Description', 'Qty in Un. of Entry',
      'Unit of Entry', 'Amount in LC'
    ]
  },
  Purchase: {
    file: 'OPG_Purchase_Headers.xlsx',
    programKeys: [
      'Invoice No', 'Document Number', 'Reference', 'Doc Date', 'Posting Date',
      'Purchasing Doc Type', 'Invoicing Party', 'Vendor', 'Vendor Name',
      'Street', 'City', 'Post Code', 'GST NO', 'Destination region', 'Region Name',
      'Deletion Indicator', 'Condition Type', 'Material', 'Purchase Order - Short Text',
      'Currency', 'Exchange Rate', 'Quantity', 'Qty in OPUn', 'Order Unit',
      'Amount', 'Total Value', 'Purchasing Document', 'Plant', 'Reference Document',
      'Value-Added Tax Amt'
    ]
  },
  Sales: {
    file: 'OPG_Sales_Headers.xlsx',
    programKeys: [
      'Sales document', 'Billing Document', 'Sales Document Type', 'Sales Document Type Desc',
      'Ship-to party', 'Party', 'Customer', 'Doc Date', 'Created on', 'Document Date',
      'Material', 'Material entered', 'Order Quantity', 'Quantity', 'Base Unit of Measure',
      'Sales unit', 'ZASH - Base Price', 'YBPR- Basic Price', 'YBPR - Basic Price',
      'zpro-base price', 'zpro - base price', 'ZPRS - base price', 'ZPRS - Base Price',
      'YTRL - Transmission', 'CESS', 'ZASH Qty * Base Price', 'YBPR - Qty * Base Price',
      'zpro - Qty * BasePrice', 'ZREL', 'ZOM2 - Without Base amt', 'JWTH',
      'TAX amount', 'TAX Amount', 'Destination region'
    ]
  },
  FI: {
    file: 'OPG_FI_Headers.xlsx',
    programKeys: [
      'Document Number', 'Doc Type', 'Doc Date', 'Posting Date', 'Reference Key',
      'Purchasing Document', 'Debit/Credit Ind.', 'Vendor', 'Customer',
      'G/L Account', 'Account Type', 'Amount', 'Amount in LC'
    ]
  }
};

for (const [mod, def] of Object.entries(modules)) {
  const wb = xlsx.readFile(path.join(headerDir, def.file));
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const allHeaders = xlsx.utils.sheet_to_json(sheet, { header: 1 })[0].filter(Boolean);
  
  console.log(`\n=================== ${mod} ===================`);
  console.log(`Total sample headers: ${allHeaders.length}`);
  
  // Find which ones in allHeaders match the programKeys
  const matchedSampleHeaders = [];
  const unmatchedProgramKeys = [];
  
  for (const pk of def.programKeys) {
    const pkNorm = norm(pk);
    const found = allHeaders.find(h => norm(h) === pkNorm);
    if (found) {
      if (!matchedSampleHeaders.includes(found)) matchedSampleHeaders.push(found);
    } else {
      unmatchedProgramKeys.push(pk);
    }
  }
  
  console.log(`Matched sample headers (${matchedSampleHeaders.length}):`, matchedSampleHeaders);
  if (unmatchedProgramKeys.length > 0) {
    console.log(`Program keys not directly in sample headers (${unmatchedProgramKeys.length}):`, unmatchedProgramKeys);
  }
}
