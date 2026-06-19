const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

let conditionTypeMap = null;

/**
 * Loads the mapping of condition types to descriptive names from condition_types_Desc_MM.xlsx
 */
function loadConditionTypeMap() {
    if (conditionTypeMap) return conditionTypeMap;
    
    conditionTypeMap = {};
    try {
        const filePath = path.resolve(__dirname, '..', 'condition_types_Desc_MM.xlsx');
        if (fs.existsSync(filePath)) {
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rawRows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
            
            // Row 0 is empty, Row 1 is headers [ 'Application', 'Condition Type', 'Access sequence', 'Name' ]
            for (let i = 2; i < rawRows.length; i++) {
                const row = rawRows[i];
                if (row && row[1] && row[3]) {
                    const condType = String(row[1]).trim().toUpperCase();
                    const name = String(row[3]).trim();
                    conditionTypeMap[condType] = name;
                }
            }
            console.log(`Loaded ${Object.keys(conditionTypeMap).length} condition types from condition_types_Desc_MM.xlsx`);
        } else {
            console.warn(`Condition types Excel file not found at ${filePath}`);
        }
    } catch (err) {
        console.error('Error loading condition types mapping:', err);
    }
    return conditionTypeMap;
}

/**
 * Extract condition code (alphanumeric prefix) from column name
 */
function getConditionCode(str) {
    if (!str) return '';
    const match = String(str).trim().match(/^([a-z0-9%]+)/i);
    return match ? match[1].toUpperCase() : '';
}

/**
 * Robustly retrieve a value from an excel row object using case-insensitive and whitespace-stripped key matching.
 */
function getRowValue(row, colName) {
    if (!row || !colName) return undefined;
    const cleanCol = colName.toLowerCase().replace(/\s/g, '');
    for (const key of Object.keys(row)) {
        if (key.toLowerCase().replace(/\s/g, '') === cleanCol) {
            return row[key];
        }
    }
    // Fallback: match by leading code prefix (e.g. 'JEXS', 'NAVS')
    const codeMatch = colName.match(/^([A-Z0-9]+)/i);
    if (codeMatch) {
        const code = codeMatch[1].toLowerCase();
        for (const key of Object.keys(row)) {
            const keyCodeMatch = key.match(/^([A-Z0-9]+)/i);
            if (keyCodeMatch && keyCodeMatch[1].toLowerCase() === code) {
                return row[key];
            }
        }
    }
    return undefined;
}

/**
 * Escapes special characters for XML compliance
 */
function escapeXML(str) {
    if (str === undefined || str === null) return '';
    return str.toString()
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * Retrieve the exact column name from the row keys matching the target column name case-insensitively and whitespace-stripped.
 */
function getExactKey(row, colName) {
    if (!row || !colName) return colName;
    const cleanCol = colName.toLowerCase().replace(/\s/g, '');
    for (const key of Object.keys(row)) {
        if (key.toLowerCase().replace(/\s/g, '') === cleanCol) {
            return key;
        }
    }
    // Fallback: match by leading code prefix (e.g. 'JEXS', 'NAVS')
    const codeMatch = colName.match(/^([A-Z0-9]+)/i);
    if (codeMatch) {
        const code = codeMatch[1].toLowerCase();
        for (const key of Object.keys(row)) {
            const keyCodeMatch = key.match(/^([A-Z0-9]+)/i);
            if (keyCodeMatch && keyCodeMatch[1].toLowerCase() === code) {
                return key;
            }
        }
    }
    return colName;
}

/**
 * Retrieve the duplicate column's value (vendor code) matching the column name.
 */
function getVendorCodeForRow(row, colName) {
    if (!row || !colName) return undefined;
    const cleanColName = colName.toLowerCase().replace(/\s/g, '');
    for (const key of Object.keys(row)) {
        const cleanKey = key.toLowerCase().replace(/\s/g, '');
        if (cleanKey.includes('vendor') && cleanKey.startsWith(cleanColName)) {
            return row[key];
        }
    }
    for (const key of Object.keys(row)) {
        const cleanKey = key.toLowerCase().replace(/\s/g, '');
        if (cleanKey === cleanColName + '_1' || cleanKey === cleanColName + '1') {
            return row[key];
        }
    }
    const codeMatch = colName.match(/^([A-Z0-9]+)/i);
    if (codeMatch) {
        const code = codeMatch[1].toLowerCase();
        for (const key of Object.keys(row)) {
            const keyCodeMatch = key.match(/^([A-Z0-9]+)(_1|-vendor)/i);
            if (keyCodeMatch && keyCodeMatch[1].toLowerCase() === code) {
                return row[key];
            }
        }
    }
    return undefined;
}

/**
 * Clean state names to match Tally's expected names
 */
function cleanStateName(stateName) {
    if (!stateName) return 'Tamil Nadu';
    const s = stateName.toString().trim().toLowerCase();
    if (s === 'tamilnadu' || s === 'tamil nadu') return 'Tamil Nadu';
    if (s === 'maharashtra') return 'Maharashtra';
    if (s === 'orissa' || s === 'odisha') return 'Odisha';
    // Title case fallback
    return stateName.toString().replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
}

/**
 * Format date to YYYYMMDD
 */
function formatDate(dateValue) {
    if (!dateValue) return '20260401';
    
    // If it's a number or numeric string (Excel serial date)
    if (typeof dateValue === 'number' || (typeof dateValue === 'string' && dateValue.trim() !== '' && !isNaN(dateValue) && !dateValue.includes('.'))) {
        const serial = parseFloat(dateValue);
        const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
        return formatDate(date);
    }
    
    // If it is a Date object
    if (dateValue instanceof Date) {
        const y = dateValue.getFullYear();
        const m = String(dateValue.getMonth() + 1).padStart(2, '0');
        const d = String(dateValue.getDate()).padStart(2, '0');
        return `${y}${m}${d}`;
    }
    
    // If it's a string, try parsing formats like DD.MM.YYYY or YYYY-MM-DD
    const str = String(dateValue).trim();
    if (str.includes('.')) {
        const parts = str.split('.');
        if (parts.length === 3) {
            // Assume DD.MM.YYYY
            const d = parts[0].padStart(2, '0');
            const m = parts[1].padStart(2, '0');
            const y = parts[2];
            return `${y}${m}${d}`;
        }
    }
    
    // Try native date parsing
    try {
        const parsed = new Date(str);
        if (!isNaN(parsed.getTime())) {
            const y = parsed.getFullYear();
            const m = String(parsed.getMonth() + 1).padStart(2, '0');
            const d = String(parsed.getDate()).padStart(2, '0');
            return `${y}${m}${d}`;
        }
    } catch (e) {
        // ignore
    }
    
    return '20260401';
}

/**
 * Pad numbers with leading zeros (e.g. 10353 -> 0000010353)
 */
function padVendor(vendorId) {
    if (vendorId === undefined || vendorId === null) return '';
    const cleanId = String(vendorId).split('.')[0].trim(); // Remove decimals if any
    return cleanId.padStart(10, '0');
}

/**
 * Construct STOCKITEMNAME with specific suffix matching the template for known items, or fallback
 */
function getStockItemName(material, shortText) {
    const matStr = material ? String(material).split('.')[0].trim() : '';
    const textStr = shortText ? String(shortText).trim() : '';
    
    // Use material code directly as the stock item name (acting as an alias in Tally)
    return matStr || textStr || 'UNKNOWN ITEM';
}

/**
 * Generate Tally XML for a grouped PO
 */
function generateTallyXML(poGroup, vendorMap = {}) {
    const condMap = loadConditionTypeMap();
    const firstRow = poGroup.items[0];
    const poNumber = escapeXML(String(getRowValue(firstRow, 'Purchasing Document')).split('.')[0].trim());
    const docType = escapeXML(String(getRowValue(firstRow, 'PO - Doc Type') || 'ZSPR').trim());
    const docDateFormatted = formatDate(getRowValue(firstRow, 'Document Date'));
    
    const vendorName = escapeXML(String(getRowValue(firstRow, 'Vendor Name') || '').trim());
    const vendorCode = padVendor(getRowValue(firstRow, 'Vendor'));
    const rawVendorName = String(getRowValue(firstRow, 'Vendor Name') || '').trim();
    const partyName = escapeXML(vendorCode ? `${vendorCode}-${rawVendorName}` : rawVendorName);
    
    const street = String(getRowValue(firstRow, 'Street') || '').trim();
    const city = String(getRowValue(firstRow, 'City') || '').trim();
    const address = escapeXML(street && city ? `${street},${city}${city}` : (street || city));
    
    const postCode = escapeXML(getRowValue(firstRow, 'Post Code') ? String(getRowValue(firstRow, 'Post Code')).split('.')[0].trim() : '');
    const gstNo = escapeXML(String(getRowValue(firstRow, 'GST NO') || '').trim());
    const regionName = escapeXML(cleanStateName(getRowValue(firstRow, 'Region Name')));
    
    const cmpState = 'Tamil Nadu'; // Default Company state is Tamil Nadu
    const isLocal = regionName.toLowerCase().replace(/\s/g, '') === cmpState.toLowerCase().replace(/\s/g, '');
    
    // Group totals
    let totalNetValue = 0;
    
    const itemsXML = poGroup.items.map(item => {
        const material = getRowValue(item, 'Material');
        const shortText = getRowValue(item, 'Short Text');
        const stockItemName = escapeXML(getStockItemName(material, shortText));
        
        const qty = parseFloat(getRowValue(item, 'Order Quantity')) || 0;
        const unit = escapeXML(String(getRowValue(item, 'Order Unit') || 'Nos').trim());
        const price = parseFloat(getRowValue(item, 'Net Order Price')) || 0;
        const amount = parseFloat(getRowValue(item, 'Net Order Value')) || (qty * price);
        
        totalNetValue += amount;
        
        const rateFormatted = `${price.toFixed(2)}/${unit}`;
        const amountFormatted = `-${amount.toFixed(2)}`;
        const qtyFormatted = ` ${qty} ${unit}`;
        
        const ledgerName = `Purchase ${docType}`;
        
        return `       <ALLINVENTORYENTRIES.LIST>
        <STOCKITEMNAME>${stockItemName}</STOCKITEMNAME>
        <GSTOVRDNINELIGIBLEITC>&#4; Applicable</GSTOVRDNINELIGIBLEITC>
        <GSTOVRDNISREVCHARGEAPPL>&#4; Not Applicable</GSTOVRDNISREVCHARGEAPPL>
        <GSTOVRDNSTOREDNATURE/>
        <GSTRATEINFERAPPLICABILITY>As per Masters/Company</GSTRATEINFERAPPLICABILITY>
        <GSTHSNINFERAPPLICABILITY>As per Masters/Company</GSTHSNINFERAPPLICABILITY>
        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
        <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
        <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
        <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
        <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>
        <ISAUTONEGATE>No</ISAUTONEGATE>
        <ISCUSTOMSCLEARANCE>No</ISCUSTOMSCLEARANCE>
        <ISTRACKCOMPONENT>No</ISTRACKCOMPONENT>
        <ISTRACKPRODUCTION>No</ISTRACKPRODUCTION>
        <ISPRIMARYITEM>No</ISPRIMARYITEM>
        <ISSCRAP>No</ISSCRAP>
        <RATE>${rateFormatted}</RATE>
        <AMOUNT>${amountFormatted}</AMOUNT>
        <ACTUALQTY>${qtyFormatted}</ACTUALQTY>
        <BILLEDQTY>${qtyFormatted}</BILLEDQTY>
        <BATCHALLOCATIONS.LIST>
         <GODOWNNAME>Main Location</GODOWNNAME>
         <BATCHNAME>Primary Batch</BATCHNAME>
         <INDENTNO>&#4; Not Applicable</INDENTNO>
         <ORDERNO>${poNumber}</ORDERNO>
         <TRACKINGNUMBER>&#4; Not Applicable</TRACKINGNUMBER>
         <DYNAMICCSTISCLEARED>No</DYNAMICCSTISCLEARED>
         <AMOUNT>${amountFormatted}</AMOUNT>
         <ACTUALQTY>${qtyFormatted}</ACTUALQTY>
         <BILLEDQTY>${qtyFormatted}</BILLEDQTY>
         <ORDERDUEDATE JD="46112" P="1-Apr-26">1-Apr-26</ORDERDUEDATE>
         <ADDITIONALDETAILS.LIST>        </ADDITIONALDETAILS.LIST>
         <VOUCHERCOMPONENTLIST.LIST>        </VOUCHERCOMPONENTLIST.LIST>
        </BATCHALLOCATIONS.LIST>
        <ACCOUNTINGALLOCATIONS.LIST>
         <OLDAUDITENTRYIDS.LIST TYPE="Number">
          <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
         </OLDAUDITENTRYIDS.LIST>
         <LEDGERNAME>${ledgerName}</LEDGERNAME>
         <GSTCLASS>&#4; Not Applicable</GSTCLASS>
         <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
         <LEDGERFROMITEM>No</LEDGERFROMITEM>
         <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
         <ISPARTYLEDGER>No</ISPARTYLEDGER>
         <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
         <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
         <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
         <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
         <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
         <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
         <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>
         <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
         <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
         <AMOUNT>${amountFormatted}</AMOUNT>
         <SERVICETAXDETAILS.LIST>        </SERVICETAXDETAILS.LIST>
         <BANKALLOCATIONS.LIST>        </BANKALLOCATIONS.LIST>
         <BILLALLOCATIONS.LIST>        </BILLALLOCATIONS.LIST>
         <INTERESTCOLLECTION.LIST>        </INTERESTCOLLECTION.LIST>
         <OLDAUDITENTRIES.LIST>        </OLDAUDITENTRIES.LIST>
         <ACCOUNTAUDITENTRIES.LIST>        </ACCOUNTAUDITENTRIES.LIST>
         <AUDITENTRIES.LIST>        </AUDITENTRIES.LIST>
         <INPUTCRALLOCS.LIST>        </INPUTCRALLOCS.LIST>
         <DUTYHEADDETAILS.LIST>        </DUTYHEADDETAILS.LIST>
         <EXCISEDUTYHEADDETAILS.LIST>        </EXCISEDUTYHEADDETAILS.LIST>
         <RATEDETAILS.LIST>        </RATEDETAILS.LIST>
         <SUMMARYALLOCS.LIST>        </SUMMARYALLOCS.LIST>
         <CENVATDUTYALLOCATIONS.LIST>        </CENVATDUTYALLOCATIONS.LIST>
         <STPYMTDETAILS.LIST>        </STPYMTDETAILS.LIST>
         <EXCISEPAYMENTALLOCATIONS.LIST>        </EXCISEPAYMENTALLOCATIONS.LIST>
         <TAXBILLALLOCATIONS.LIST>        </TAXBILLALLOCATIONS.LIST>
         <TAXOBJECTALLOCATIONS.LIST>        </TAXOBJECTALLOCATIONS.LIST>
         <TDSEXPENSEALLOCATIONS.LIST>        </TDSEXPENSEALLOCATIONS.LIST>
         <VATSTATUTORYDETAILS.LIST>        </VATSTATUTORYDETAILS.LIST>
         <COSTTRACKALLOCATIONS.LIST>        </COSTTRACKALLOCATIONS.LIST>
         <REFVOUCHERDETAILS.LIST>        </REFVOUCHERDETAILS.LIST>
         <INVOICEWISEDETAILS.LIST>        </INVOICEWISEDETAILS.LIST>
         <VATITCDETAILS.LIST>        </VATITCDETAILS.LIST>
         <ADVANCETAXDETAILS.LIST>        </ADVANCETAXDETAILS.LIST>
         <TAXTYPEALLOCATIONS.LIST>        </TAXTYPEALLOCATIONS.LIST>
        </ACCOUNTINGALLOCATIONS.LIST>
        <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>
        <RATEDETAILS.LIST>
         <GSTRATEDUTYHEAD>CGST</GSTRATEDUTYHEAD>
        </RATEDETAILS.LIST>
        <RATEDETAILS.LIST>
         <GSTRATEDUTYHEAD>SGST/UTGST</GSTRATEDUTYHEAD>
        </RATEDETAILS.LIST>
        <RATEDETAILS.LIST>
         <GSTRATEDUTYHEAD>IGST</GSTRATEDUTYHEAD>
        </RATEDETAILS.LIST>
        <RATEDETAILS.LIST>
         <GSTRATEDUTYHEAD>Cess</GSTRATEDUTYHEAD>
        </RATEDETAILS.LIST>
        <RATEDETAILS.LIST>
         <GSTRATEDUTYHEAD>State Cess</GSTRATEDUTYHEAD>
        </RATEDETAILS.LIST>
        <SUPPLEMENTARYDUTYHEADDETAILS.LIST>       </SUPPLEMENTARYDUTYHEADDETAILS.LIST>
        <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>
        <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>
        <EXCISEALLOCATIONS.LIST>       </EXCISEALLOCATIONS.LIST>
        <EXPENSEALLOCATIONS.LIST>       </EXPENSEALLOCATIONS.LIST>
       </ALLINVENTORYENTRIES.LIST>`;
    }).join('\n');
    
    // Ledger columns that need to be parsed and created as separate ledgers if they contain non-zero value
    const ledgerColumns = [
        'FRB1', 'FRB2', 'FRC2', 'JEXS', 'NAVS', 'P001', 'P101', 'PB00', 'PBXX  -  Gross Price', 'R003', 'RA00', 'RA01', 'SKTO', 'WOTB',
        'ZBCA', 'ZBCD', 'ZBED', 'ZCEC', 'ZCEQ', 'ZCST', 'ZDSD', 'ZEQP', 'ZFRQ', 'ZFRV', 'ZHAN', 'ZINS', 'ZLAN', 'ZMFR', 'ZMIS', 'ZNE1', 'ZNE2', 'ZPAC', 'ZPF%', 'ZPNF', 'ZROY', 'ZRTQ', 'ZRUQ', 'ZSIZ', 'ZSTC', 'ZSTP', 'ZSTV', 'ZVAT', 'ZVIN', 'ZWRF'
    ];
    
    const udfSlots = [
        { key: 'TransOcFre', inr: 'TRANSOCFREINR', str: 'TRANSOCFRE', sub: 'TRANSOCFRESUB', idxInr: '2088', idxStr: '2086', idxSub: '2087' },
        { key: 'TransDDwgt', inr: 'TRANSDDWGTINR', str: 'TRANSDDWGT', sub: 'TRANSDDWGTSUB', idxInr: '2091', idxStr: '2089', idxSub: '2090' },
        { key: 'TransDem', inr: 'TRANSDEMINR', str: 'TRANSDEM', sub: 'TRANSDEMSUB', idxInr: '2094', idxStr: '2092', idxSub: '2093' },
        { key: 'TransInsur', inr: 'TRANSINSURINR', str: 'TRANSINSUR', sub: 'TRANSINSURSUB', idxInr: '2097', idxStr: '2095', idxSub: '2096' },
        { key: 'TransSample', inr: 'TRANSSAMPLEINR', str: 'TRANSSAMPLE', sub: 'TRANSSAMPLESUB', idxInr: '2100', idxStr: '20981', idxSub: '2099' },
        { key: 'TransSTock', inr: 'TRANSSTOCKINR', str: 'TRANSSTOCK', sub: 'TRANSSTOCKSUB', idxInr: '2103', idxStr: '2101', idxSub: '2102' },
        { key: 'TransDraft', inr: 'TRANSDRAFTINR', str: 'TRANSDRAFT', sub: 'TRANSDRAFTSUB', idxInr: '2106', idxStr: '2104', idxSub: '2105' },
        { key: 'TransLia', inr: 'TRANSLIAINR', str: 'TRANSLIA', sub: 'TRANSLIASUB', idxInr: '2109', idxStr: '2107', idxSub: '2108' },
        { key: 'TransWhar', inr: 'TRANSWHARINR', str: 'TRANSWHAR', sub: 'TRANSWHARSUB', idxInr: '2112', idxStr: '2110', idxSub: '2111' },
        { key: 'TransSample2', inr: 'TRANSSAMPLE2INR', str: 'TRANSSAMPLE2', sub: 'TRANSSAMPLE2SUB', idxInr: '2146', idxStr: '2144', idxSub: '2145' }
    ];
    
    let udfCount = 0;
    const udfXmls = [];
    const activeLedgers = [];
    let totalTaxesAndCharges = 0;
    
    ledgerColumns.forEach(col => {
        let colSum = 0;
        poGroup.items.forEach(item => {
            const rawVal = getRowValue(item, col);
            if (rawVal !== undefined && rawVal !== null) {
                const cleanVal = String(rawVal).replace(/,/g, '').trim();
                colSum += parseFloat(cleanVal) || 0;
            }
        });
        
        if (Math.abs(colSum) > 0.001) {
            const exactKey = getExactKey(firstRow, col);
            const condCode = getConditionCode(col);
            const mappedName = condMap[condCode];
            
            const finalLedgerName = mappedName ? mappedName : exactKey.toString().replace(/\s+/g, ' ').trim();
            if (finalLedgerName.toLowerCase().replace(/\s/g, '') === 'grossprice') {
                return;
            }
            const normalizedName = escapeXML(finalLedgerName);
            const cleanColName = col.toLowerCase().replace(/\s/g, '');
            
            if (docType === 'ZCOL' && cleanColName !== 'jexs' && cleanColName !== 'navs') {
                if (udfCount < udfSlots.length) {
                    const slot = udfSlots[udfCount];
                    udfCount++;
                    
                    let vendorCode = '';
                    for (const item of poGroup.items) {
                        const vCode = getVendorCodeForRow(item, col);
                        if (vCode !== undefined && vCode !== null && String(vCode).trim() !== '' && parseFloat(vCode) !== 0) {
                            vendorCode = String(vCode).split('.')[0].trim();
                            break;
                        }
                    }
                    if (!vendorCode && poGroup.items.length > 0) {
                        const vCode = getVendorCodeForRow(poGroup.items[0], col);
                        if (vCode !== undefined && vCode !== null) {
                            vendorCode = String(vCode).split('.')[0].trim();
                        }
                    }
                    
                    let vendorString = 'Unknown';
                    if (vendorCode && vendorCode !== '0' && vendorCode !== '0000000000') {
                        const paddedCode = padVendor(vendorCode);
                        const nameFromMap = vendorMap[vendorCode];
                        vendorString = nameFromMap ? `${paddedCode}-${nameFromMap}` : paddedCode;
                    }
                    
                    const escAmount = colSum.toFixed(2);
                    const escVendor = escapeXML(vendorString);
                    const escColName = normalizedName;
                    
                    const slotXml = `      <UDF:${slot.inr}.LIST DESC="\`${slot.key}Inr\`" ISLIST="YES" TYPE="Amount" INDEX="${slot.idxInr}">
       <UDF:${slot.inr} DESC="\`${slot.key}Inr\`">${escAmount}</UDF:${slot.inr}>
      </UDF:${slot.inr}.LIST>
      <UDF:${slot.sub}.LIST DESC="\`${slot.key}Sub\`" ISLIST="YES" TYPE="String" INDEX="${slot.idxSub}">
       <UDF:${slot.sub} DESC="\`${slot.key}Sub\`">${escVendor}</UDF:${slot.sub}>
      </UDF:${slot.sub}.LIST>
      <UDF:${slot.str}.LIST DESC="\`${slot.key}\`" ISLIST="YES" TYPE="String" INDEX="${slot.idxStr}">
       <UDF:${slot.str} DESC="\`${slot.key}\`">${escColName}</UDF:${slot.str}>
      </UDF:${slot.str}.LIST>`;
                    udfXmls.push(slotXml);
                }
            } else {
                activeLedgers.push({
                    name: normalizedName,
                    sum: colSum
                });
                totalTaxesAndCharges += colSum;
            }
        }
    });
    
    const totalVoucherAmount = totalNetValue + totalTaxesAndCharges;
    
    const taxLedgerXML = activeLedgers.map(led => {
        const isDeemedPositive = led.sum > 0 ? 'Yes' : 'No';
        // Positive columns are debits (negative in XML), negative columns like discount are credits (positive in XML)
        const formattedAmount = led.sum > 0 ? `-${led.sum.toFixed(2)}` : `${Math.abs(led.sum).toFixed(2)}`;
        
        return `       <LEDGERENTRIES.LIST>
        <OLDAUDITENTRYIDS.LIST TYPE="Number">
         <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
        </OLDAUDITENTRYIDS.LIST>
        <APPROPRIATEFOR>&#4; Not Applicable</APPROPRIATEFOR>
        <LEDGERNAME>${led.name}</LEDGERNAME>
        <GSTCLASS>&#4; Not Applicable</GSTCLASS>
        <ISDEEMEDPOSITIVE>${isDeemedPositive}</ISDEEMEDPOSITIVE>
        <LEDGERFROMITEM>No</LEDGERFROMITEM>
        <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
        <ISPARTYLEDGER>No</ISPARTYLEDGER>
        <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
        <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
        <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
        <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
        <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
        <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
        <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>
        <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
        <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
        <AMOUNT>${formattedAmount}</AMOUNT>
        <VATEXPAMOUNT>${formattedAmount}</VATEXPAMOUNT>
        <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>
        <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>
        <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>
        <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>
        <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>
        <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>
        <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>
        <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>
        <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>
        <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>
        <RATEDETAILS.LIST>       </RATEDETAILS.LIST>
        <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>
        <CENVATDUTYALLOCATIONS.LIST>       </CENVATDUTYALLOCATIONS.LIST>
        <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>
        <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>
        <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>
        <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>
        <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>
        <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>
        <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>
        <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>
        <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>
        <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>
        <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>
        <TAXTYPEALLOCATIONS.LIST>       </TAXTYPEALLOCATIONS.LIST>
       </LEDGERENTRIES.LIST>`;
    }).join('\n');
    
    // Generate UUID or REMOTEID format matching "2786887a-d92f-46c7-bf13-d8a373da8523-00000001"
    const remoteId = `2786887a-d92f-46c7-bf13-d8a373da8523-${poNumber.padStart(8, '0')}`;
    const vchKey = `2786887a-d92f-46c7-bf13-d8a373da8523-0000b420:${poNumber.padStart(8, '0')}`;

    return `<ENVELOPE>
 <HEADER>
  <TALLYREQUEST>Import Data</TALLYREQUEST>
 </HEADER>
 <BODY>
  <IMPORTDATA>
   <REQUESTDESC>
    <REPORTNAME>Vouchers</REPORTNAME>
    <STATICVARIABLES>
     <SVCURRENTCOMPANY>OPG Master</SVCURRENTCOMPANY>
    </STATICVARIABLES>
   </REQUESTDESC>
   <REQUESTDATA>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <VOUCHER REMOTEID="${remoteId}" VCHKEY="${vchKey}" VCHTYPE="Purcahse Order ${docType}" ACTION="Create" OBJVIEW="Invoice Voucher View">
      <ADDRESS.LIST TYPE="String">
       <ADDRESS>${address}</ADDRESS>
      </ADDRESS.LIST>
      <OLDAUDITENTRYIDS.LIST TYPE="Number">
       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
      </OLDAUDITENTRYIDS.LIST>
      <DATE>${docDateFormatted}</DATE>
      <VCHSTATUSDATE>${docDateFormatted}</VCHSTATUSDATE>
      <GUID>${remoteId}</GUID>
      <GSTREGISTRATIONTYPE>&#4; Unknown</GSTREGISTRATIONTYPE>
      <VATDEALERTYPE>&#4; Unknown</VATDEALERTYPE>
      <STATENAME>${regionName}</STATENAME>
      <OBJECTUPDATEACTION/>
      <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
      <PARTYGSTIN>${gstNo}</PARTYGSTIN>
      <PLACEOFSUPPLY>${regionName}</PLACEOFSUPPLY>
      <VOUCHERTYPENAME>Purcahse Order ${docType}</VOUCHERTYPENAME>
      <PARTYNAME>${partyName}</PARTYNAME>
      <GSTREGISTRATION TAXTYPE="GST" TAXREGISTRATION="">${regionName} Registration</GSTREGISTRATION>
      <PARTYLEDGERNAME>${partyName}</PARTYLEDGERNAME>
      <VOUCHERNUMBER>${poNumber}</VOUCHERNUMBER>
      <BASICBUYERNAME>OPG Master</BASICBUYERNAME>
      <CMPGSTREGISTRATIONTYPE>Regular</CMPGSTREGISTRATIONTYPE>
      <REFERENCE>${poNumber}</REFERENCE>
      <PARTYMAILINGNAME>${vendorName}</PARTYMAILINGNAME>
      <PARTYPINCODE>${postCode}</PARTYPINCODE>
      <CONSIGNEEMAILINGNAME>OPG Master</CONSIGNEEMAILINGNAME>
      <CONSIGNEESTATENAME>${cmpState}</CONSIGNEESTATENAME>
      <CMPGSTSTATE>${cmpState}</CMPGSTSTATE>
      <CONSIGNEECOUNTRYNAME>India</CONSIGNEECOUNTRYNAME>
      <BASICBASEPARTYNAME>${partyName}</BASICBASEPARTYNAME>
      <NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
      <CSTFORMISSUETYPE>&#4; Not Applicable</CSTFORMISSUETYPE>
      <CSTFORMRECVTYPE>&#4; Not Applicable</CSTFORMRECVTYPE>
      <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>
      <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
      <VCHSTATUSTAXADJUSTMENT>Default</VCHSTATUSTAXADJUSTMENT>
      <VCHSTATUSVOUCHERTYPE>Purcahse Order ${docType}</VCHSTATUSVOUCHERTYPE>
      <VCHSTATUSTAXUNIT>${cmpState} Registration</VCHSTATUSTAXUNIT>
      <VCHGSTCLASS>&#4; Not Applicable</VCHGSTCLASS>
      <BUYERPINNUMBER>${gstNo.length >= 12 ? gstNo.substring(2, 12) : ''}</BUYERPINNUMBER>
      <DIFFACTUALQTY>No</DIFFACTUALQTY>
      <ISMSTFROMSYNC>No</ISMSTFROMSYNC>
      <ISDELETED>No</ISDELETED>
      <ISSECURITYONWHENENTERED>No</ISSECURITYONWHENENTERED>
      <ASORIGINAL>No</ASORIGINAL>
      <AUDITED>No</AUDITED>
      <ISCOMMONPARTY>No</ISCOMMONPARTY>
      <FORJOBCOSTING>No</FORJOBCOSTING>
      <ISOPTIONAL>No</ISOPTIONAL>
      <EFFECTIVEDATE>${docDateFormatted}</EFFECTIVEDATE>
      <USEFOREXCISE>No</USEFOREXCISE>
      <ISFORJOBWORKIN>No</ISFORJOBWORKIN>
      <ALLOWCONSUMPTION>No</ALLOWCONSUMPTION>
      <USEFORINTEREST>No</USEFORINTEREST>
      <USEFORGAINLOSS>No</USEFORGAINLOSS>
      <USEFORGODOWNTRANSFER>No</USEFORGODOWNTRANSFER>
      <USEFORCOMPOUND>No</USEFORCOMPOUND>
      <USEFORSERVICETAX>No</USEFORSERVICETAX>
      <ISREVERSECHARGEAPPLICABLE>No</ISREVERSECHARGEAPPLICABLE>
      <ISSYSTEM>No</ISSYSTEM>
      <ISFETCHEDONLY>No</ISFETCHEDONLY>
      <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>
      <ISCANCELLED>No</ISCANCELLED>
      <ISONHOLD>No</ISONHOLD>
      <ISSUMMARY>No</ISSUMMARY>
      <ISECOMMERCESUPPLY>No</ISECOMMERCESUPPLY>
      <ISBOENOTAPPLICABLE>No</ISBOENOTAPPLICABLE>
      <ISGSTSECSEVENAPPLICABLE>No</ISGSTSECSEVENAPPLICABLE>
      <IGNOREEINVVALIDATION>No</IGNOREEINVVALIDATION>
      <CMPGSTISOTHTERRITORYASSESSEE>No</CMPGSTISOTHTERRITORYASSESSEE>
      <PARTYGSTISOTHTERRITORYASSESSEE>No</PARTYGSTISOTHTERRITORYASSESSEE>
      <IRNJSONEXPORTED>No</IRNJSONEXPORTED>
      <IRNCANCELLED>No</IRNCANCELLED>
      <IGNOREGSTCONFLICTINMIG>No</IGNOREGSTCONFLICTINMIG>
      <ISOPBALTRANSACTION>No</ISOPBALTRANSACTION>
      <IGNOREGSTFORMATVALIDATION>No</IGNOREGSTFORMATVALIDATION>
      <ISELIGIBLEFORITC>Yes</ISELIGIBLEFORITC>
      <IGNOREGSTOPTIONALUNCERTAIN>No</IGNOREGSTOPTIONALUNCERTAIN>
      <UPDATESUMMARYVALUES>No</UPDATESUMMARYVALUES>
      <ISEWAYBILLAPPLICABLE>No</ISEWAYBILLAPPLICABLE>
      <ISDELETEDRETAINED>No</ISDELETEDRETAINED>
      <ISNULL>No</ISNULL>
      <ISEXCISEVOUCHER>No</ISEXCISEVOUCHER>
      <EXCISETAXOVERRIDE>No</EXCISETAXOVERRIDE>
      <USEFORTAXUNITTRANSFER>No</USEFORTAXUNITTRANSFER>
      <ISEXER1NOPOVERWRITE>No</ISEXER1NOPOVERWRITE>
      <ISEXF2NOPOVERWRITE>No</ISEXF2NOPOVERWRITE>
      <ISEXER3NOPOVERWRITE>No</ISEXER3NOPOVERWRITE>
      <IGNOREPOSVALIDATION>No</IGNOREPOSVALIDATION>
      <EXCISEOPENING>No</EXCISEOPENING>
      <USEFORFINALPRODUCTION>No</USEFORFINALPRODUCTION>
      <ISTDSOVERRIDDEN>No</ISTDSOVERRIDDEN>
      <ISTCSOVERRIDDEN>No</ISTCSOVERRIDDEN>
      <ISTDSTCSCASHVCH>No</ISTDSTCSCASHVCH>
      <INCLUDEADVPYMTVCH>No</INCLUDEADVPYMTVCH>
      <ISSUBWORKSCONTRACT>No</ISSUBWORKSCONTRACT>
      <ISVATOVERRIDDEN>No</ISVATOVERRIDDEN>
      <IGNOREORIGVCHDATE>No</IGNOREORIGVCHDATE>
      <ISVATPAIDATCUSTOMS>No</ISVATPAIDATCUSTOMS>
      <ISDECLAREDTOCUSTOMS>No</ISDECLAREDTOCUSTOMS>
      <VATADVANCEPAYMENT>No</VATADVANCEPAYMENT>
      <VATADVPAY>No</VATADVPAY>
      <ISCSTDELCAREDGOODSSALES>No</ISCSTDELCAREDGOODSSALES>
      <ISVATRESTAXINV>No</ISVATRESTAXINV>
      <ISSERVICETAXOVERRIDDEN>No</ISSERVICETAXOVERRIDDEN>
      <ISISDVOUCHER>No</ISISDVOUCHER>
      <ISEXCISEOVERRIDDEN>No</ISEXCISEOVERRIDDEN>
      <ISEXCISESUPPLYVCH>No</ISEXCISESUPPLYVCH>
      <GSTNOTEXPORTED>No</GSTNOTEXPORTED>
      <IGNOREGSTINVALIDATION>No</IGNOREGSTINVALIDATION>
      <ISGSTREFUND>No</ISGSTREFUND>
      <OVRDNEWAYBILLAPPLICABILITY>No</OVRDNEWAYBILLAPPLICABILITY>
      <ISVATPRINCIPALACCOUNT>No</ISVATPRINCIPALACCOUNT>
      <VCHSTATUSISVCHNUMUSED>No</VCHSTATUSISVCHNUMUSED>
      <VCHGSTSTATUSISINCLUDED>No</VCHGSTSTATUSISINCLUDED>
      <VCHGSTSTATUSISUNCERTAIN>No</VCHGSTSTATUSISUNCERTAIN>
      <VCHGSTSTATUSISEXCLUDED>No</VCHGSTSTATUSISEXCLUDED>
      <VCHGSTSTATUSISAPPLICABLE>No</VCHGSTSTATUSISAPPLICABLE>
      <VCHGSTSTATUSISGSTR2BRECONCILED>No</VCHGSTSTATUSISGSTR2BRECONCILED>
      <VCHGSTSTATUSISGSTR2BONLYINPORTAL>No</VCHGSTSTATUSISGSTR2BONLYINPORTAL>
      <VCHGSTSTATUSISGSTR2BONLYINBOOKS>No</VCHGSTSTATUSISGSTR2BONLYINBOOKS>
      <VCHGSTSTATUSISGSTR2BMISMATCH>No</VCHGSTSTATUSISGSTR2BMISMATCH>
      <VCHGSTSTATUSISGSTR2BINDIFFPERIOD>No</VCHGSTSTATUSISGSTR2BINDIFFPERIOD>
      <VCHGSTSTATUSISRETEFFDATEOVERRDN>No</VCHGSTSTATUSISRETEFFDATEOVERRDN>
      <VCHGSTSTATUSISOVERRDN>No</VCHGSTSTATUSISOVERRDN>
      <VCHGSTSTATUSISSTATINDIFFDATE>No</VCHGSTSTATUSISSTATINDIFFDATE>
      <VCHGSTSTATUSISRETINDIFFDATE>No</VCHGSTSTATUSISRETINDIFFDATE>
      <VCHGSTSTATUSMAINSECTIONEXCLUDED>No</VCHGSTSTATUSMAINSECTIONEXCLUDED>
      <VCHGSTSTATUSISBRANCHTRANSFEROUT>No</VCHGSTSTATUSISBRANCHTRANSFEROUT>
      <VCHGSTSTATUSISSYSTEMSUMMARY>No</VCHGSTSTATUSISSYSTEMSUMMARY>
      <VCHSTATUSISUNREGISTEREDRCM>No</VCHSTATUSISUNREGISTEREDRCM>
      <VCHSTATUSISOPTIONAL>No</VCHSTATUSISOPTIONAL>
      <VCHSTATUSISCANCELLED>No</VCHSTATUSISCANCELLED>
      <VCHSTATUSISDELETED>No</VCHSTATUSISDELETED>
      <VCHSTATUSISOPENINGBALANCE>No</VCHSTATUSISOPENINGBALANCE>
      <VCHSTATUSISFETCHEDONLY>No</VCHSTATUSISFETCHEDONLY>
      <VCHGSTSTATUSISOPTIONALUNCERTAIN>No</VCHGSTSTATUSISOPTIONALUNCERTAIN>
      <VCHSTATUSISREACCEPTFORHSNDONE>No</VCHSTATUSISREACCEPTFORHSNDONE>
      <VCHSTATUSISREACCEPHSNSIXONEDONE>No</VCHSTATUSISREACCEPHSNSIXONEDONE>
      <PAYMENTLINKHASMULTIREF>No</PAYMENTLINKHASMULTIREF>
      <ISSHIPPINGWITHINSTATE>No</ISSHIPPINGWITHINSTATE>
      <ISOVERSEASTOURISTTRANS>No</ISOVERSEASTOURISTTRANS>
      <ISDESIGNATEDZONEPARTY>No</ISDESIGNATEDZONEPARTY>
      <HASCASHFLOW>No</HASCASHFLOW>
      <ISPOSTDATED>No</ISPOSTDATED>
      <USETRACKINGNUMBER>No</USETRACKINGNUMBER>
      <ISINVOICE>No</ISINVOICE>
      <MFGJOURNAL>No</MFGJOURNAL>
      <HASDISCOUNTS>No</HASDISCOUNTS>
      <ASPAYSLIP>No</ASPAYSLIP>
      <ISCOSTCENTRE>No</ISCOSTCENTRE>
      <ISSTXNONREALIZEDVCH>No</ISSTXNONREALIZEDVCH>
      <ISEXCISEMANUFACTURERON>No</ISEXCISEMANUFACTURERON>
      <ISBLANKCHEQUE>No</ISBLANKCHEQUE>
      <ISVOID>No</ISVOID>
      <ORDERLINESTATUS>No</ORDERLINESTATUS>
      <VATISAGNSTCANCSALES>No</VATISAGNSTCANCSALES>
      <VATISPURCEXEMPTED>No</VATISPURCEXEMPTED>
      <ISVATRESTAXINVOICE>No</ISVATRESTAXINVOICE>
      <VATISASSESABLECALCVCH>No</VATISASSESABLECALCVCH>
      <ISVATDUTYPAID>Yes</ISVATDUTYPAID>
      <ISDELIVERYSAMEASCONSIGNEE>No</ISDELIVERYSAMEASCONSIGNEE>
      <ISDISPATCHSAMEASCONSIGNOR>No</ISDISPATCHSAMEASCONSIGNOR>
      <ISDELETEDVCHRETAINED>No</ISDELETEDVCHRETAINED>
      <VCHONLYADDLINFOUPDATED>No</VCHONLYADDLINFOUPDATED>
      <CHANGEVCHMODE>No</CHANGEVCHMODE>
      <RESETIRNQRCODE>No</RESETIRNQRCODE>
      <ALTERID> 4</ALTERID>
      <MASTERID> 1</MASTERID>
      <VOUCHERKEY>198049531953160</VOUCHERKEY>
      <VOUCHERRETAINKEY>1</VOUCHERRETAINKEY>
      <VOUCHERNUMBERSERIES>Default</VOUCHERNUMBERSERIES>
      <UPDATEDDATETIME>20260603115029000</UPDATEDDATETIME>
      <EWAYBILLDETAILS.LIST>      </EWAYBILLDETAILS.LIST>
      <EXCLUDEDTAXATIONS.LIST>      </EXCLUDEDTAXATIONS.LIST>
      <OLDAUDITENTRIES.LIST>      </OLDAUDITENTRIES.LIST>
      <ACCOUNTAUDITENTRIES.LIST>      </ACCOUNTAUDITENTRIES.LIST>
      <AUDITENTRIES.LIST>      </AUDITENTRIES.LIST>
      <DUTYHEADDETAILS.LIST>      </DUTYHEADDETAILS.LIST>
      <GSTADVADJDETAILS.LIST>      </GSTADVADJDETAILS.LIST>
${itemsXML}
      <CONTRITRANS.LIST>      </CONTRITRANS.LIST>
      <EWAYBILLERRORLIST.LIST>      </EWAYBILLERRORLIST.LIST>
      <IRNERRORLIST.LIST>      </IRNERRORLIST.LIST>
      <HARYANAVAT.LIST>      </HARYANAVAT.LIST>
      <SUPPLEMENTARYDUTYHEADDETAILS.LIST>      </SUPPLEMENTARYDUTYHEADDETAILS.LIST>
      <INVOICEDELNOTES.LIST>      </INVOICEDELNOTES.LIST>
      <INVOICEORDERLIST.LIST>      </INVOICEORDERLIST.LIST>
      <INVOICEINDENTLIST.LIST>      </INVOICEINDENTLIST.LIST>
      <ATTENDANCEENTRIES.LIST>      </ATTENDANCEENTRIES.LIST>
      <ORIGINVOICEDETAILS.LIST>      </ORIGINVOICEDETAILS.LIST>
      <INVOICEEXPORTLIST.LIST>      </INVOICEEXPORTLIST.LIST>
      <LEDGERENTRIES.LIST>
        <OLDAUDITENTRYIDS.LIST TYPE="Number">
         <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
        </OLDAUDITENTRYIDS.LIST>
        <LEDGERNAME>${partyName}</LEDGERNAME>
        <GSTCLASS>&#4; Not Applicable</GSTCLASS>
        <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
        <LEDGERFROMITEM>No</LEDGERFROMITEM>
        <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
        <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
        <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
        <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
        <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
        <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
        <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
        <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
        <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
        <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
        <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
        <AMOUNT>${totalVoucherAmount.toFixed(2)}</AMOUNT>
        <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>
        <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>
        <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>
        <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>
        <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>
        <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>
        <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>
        <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>
        <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>
        <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>
        <RATEDETAILS.LIST>       </RATEDETAILS.LIST>
        <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>
        <CENVATDUTYALLOCATIONS.LIST>       </CENVATDUTYALLOCATIONS.LIST>
        <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>
        <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>
        <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>
        <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>
        <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>
        <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>
        <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>
        <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>
        <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>
        <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>
        <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>
        <TAXTYPEALLOCATIONS.LIST>       </TAXTYPEALLOCATIONS.LIST>
      </LEDGERENTRIES.LIST>
${taxLedgerXML}
      <GST.LIST>      </GST.LIST>
      <STKJRNLADDLCOSTDETAILS.LIST>      </STKJRNLADDLCOSTDETAILS.LIST>
      <PAYROLLMODEOFPAYMENT.LIST>      </PAYROLLMODEOFPAYMENT.LIST>
      <ATTDRECORDS.LIST>      </ATTDRECORDS.LIST>
      <GSTEWAYCONSIGNORADDRESS.LIST>      </GSTEWAYCONSIGNORADDRESS.LIST>
      <GSTEWAYCONSIGNEEADDRESS.LIST>      </GSTEWAYCONSIGNEEADDRESS.LIST>
      <TEMPGSTRATEDETAILS.LIST>      </TEMPGSTRATEDETAILS.LIST>
      <TEMPGSTADVADJUSTED.LIST>      </TEMPGSTADVADJUSTED.LIST>
      <GSTBUYERADDRESS.LIST>      </GSTBUYERADDRESS.LIST>
      <GSTCONSIGNEEADDRESS.LIST>      </GSTCONSIGNEEADDRESS.LIST>${udfXmls.length > 0 ? '\n' + udfXmls.join('\n') : ''}
     </VOUCHER>
    </TALLYMESSAGE>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <COMPANY>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>2786887a-d92f-46c7-bf13-d8a373da8523</NAME>
       <REMOTECMPNAME>OPG Master</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
     </COMPANY>
    </TALLYMESSAGE>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <COMPANY>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>2786887a-d92f-46c7-bf13-d8a373da8523</NAME>
       <REMOTECMPNAME>OPG Master</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
     </COMPANY>
    </TALLYMESSAGE>
   </REQUESTDATA>
  </IMPORTDATA>
 </BODY>
</ENVELOPE>`;
}

module.exports = {
    getRowValue,
    cleanStateName,
    formatDate,
    padVendor,
    getStockItemName,
    generateTallyXML
};
