const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

let conditionTypeMap = null;
const COMPANY_NAME = 'Opg Legacy Data (22-23)';

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
    // Fallback: match by leading code prefix (e.g. 'JEXS', 'NAVS') for condition type columns
    const codeMatch = colName.match(/^([A-Z0-9]+)/i);
    if (codeMatch) {
        const code = codeMatch[1].toLowerCase();
        const reservedWords = ['po', 'document', 'purchasing', 'material', 'invoicing', 'vendor', 'reference', 'item', 'order', 'quantity', 'amount', 'unit', 'plant', 'stock', 'total'];
        if (code.length >= 3 && !reservedWords.includes(code)) {
            for (const key of Object.keys(row)) {
                const keyCodeMatch = key.match(/^([A-Z0-9]+)/i);
                if (keyCodeMatch && keyCodeMatch[1].toLowerCase() === code) {
                    return row[key];
                }
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
 * Format quantity preserving exact unit accuracy (up to 3 decimal places for fractional units like Coal)
 * without altering integer quantities.
 */
function formatQuantity(val) {
    if (val === undefined || val === null || val === '') return '0';
    const num = parseFloat(String(val).trim());
    if (isNaN(num)) return '0';
    if (Number.isInteger(num)) {
        return String(num);
    }
    return num.toFixed(3);
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
        if (code.length >= 3 && code !== 'po') {
            for (const key of Object.keys(row)) {
                const keyCodeMatch = key.match(/^([A-Z0-9]+)/i);
                if (keyCodeMatch && keyCodeMatch[1].toLowerCase() === code) {
                    return key;
                }
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
    console.log("date before format:", dateValue)
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
    if (vendorId === undefined || vendorId === null || String(vendorId).trim() === '' || String(vendorId).trim() === '0') return '';
    const cleanId = String(vendorId).split('.')[0].trim(); // Remove decimals if any
    return cleanId.padStart(10, '0');
}

/**
 * Construct STOCKITEMNAME with specific suffix matching the template for known items, or fallback
 */
function getStockItemName(material, shortText) {
    let matStr = material ? String(material).split('.')[0].trim() : '';
    if (matStr === '0' || /^0+$/.test(matStr)) {
        matStr = '';
    }
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
    const docType = escapeXML(String(getRowValue(firstRow, 'Doc Type') || 'ZSPR').trim());
    const docDateFormatted = formatDate(getRowValue(firstRow, 'Doc Date'));

    const vendorCode = padVendor(getRowValue(firstRow, 'Vendor'));
    const rawVendorName = String(getRowValue(firstRow, 'Vendor Name') || (vendorCode ? vendorMap[vendorCode] || '' : '')).trim();
    const vendorName = escapeXML(rawVendorName || vendorCode);
    const partyName = escapeXML(vendorCode ? (rawVendorName ? `${vendorCode}-${rawVendorName}` : vendorCode) : rawVendorName);

    const street = String(getRowValue(firstRow, 'Street') || '').trim();
    const city = String(getRowValue(firstRow, 'City') || '').trim();
    const address = escapeXML(street && city ? `${street},${city}${city}` : (street || city));

    const postCode = escapeXML(getRowValue(firstRow, 'Post Code') ? String(getRowValue(firstRow, 'Post Code')).split('.')[0].trim() : '');
    const gstNo = escapeXML(String(getRowValue(firstRow, 'GST NO') || '').trim());
    const isGst33OrBlank = !gstNo || gstNo.startsWith('33');
    const regionName = escapeXML(cleanStateName(getRowValue(firstRow, 'Region Name')));

    const cmpState = 'Tamil Nadu'; // Default Company state is Tamil Nadu
    const isLocal = regionName.toLowerCase().replace(/\s/g, '') === cmpState.toLowerCase().replace(/\s/g, '');

    // Group totals
    let totalNetValue = 0;

    const activeItems = poGroup.items.filter(item => {
        const delInd = getRowValue(item, 'Deletion Indicator');
        return !(delInd && String(delInd).trim().toUpperCase() === 'L');
    });

    if (activeItems.length === 0) {
        return '';
    }

    const itemsXML = activeItems.map(item => {
        const material = getRowValue(item, 'Material');
        const shortText = getRowValue(item, 'Material Description') || getRowValue(item, 'Purchase Order - Short Text') || getRowValue(item, 'Short Text') || getRowValue(item, 'Text');
        const stockItemName = escapeXML(getStockItemName(material, shortText));

        const qty = parseFloat(getRowValue(item, 'Order Quantity')) || 0;
        const unit = escapeXML(String(getRowValue(item, 'Order Unit') || 'Nos').trim());
        const price = parseFloat(getRowValue(item, 'Net Order Price')) || 0;
        const amount = parseFloat(getRowValue(item, 'Net Order Value')) || (qty * price);

        totalNetValue += amount;

        const rateFormatted = `${price.toFixed(2)}/${unit}`;
        const amountFormatted = `-${amount.toFixed(2)}`;
        const qtyFormatted = ` ${formatQuantity(qty)} ${unit}`;

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
        { key: 'TransSample2', inr: 'TRANSSAMPLE2INR', str: 'TRANSSAMPLE2', sub: 'TRANSSAMPLE2SUB', idxInr: '2146', idxStr: '2144', idxSub: '2145' },
        { key: 'TransSTock', inr: 'TRANSSTOCKINR', str: 'TRANSSTOCK', sub: 'TRANSSTOCKSUB', idxInr: '2103', idxStr: '2101', idxSub: '2102' },
        { key: 'TransDraft', inr: 'TRANSDRAFTINR', str: 'TRANSDRAFT', sub: 'TRANSDRAFTSUB', idxInr: '2106', idxStr: '2104', idxSub: '2105' },
        { key: 'TransLia', inr: 'TRANSLIAINR', str: 'TRANSLIA', sub: 'TRANSLIASUB', idxInr: '2109', idxStr: '2107', idxSub: '2108' },
        { key: 'TransWhar', inr: 'TRANSWHARINR', str: 'TRANSWHAR', sub: 'TRANSWHARSUB', idxInr: '2112', idxStr: '2110', idxSub: '2111' },
        { key: 'TransEquip', inr: 'TRANSEQUIPINR', str: 'TRANSEQUIP', sub: 'TRANSEQUIPSUB', idxInr: '2115', idxStr: '2113', idxSub: '2114' },
        { key: 'TransDisc', inr: 'TRANSDISCINR', str: 'TRANSDISC', sub: 'TRANSDISCSUB', idxInr: '2118', idxStr: '2116', idxSub: '2117' },
        { key: 'TransStev', inr: 'TRANSSTEVINR', str: 'TRANSSTEV', sub: 'TRANSSTEVSUB', idxInr: '2121', idxStr: '2119', idxSub: '2120' },
        { key: 'TransWharEn', inr: 'TRANSWHARENINR', str: 'TRANSWHAREN', sub: 'TRANSWHARENSUB', idxInr: '2124', idxStr: '2122', idxSub: '2123' },
        { key: 'TransTrans', inr: 'TRANSTRANSINR', str: 'TRANSTRANS', sub: 'TRANSTRANSSUB', idxInr: '2127', idxStr: '21251', idxSub: '2126' },
        { key: 'TransHandling', inr: 'TRANSHANDLINGINR', str: 'TRANSHANDLING', sub: 'TRANSHANDLINGSUB', idxInr: '2130', idxStr: '2128', idxSub: '2129' }
    ];

    let udfCount = 0;
    const udfXmls = [];
    const activeLedgers = [];
    let totalTaxesAndCharges = 0;

    ledgerColumns.forEach(col => {
        let colSum = 0;
        activeItems.forEach(item => {
            const rawVal = getRowValue(item, col);
            if (rawVal !== undefined && rawVal !== null) {
                const cleanVal = String(rawVal).replace(/,/g, '').trim();
                colSum += parseFloat(cleanVal) || 0;
            }
        });

        if (Math.abs(colSum) > 0.001) {
            const cleanColName = col.toLowerCase().replace(/\s/g, '');

            // JEXS / NAVS split logic
            if (cleanColName === 'jexs' || cleanColName === 'navs') {
                if (isGst33OrBlank) {
                    const halfSum = colSum / 2;
                    activeLedgers.push({
                        name: 'CGST',
                        sum: halfSum
                    });
                    activeLedgers.push({
                        name: 'SGST',
                        sum: halfSum
                    });
                } else {
                    activeLedgers.push({
                        name: 'IGST',
                        sum: colSum
                    });
                }
                totalTaxesAndCharges += colSum;
                return;
            }

            const exactKey = getExactKey(firstRow, col);
            const condCode = getConditionCode(col);
            const mappedName = condMap[condCode];

            const finalLedgerName = mappedName ? mappedName : exactKey.toString().replace(/\s+/g, ' ').trim();
            if (finalLedgerName.toLowerCase().replace(/\s/g, '') === 'grossprice') {
                return;
            }
            const normalizedName = escapeXML(finalLedgerName);

            if (docType === 'ZCOL' && cleanColName !== 'jexs' && cleanColName !== 'navs' && cleanColName !== 'zcec' && cleanColName !== 'zceq') {
                if (udfCount < udfSlots.length) {
                    const slot = udfSlots[udfCount];
                    udfCount++;

                    let vendorCode = '';
                    for (const item of activeItems) {
                        const vCode = getVendorCodeForRow(item, col);
                        if (vCode !== undefined && vCode !== null && String(vCode).trim() !== '' && parseFloat(vCode) !== 0) {
                            vendorCode = String(vCode).split('.')[0].trim();
                            break;
                        }
                    }
                    if (!vendorCode && activeItems.length > 0) {
                        const vCode = getVendorCodeForRow(activeItems[0], col);
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
    const remoteId = `2786887a-d92f-46c7-bf13-d8a373da8523x-${poNumber.padStart(8, '0')}`;
    const vchKey = `2786887a-d92f-46c7-bf13-d8a373da8523x-0000b420:${poNumber.padStart(8, '0')}`;

    return `<ENVELOPE>
 <HEADER>
  <TALLYREQUEST>Import Data</TALLYREQUEST>
 </HEADER>
 <BODY>
  <IMPORTDATA>
   <REQUESTDESC>
    <REPORTNAME>Vouchers</REPORTNAME>
    <STATICVARIABLES>
     <SVCURRENTCOMPANY>${escapeXML(COMPANY_NAME)}</SVCURRENTCOMPANY>
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
      <BASICBUYERNAME>${escapeXML(COMPANY_NAME)}</BASICBUYERNAME>
      <CMPGSTREGISTRATIONTYPE>Regular</CMPGSTREGISTRATIONTYPE>
      <REFERENCE>${poNumber}</REFERENCE>
      <PARTYMAILINGNAME>${vendorName}</PARTYMAILINGNAME>
      <PARTYPINCODE>${postCode}</PARTYPINCODE>
      <CONSIGNEEMAILINGNAME>${escapeXML(COMPANY_NAME)}</CONSIGNEEMAILINGNAME>
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
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
     </COMPANY>
    </TALLYMESSAGE>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <COMPANY>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>2786887a-d92f-46c7-bf13-d8a373da8523</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
     </COMPANY>
    </TALLYMESSAGE>
   </REQUESTDATA>
  </IMPORTDATA>
 </BODY>
</ENVELOPE>`;
}

let vendorMasterMap = null;
let poMasterMap = null;

function loadPOMaster() {
    if (poMasterMap) return poMasterMap;
    poMasterMap = {};
    try {
        const parentDir = path.resolve(__dirname, '..');
        const files = fs.readdirSync(parentDir);
        for (const file of files) {
            if (file.toLowerCase().endsWith('.xlsx') && !file.toLowerCase().includes('grn') && !file.toLowerCase().includes('invoice')) {
                const filePath = path.join(parentDir, file);
                try {
                    const workbook = xlsx.readFile(filePath);
                    for (const sheetName of workbook.SheetNames) {
                        if (sheetName.toLowerCase().includes('detail') || sheetName.toLowerCase().includes('sheet')) {
                            const sheet = workbook.Sheets[sheetName];
                            const rawGrid = xlsx.utils.sheet_to_json(sheet, { header: 1 });
                            if (rawGrid.length === 0) continue;

                            let headerRowIndex = -1;
                            for (let i = 0; i < Math.min(rawGrid.length, 20); i++) {
                                const row = rawGrid[i];
                                if (row && (row.includes('Purchasing Document') || row.includes('PO Number'))) {
                                    headerRowIndex = i;
                                    break;
                                }
                            }
                            if (headerRowIndex === -1) continue;

                            const headers = rawGrid[headerRowIndex].map(h => String(h || '').trim());
                            const poColIdx = headers.findIndex(h => h.includes('Purchasing Document') || h === 'PO Number');
                            const typeColIdx = headers.findIndex(h => h.includes('Doc Type') || h.includes('PO - Doc Type') || h.includes('Document Type'));

                            if (poColIdx === -1) continue;

                            for (let i = headerRowIndex + 1; i < rawGrid.length; i++) {
                                const row = rawGrid[i];
                                if (row && row[poColIdx]) {
                                    const poNum = String(row[poColIdx]).split('.')[0].trim();
                                    const docType = typeColIdx !== -1 && row[typeColIdx] ? String(row[typeColIdx]).trim() : 'ZSPR';
                                    poMasterMap[poNum] = docType;
                                }
                            }
                        }
                    }
                } catch (err) {
                    console.error('Error reading file for PO master:', file, err);
                }
            }
        }
        console.log(`Loaded ${Object.keys(poMasterMap).length} PO doc types into master map.`);
    } catch (err) {
        console.error('Error building PO master map:', err);
    }
    return poMasterMap;
}

function generatePurchaseTallyXML(purchaseGroup, vendorMap = {}) {
    const vendors = loadVendorMaster();
    const poMaster = loadPOMaster();
    const firstRow = purchaseGroup.items[0];

    const rawVoucherNo = getRowValue(firstRow, 'Invoice No') ||
                         getRowValue(firstRow, 'Invoice Number') ||
                         getRowValue(firstRow, 'Document Number') ||
                         '';
    const voucherNumber = escapeXML(String(rawVoucherNo).split('.')[0].trim());
    const reference = escapeXML(String(getRowValue(firstRow, 'Reference')).trim());
    const docDateFormatted = formatDate(getRowValue(firstRow, 'Document Date') || getRowValue(firstRow, 'Posting Date'));

    const firstRowDocType = String(
        getRowValue(firstRow, 'Purchasing Doc Type') ||
        getRowValue(firstRow, 'Purchasing Doc. Type') ||
        getRowValue(firstRow, 'PO - Doc Type') ||
        getRowValue(firstRow, 'Doc Type') ||
        ''
    ).trim();

    const poNumberHeader = escapeXML(String(getRowValue(firstRow, 'Purchasing Document') || getRowValue(firstRow, 'Purchase Order') || '').split('.')[0].trim());
    const headerDocType = firstRowDocType || (poNumberHeader && poMaster[poNumberHeader] ? poMaster[poNumberHeader] : 'ZSPR');
    const voucherTypeName = `Purchase ${headerDocType}`;

    const rawVendorCode = getRowValue(firstRow, 'Invoicing Party') || getRowValue(firstRow, 'Vendor');
    const vendorCode = padVendor(rawVendorCode);
    const rawVendorName = String(getRowValue(firstRow, 'Vendor Name') || (vendorCode ? (vendorMap[vendorCode] || vendors[vendorCode]?.vendorName || '') : '')).trim();
    const vendorName = escapeXML(rawVendorName || vendorCode);
    const partyName = escapeXML(vendorCode ? (rawVendorName ? `${vendorCode}-${rawVendorName}` : vendorCode) : rawVendorName);

    const street = String(vendors[vendorCode]?.street || getRowValue(firstRow, 'Street') || '').trim();
    const city = String(vendors[vendorCode]?.city || getRowValue(firstRow, 'City') || '').trim();
    const address = escapeXML(street && city ? `${street},${city}` : (street || city || ''));

    const postCode = escapeXML(vendors[vendorCode]?.postCode || getRowValue(firstRow, 'Post Code') || '');
    const gstNo = escapeXML(vendors[vendorCode]?.gstNo || getRowValue(firstRow, 'GST NO') || '');
    const isGst33OrBlank = !gstNo || gstNo.startsWith('33');
    const regionName = escapeXML(cleanStateName(vendors[vendorCode]?.regionName || getRowValue(firstRow, 'Region Name') || 'Tamil Nadu'));

    const cmpState = 'Tamil Nadu';
    const isLocal = regionName.toLowerCase().replace(/\s/g, '') === cmpState.toLowerCase().replace(/\s/g, '');

    const activeItems = purchaseGroup.items.filter(item => {
        const delInd = getRowValue(item, 'Deletion Indicator');
        return !(delInd && String(delInd).trim().toUpperCase() === 'L');
    });

    if (activeItems.length === 0) {
        return '';
    }

    const stockItems = activeItems.filter(item => {
        const condType = getRowValue(item, 'Condition Type');
        return !condType || String(condType).trim() === '';
    });

    const additionalLedgerItems = activeItems.filter(item => {
        const condType = getRowValue(item, 'Condition Type');
        return condType && String(condType).trim() !== '';
    });

    let totalNetValue = 0;

    const itemsXML = stockItems.map(item => {
        const material = getRowValue(item, 'Material');
        const shortText = getRowValue(item, 'Material Description') || getRowValue(item, 'Purchase Order - Short Text') || getRowValue(item, 'Text') || getRowValue(item, 'Short Text');
        const stockItemName = escapeXML(getStockItemName(material, shortText));

        const qty = parseFloat(getRowValue(item, 'Quantity') || getRowValue(item, 'Qty in OPUn')) || 0;
        const unit = escapeXML(String(getRowValue(item, 'Order Unit') || getRowValue(item, 'Order Price Unit') || getRowValue(item, 'Base Unit of Measure') || 'Nos').trim());
        const amount = parseFloat(getRowValue(item, 'Amount') || getRowValue(item, 'Total Value')) || 0;
        const price = qty > 0 ? (amount / qty) : 0;

        totalNetValue += amount;

        const rateFormatted = `${price.toFixed(2)}/${unit}`;
        const amountFormatted = `-${amount.toFixed(2)}`;
        const qtyFormatted = ` ${formatQuantity(qty)} ${unit}`;

        const poNumber = escapeXML(String(getRowValue(item, 'Purchasing Document') || getRowValue(item, 'Purchase Order') || '').split('.')[0].trim());
        const itemDocType = String(
            getRowValue(item, 'Purchasing Doc Type') ||
            getRowValue(item, 'Purchasing Doc. Type') ||
            getRowValue(item, 'PO - Doc Type') ||
            getRowValue(item, 'Doc Type') ||
            ''
        ).trim() || (poNumber && poMaster[poNumber] ? poMaster[poNumber] : headerDocType);
        const ledgerName = `Purchase ${itemDocType}`;

        const godownName = escapeXML(String(getRowValue(item, 'Plant') || '1000').split('.')[0].trim());

        const rawRefDoc = getRowValue(item, 'Reference Document') || getRowValue(item, 'Reference Document Item') || getRowValue(item, 'Ref Document') || getRowValue(item, 'GRN Number') || getRowValue(item, 'GRN') || '';
        let trackingNumber = '&#4; Not Applicable';
        if (rawRefDoc !== undefined && rawRefDoc !== null && String(rawRefDoc).trim() !== '') {
            trackingNumber = escapeXML(String(rawRefDoc).split('.')[0].trim());
        }

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
         <GODOWNNAME>${godownName}</GODOWNNAME>
         <BATCHNAME>Primary Batch</BATCHNAME>
         <DESTINATIONGODOWNNAME>${godownName}</DESTINATIONGODOWNNAME>
         <INDENTNO>&#4; Not Applicable</INDENTNO>
         <ORDERNO>${poNumber}</ORDERNO>
         <TRACKINGNUMBER>${trackingNumber}</TRACKINGNUMBER>
         <DYNAMICCSTISCLEARED>No</DYNAMICCSTISCLEARED>
         <AMOUNT>${amountFormatted}</AMOUNT>
         <ACTUALQTY>${qtyFormatted}</ACTUALQTY>
         <BILLEDQTY>${qtyFormatted}</BILLEDQTY>
         <ORDERDUEDATE JD="45349" P="1-Apr-26">1-Apr-26</ORDERDUEDATE>
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
        <RATEDETAILS.LIST>       </RATEDETAILS.LIST>
        <SUPPLEMENTARYDUTYHEADDETAILS.LIST>       </SUPPLEMENTARYDUTYHEADDETAILS.LIST>
        <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>
        <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>
        <EXCISEALLOCATIONS.LIST>       </EXCISEALLOCATIONS.LIST>
        <EXPENSEALLOCATIONS.LIST>       </EXPENSEALLOCATIONS.LIST>
       </ALLINVENTORYENTRIES.LIST>`;
    }).join('\n');

    // Group additionalLedgerItems by condition type
    const condGroups = {};
    let totalAdditionalValue = 0;
    additionalLedgerItems.forEach(item => {
        const condType = String(getRowValue(item, 'Condition Type')).trim();
        const amt = parseFloat(getRowValue(item, 'Amount')) || 0;
        if (condType) {
            condGroups[condType] = (condGroups[condType] || 0) + amt;
            totalAdditionalValue += amt;
        }
    });

    let additionalLedgersXML = '';
    Object.entries(condGroups).forEach(([condType, amount]) => {
        additionalLedgersXML += `      <LEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <LEDGERNAME>${escapeXML(condType)}</LEDGERNAME>
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
       <AMOUNT>-${amount.toFixed(2)}</AMOUNT>
       <VATEXPAMOUNT>-${amount.toFixed(2)}</VATEXPAMOUNT>
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
      </LEDGERENTRIES.LIST>\n`;
    });

    let taxAmt = 0;
    for (const item of activeItems) {
        const val = parseFloat(getRowValue(item, 'Value-Added Tax Amt'));
        if (val && !isNaN(val)) {
            taxAmt = val;
            break;
        }
    }
    const totalPartyValueFormatted = (totalNetValue + totalAdditionalValue + taxAmt).toFixed(2);

    let taxesLedgerXML = '';
    if (taxAmt > 0) {
        taxesLedgerXML = `      <LEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <LEDGERNAME>Taxes</LEDGERNAME>
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
       <AMOUNT>-${taxAmt.toFixed(2)}</AMOUNT>
       <VATEXPAMOUNT>-${taxAmt.toFixed(2)}</VATEXPAMOUNT>
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
    }

    const guid = `81f73e2b-a3c5-4ff2-a56f-${voucherNumber}`;
    const vchKey = `${guid}:00000000`;

    return `<ENVELOPE>
 <HEADER>
  <TALLYREQUEST>Import Data</TALLYREQUEST>
 </HEADER>
 <BODY>
  <IMPORTDATA>
   <REQUESTDESC>
    <REPORTNAME>Vouchers</REPORTNAME>
    <STATICVARIABLES>
     <SVCURRENTCOMPANY>${escapeXML(COMPANY_NAME)}</SVCURRENTCOMPANY>
    </STATICVARIABLES>
   </REQUESTDESC>
   <REQUESTDATA>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <VOUCHER REMOTEID="${guid}" VCHKEY="${vchKey}" VCHTYPE="${voucherTypeName}" ACTION="Create" OBJVIEW="Invoice Voucher View">
      <ADDRESS.LIST TYPE="String">
       <ADDRESS>${address}</ADDRESS>
      </ADDRESS.LIST>
      <OLDAUDITENTRYIDS.LIST TYPE="Number">
       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
      </OLDAUDITENTRYIDS.LIST>
      <DATE>${docDateFormatted}</DATE>
      <REFERENCEDATE>${docDateFormatted}</REFERENCEDATE>
      <VCHSTATUSDATE>${docDateFormatted}</VCHSTATUSDATE>
      <GUID>${guid}</GUID>
      <GSTREGISTRATIONTYPE>&#4; Unknown</GSTREGISTRATIONTYPE>
      <VOUCHERNUMBER>${voucherNumber}</VOUCHERNUMBER>
      <PARTYLEDGERNAME>${partyName}</PARTYLEDGERNAME>
      <REFERENCE>${reference}</REFERENCE>
      <BASICBUYERNAME>${escapeXML(COMPANY_NAME)}</BASICBUYERNAME>
      <SVCURRENTCOMPANY>${escapeXML(COMPANY_NAME)}</SVCURRENTCOMPANY>
      <VOUCHERTYPENAME>${voucherTypeName}</VOUCHERTYPENAME>
      <ISINVOICE>Yes</ISINVOICE>
      <VCHENTRYMODE>Item Invoice</VCHENTRYMODE>
      <EFFECTIVEDATE>${docDateFormatted}</EFFECTIVEDATE>
      <PARTYNAME>${partyName}</PARTYNAME>
      <USETRACKINGNUMBER>No</USETRACKINGNUMBER>
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
      <CHANGEVCHMODE>No</CHANGEVCHMODE>
      <RESETIRNQRCODE>No</RESETIRNQRCODE>
      <VOUCHERNUMBERSERIES>Default</VOUCHERNUMBERSERIES>
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
       <AMOUNT>${totalPartyValueFormatted}</AMOUNT>
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
${additionalLedgersXML}
${taxesLedgerXML}
      <GST.LIST>      </GST.LIST>
      <STKJRNLADDLCOSTDETAILS.LIST>      </STKJRNLADDLCOSTDETAILS.LIST>
      <PAYROLLMODEOFPAYMENT.LIST>      </PAYROLLMODEOFPAYMENT.LIST>
      <ATTDRECORDS.LIST>      </ATTDRECORDS.LIST>
      <GSTEWAYCONSIGNORADDRESS.LIST>      </GSTEWAYCONSIGNORADDRESS.LIST>
      <GSTEWAYCONSIGNEEADDRESS.LIST>      </GSTEWAYCONSIGNEEADDRESS.LIST>
      <TEMPGSTRATEDETAILS.LIST>      </TEMPGSTRATEDETAILS.LIST>
      <TEMPGSTADVADJUSTED.LIST>      </TEMPGSTADVADJUSTED.LIST>
      <GSTBUYERADDRESS.LIST>      </GSTBUYERADDRESS.LIST>
      <GSTCONSIGNEEADDRESS.LIST>      </GSTCONSIGNEEADDRESS.LIST>
     </VOUCHER>
    </TALLYMESSAGE>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <COMPANY>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>81f73e2b-a3c5-4ff2-a56f-49d15ff7c0f6</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>2786887a-d92f-46c7-bf13-d8a373da8523</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
     </COMPANY>
    </TALLYMESSAGE>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <COMPANY>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>81f73e2b-a3c5-4ff2-a56f-49d15ff7c0f6</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>2786887a-d92f-46c7-bf13-d8a373da8523</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
     </COMPANY>
    </TALLYMESSAGE>
   </REQUESTDATA>
  </IMPORTDATA>
 </BODY>
</ENVELOPE>`;
}

function loadVendorMaster() {
    if (vendorMasterMap) return vendorMasterMap;
    vendorMasterMap = {};
    try {
        const parentDir = path.resolve(__dirname, '..');
        const files = fs.readdirSync(parentDir);
        for (const file of files) {
            if (file.toLowerCase().endsWith('.xlsx') && !file.toLowerCase().includes('grn')) {
                const filePath = path.join(parentDir, file);
                try {
                    const workbook = xlsx.readFile(filePath);
                    for (const sheetName of workbook.SheetNames) {
                        if (sheetName.toLowerCase().includes('detail')) {
                            const sheet = workbook.Sheets[sheetName];
                            const rawGrid = xlsx.utils.sheet_to_json(sheet, { header: 1 });
                            if (rawGrid.length === 0) continue;

                            // Find header row
                            let headerRowIndex = -1;
                            for (let i = 0; i < Math.min(rawGrid.length, 20); i++) {
                                const row = rawGrid[i];
                                if (row && (row.includes('Vendor') || row.includes('Vendor ') || row.includes('Vendor Name') || row.includes('Purchasing Document'))) {
                                    headerRowIndex = i;
                                    break;
                                }
                            }
                            if (headerRowIndex === -1) continue;

                            const headers = rawGrid[headerRowIndex].map(h => String(h || '').trim());
                            for (let i = headerRowIndex + 1; i < rawGrid.length; i++) {
                                const row = rawGrid[i];
                                if (!row || row.length === 0) continue;

                                const obj = {};
                                headers.forEach((h, idx) => {
                                    if (h) {
                                        const newVal = row[idx];
                                        const hasCurrent = obj[h] !== undefined && obj[h] !== null && String(obj[h]).trim() !== '';
                                        const hasNew = newVal !== undefined && newVal !== null && String(newVal).trim() !== '';
                                        if (hasCurrent && !hasNew) {
                                            // Keep current populated value
                                        } else {
                                            obj[h] = newVal;
                                        }
                                    }
                                });

                                const vendorVal = getRowValue(obj, 'Vendor');
                                if (vendorVal !== undefined && vendorVal !== null) {
                                    const cleanCode = String(vendorVal).split('.')[0].trim().padStart(10, '0');
                                    if (cleanCode && cleanCode !== '0000000000' && !vendorMasterMap[cleanCode]) {
                                        vendorMasterMap[cleanCode] = {
                                            vendorName: String(getRowValue(obj, 'Vendor Name') || '').trim(),
                                            city: String(getRowValue(obj, 'City') || '').trim(),
                                            postCode: String(getRowValue(obj, 'Post Code') || '').split('.')[0].trim(),
                                            street: String(getRowValue(obj, 'Street') || '').trim(),
                                            regionName: String(getRowValue(obj, 'Region Name') || '').trim(),
                                            gstNo: String(getRowValue(obj, 'GST NO') || '').trim()
                                        };
                                    }
                                }
                            }
                        }
                    }
                } catch (e) {
                    console.error(`Error reading vendor details from ${file}:`, e.message);
                }
            }
        }
        console.log(`Loaded ${Object.keys(vendorMasterMap).length} vendors into master map.`);
    } catch (err) {
        console.error('Error loading vendor master:', err);
    }
    return vendorMasterMap;
}

//xml generation for GRN

function generateGRNTallyXML(grnGroup) {
    const firstRow = grnGroup.items[0];
    const eventType = String(getRowValue(firstRow, 'Trans./Event Type') || getRowValue(firstRow, 'Trans./Event TypeA') || '').trim().toUpperCase();
    if (eventType === 'WA') {
        return generateStockJournalTallyXML(grnGroup);
    }

    const vendors = loadVendorMaster();

    // Material Document is the GRN number
    const grnNumber = escapeXML(String(getRowValue(firstRow, 'Material Document')).split('.')[0].trim());
    const poNumber = escapeXML(String(getRowValue(firstRow, 'Purchase Order')).split('.')[0].trim());
    // Date formatting
    const docDateFormatted = formatDate(getRowValue(firstRow, 'Document Date') || getRowValue(firstRow, 'Posting Date'));

    // Vendor Info
    const rawVendorCode = getRowValue(firstRow, 'Vendor');
    const vendorCode = padVendor(rawVendorCode);
    const rawVendorName = String(getRowValue(firstRow, 'Vendor Description') || '').trim();

    // Look up in our master vendor database
    let masterVendor = vendors[vendorCode] || {};
    const finalVendorName = masterVendor.vendorName || rawVendorName;
    const vendorName = escapeXML(finalVendorName || vendorCode);
    const partyName = escapeXML(vendorCode ? (finalVendorName ? `${vendorCode}-${finalVendorName}` : vendorCode) : finalVendorName);

    const street = String(masterVendor.street || '').trim();
    const city = String(masterVendor.city || '').trim();
    const address = escapeXML(street && city ? `${street},${city}` : (street || city || ''));

    const postCode = escapeXML(masterVendor.postCode || '');
    const gstNo = escapeXML(masterVendor.gstNo || '');
    const isGst33OrBlank = !gstNo || gstNo.startsWith('33');
    const regionName = escapeXML(cleanStateName(masterVendor.regionName || 'Tamil Nadu'));

    const cmpState = 'Tamil Nadu';
    const isLocal = regionName.toLowerCase().replace(/\s/g, '') === cmpState.toLowerCase().replace(/\s/g, '');

    // Determine Voucher Type
    // Trans./Event Type column: if "WE", type is "Receipt Note WS", else "Receipt Note"
    const voucherType = `Receipt Note ${eventType}`;

    // Group totals
    let totalNetValue = 0;

    const activeItems = grnGroup.items.filter(item => {
        const delInd = getRowValue(item, 'Deletion Indicator');
        return !(delInd && String(delInd).trim().toUpperCase() === 'L');
    });

    if (activeItems.length === 0) {
        return '';
    }

    const itemsXML = activeItems.map(item => {
        const material = getRowValue(item, 'Material');
        const shortText = getRowValue(item, 'Material Description') || getRowValue(item, 'Purchase Order - Short Text') || getRowValue(item, 'Short Text') || getRowValue(item, 'Text');

        // Stock Item Name is the material column value (acting as an alias in Tally)
        const stockItemName = escapeXML(getStockItemName(material, shortText));

        const qty = parseFloat(getRowValue(item, 'Qty in Un. of Entry') || getRowValue(item, 'Order Quantity')) || 0;
        const unit = escapeXML(String(getRowValue(item, 'Unit of Entry') || getRowValue(item, 'Order Unit') || 'Nos').trim());

        // Amount in LC in Excel is tax-inclusive. We divide by 1.18 to get the net amount.
        const amountLC = parseFloat(getRowValue(item, 'Amount in LC')) || 0;
        const amount = amountLC / 1.18;
        const price = qty > 0 ? (amount / qty) : 0;

        totalNetValue += amount;

        const rateFormatted = `${price.toFixed(2)}/${unit}`;
        const amountFormatted = `-${amount.toFixed(2)}`;
        const qtyFormatted = ` ${formatQuantity(qty)} ${unit}`;

        // For GRN, ledger name is Purchase ZSPR (or Purchase Doc Type if we can determine Doc Type)
        // Since we don't have Doc Type in GRN Excel, we default it to ZSPR
        const docType = 'ZSPR';
        const ledgerName = `Purchase ${docType}`;

        const poNumber = escapeXML(String(getRowValue(item, 'Purchase Order') || '').split('.')[0].trim());

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
         <TRACKINGNUMBER>${grnNumber}</TRACKINGNUMBER>
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

    // Tax calculation
    const activeLedgers = [];
    let totalTaxesAndCharges = 0;

    // CGST/SGST/IGST tax allocation
    // 9% CGST + 9% SGST if local, or 18% IGST if interstate
    const totalTax = totalNetValue * 0.18;
    if (isGst33OrBlank) {
        const halfTax = totalTax / 2;
        activeLedgers.push({
            name: 'CGST Tax',
            sum: halfTax
        });
        activeLedgers.push({
            name: 'SGST Tax',
            sum: halfTax
        });
    } else {
        activeLedgers.push({
            name: 'IGST Tax',
            sum: totalTax
        });
    }
    totalTaxesAndCharges += totalTax;

    const totalVoucherAmount = totalNetValue + totalTaxesAndCharges;

    const taxLedgerXML = activeLedgers.map(led => {
        const isDeemedPositive = led.sum > 0 ? 'Yes' : 'No';
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

    const remoteId = `81f73e2b-a3c5-4ff2-a56f-49d15ff7c0f6x-${grnNumber.padStart(8, '0')}`;
    const vchKey = `81f73e2b-a3c5-4ff2-a56f-49d15ff7c0f6x-0000b097:${grnNumber.padStart(8, '0')}`;

    return `<ENVELOPE>
 <HEADER>
  <TALLYREQUEST>Import Data</TALLYREQUEST>
 </HEADER>
 <BODY>
  <IMPORTDATA>
   <REQUESTDESC>
    <REPORTNAME>Vouchers</REPORTNAME>
    <STATICVARIABLES>
     <SVCURRENTCOMPANY>${escapeXML(COMPANY_NAME)}</SVCURRENTCOMPANY>
    </STATICVARIABLES>
   </REQUESTDESC>
   <REQUESTDATA>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <VOUCHER REMOTEID="${remoteId}" VCHKEY="${vchKey}" VCHTYPE="${voucherType}" ACTION="Create" OBJVIEW="Invoice Voucher View">
      <ADDRESS.LIST TYPE="String">
       <ADDRESS>${address}</ADDRESS>
      </ADDRESS.LIST>
      <OLDAUDITENTRYIDS.LIST TYPE="Number">
       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
      </OLDAUDITENTRYIDS.LIST>
      <DATE>${docDateFormatted}</DATE>
      <REFERENCEDATE>${docDateFormatted}</REFERENCEDATE>
      <VCHSTATUSDATE>${docDateFormatted}</VCHSTATUSDATE>
      <GUID>${remoteId}</GUID>
      <GSTREGISTRATIONTYPE>&#4; Unknown</GSTREGISTRATIONTYPE>
      <VATDEALERTYPE>&#4; Unknown</VATDEALERTYPE>
      <STATENAME>${regionName}</STATENAME>
      <OBJECTUPDATEACTION/>
      <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
      <PARTYGSTIN>${gstNo}</PARTYGSTIN>
      <PLACEOFSUPPLY>${regionName}</PLACEOFSUPPLY>
      <PARTYNAME>${partyName}</PARTYNAME>
      <GSTREGISTRATION TAXTYPE="GST" TAXREGISTRATION="">${regionName} Registration</GSTREGISTRATION>
      <VOUCHERTYPENAME>${voucherType}</VOUCHERTYPENAME>
      <PARTYLEDGERNAME>${partyName}</PARTYLEDGERNAME>
      <VOUCHERNUMBER>${grnNumber}</VOUCHERNUMBER>
      <BASICBUYERNAME>${escapeXML(COMPANY_NAME)}</BASICBUYERNAME>
      <REFERENCE>${poNumber}</REFERENCE>
      <PARTYMAILINGNAME>${vendorName}</PARTYMAILINGNAME>
      <CONSIGNEEMAILINGNAME>${escapeXML(COMPANY_NAME)}</CONSIGNEEMAILINGNAME>
      <CONSIGNEESTATENAME>${cmpState}</CONSIGNEESTATENAME>
      <CONSIGNEECOUNTRYNAME>India</CONSIGNEECOUNTRYNAME>
      <BASICBASEPARTYNAME>${partyName}</BASICBASEPARTYNAME>
      <NUMBERINGSTYLE>Auto Retain</NUMBERINGSTYLE>
      <CSTFORMISSUETYPE>&#4; Not Applicable</CSTFORMISSUETYPE>
      <CSTFORMRECVTYPE>&#4; Not Applicable</CSTFORMRECVTYPE>
      <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>
      <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
      <VCHSTATUSTAXADJUSTMENT>Default</VCHSTATUSTAXADJUSTMENT>
      <VCHSTATUSVOUCHERTYPE>${voucherType}</VCHSTATUSVOUCHERTYPE>
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
      <CHANGEVCHMODE>No</CHANGEVCHMODE>
      <RESETIRNQRCODE>No</RESETIRNQRCODE>
      <ALTERID> 16436</ALTERID>
      <MASTERID> 2100</MASTERID>
      <VOUCHERKEY>194162586550288</VOUCHERKEY>
      <VOUCHERRETAINKEY>1</VOUCHERRETAINKEY>
      <VOUCHERNUMBERSERIES>Default</VOUCHERNUMBERSERIES>
      <UPDATEDDATETIME>20260704113404000</UPDATEDDATETIME>
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
      <INVOICEORDERLIST.LIST>
       <BASICORDERDATE>${docDateFormatted}</BASICORDERDATE>
       <ORDERTYPE>Purchase Order</ORDERTYPE>
       <BASICPURCHASEORDERNO>${poNumber}</BASICPURCHASEORDERNO>
      </INVOICEORDERLIST.LIST>
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
      <GSTCONSIGNEEADDRESS.LIST>      </GSTCONSIGNEEADDRESS.LIST>
     </VOUCHER>
    </TALLYMESSAGE>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <COMPANY>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>81f73e2b-a3c5-4ff2-a56f-49d15ff7c0f6</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>2786887a-d92f-46c7-bf13-d8a373da8523</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
     </COMPANY>
    </TALLYMESSAGE>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <COMPANY>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>81f73e2b-a3c5-4ff2-a56f-49d15ff7c0f6</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>2786887a-d92f-46c7-bf13-d8a373da8523</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
     </COMPANY>
    </TALLYMESSAGE>
   </REQUESTDATA>
  </IMPORTDATA>
 </BODY>
</ENVELOPE>`;
}

function generateStockJournalTallyXML(grnGroup) {
    const firstRow = grnGroup.items[0];

    // Material Document is the voucher number
    const grnNumber = escapeXML(String(getRowValue(firstRow, 'Material Document') || '').split('.')[0].trim());
    const docDateFormatted = formatDate(getRowValue(firstRow, 'Document Date') || getRowValue(firstRow, 'Posting Date'));

    const activeItems = grnGroup.items.filter(item => {
        const delInd = getRowValue(item, 'Deletion Indicator');
        return !(delInd && String(delInd).trim().toUpperCase() === 'L');
    });

    if (activeItems.length === 0) {
        return '';
    }

    // Filter to only positive items to avoid duplicate processing of the negative pair row
    let transferItems = activeItems.filter(item => {
        const qty = parseFloat(getRowValue(item, 'Qty in Un. of Entry')) || 0;
        return qty > 0;
    });

    // Fallback if no positive items are found (e.g. data anomaly)
    if (transferItems.length === 0) {
        transferItems = activeItems;
    }

    const firstItem = transferItems[0];
    const receivingPlant = escapeXML(String(getRowValue(firstItem, 'Receiving Plant') || getRowValue(firstItem, 'Plant') || '').split('.')[0].trim());
    const cmpState = 'Tamil Nadu';

    // Generate IN inventory entries (deemed positive: Yes)
    const inventoryInXML = transferItems.map(item => {
        const material = getRowValue(item, 'Material');
        const shortText = getRowValue(item, 'Material Description') || getRowValue(item, 'Purchase Order - Short Text') || getRowValue(item, 'Short Text') || getRowValue(item, 'Text');
        const stockItemName = escapeXML(getStockItemName(material, shortText));

        const qty = Math.abs(parseFloat(getRowValue(item, 'Qty in Un. of Entry')) || 0);
        const unit = escapeXML(String(getRowValue(item, 'Unit of Entry') || 'Nos').trim());
        const amount = Math.abs(parseFloat(getRowValue(item, 'Amount in LC')) || 0);
        const price = qty > 0 ? (amount / qty) : 0;

        const rateFormatted = `${price.toFixed(2)}/${unit}`;
        const amountFormatted = `-${amount.toFixed(2)}`;
        const qtyFormatted = ` ${formatQuantity(qty)} ${unit}`;
        const destGodown = escapeXML(String(getRowValue(item, 'Receiving Plant') || '').split('.')[0].trim() || receivingPlant);

        return `       <INVENTORYENTRIESIN.LIST>
        <STOCKITEMNAME>${stockItemName}</STOCKITEMNAME>
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
         <GODOWNNAME>${destGodown}</GODOWNNAME>
         <BATCHNAME>Primary Batch</BATCHNAME>
         <DESTINATIONGODOWNNAME>${destGodown}</DESTINATIONGODOWNNAME>
         <INDENTNO>&#4; Not Applicable</INDENTNO>
         <ORDERNO>&#4; Not Applicable</ORDERNO>
         <TRACKINGNUMBER>&#4; Not Applicable</TRACKINGNUMBER>
         <DYNAMICCSTISCLEARED>No</DYNAMICCSTISCLEARED>
         <AMOUNT>${amountFormatted}</AMOUNT>
         <ACTUALQTY>${qtyFormatted}</ACTUALQTY>
         <BILLEDQTY>${qtyFormatted}</BILLEDQTY>
         <ADDITIONALDETAILS.LIST>        </ADDITIONALDETAILS.LIST>
         <VOUCHERCOMPONENTLIST.LIST>        </VOUCHERCOMPONENTLIST.LIST>
        </BATCHALLOCATIONS.LIST>
        <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>
        <RATEDETAILS.LIST>       </RATEDETAILS.LIST>
        <SUPPLEMENTARYDUTYHEADDETAILS.LIST>       </SUPPLEMENTARYDUTYHEADDETAILS.LIST>
        <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>
        <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>
        <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>
        <EXCISEALLOCATIONS.LIST>       </EXCISEALLOCATIONS.LIST>
        <EXPENSEALLOCATIONS.LIST>       </EXPENSEALLOCATIONS.LIST>
       </INVENTORYENTRIESIN.LIST>`;
    }).join('\n');

    // Generate OUT inventory entries (deemed positive: No)
    const inventoryOutXML = transferItems.map(item => {
        const material = getRowValue(item, 'Material');
        const shortText = getRowValue(item, 'Material Description') || getRowValue(item, 'Purchase Order - Short Text') || getRowValue(item, 'Short Text') || getRowValue(item, 'Text');
        const stockItemName = escapeXML(getStockItemName(material, shortText));

        const qty = Math.abs(parseFloat(getRowValue(item, 'Qty in Un. of Entry')) || 0);
        const unit = escapeXML(String(getRowValue(item, 'Unit of Entry') || 'Nos').trim());
        const amount = Math.abs(parseFloat(getRowValue(item, 'Amount in LC')) || 0);
        const price = qty > 0 ? (amount / qty) : 0;

        const rateFormatted = `${price.toFixed(2)}/${unit}`;
        const amountFormatted = `${amount.toFixed(2)}`;
        const qtyFormatted = ` ${formatQuantity(qty)} ${unit}`;
        const sourceGodown = escapeXML(String(getRowValue(item, 'Plant') || '').split('.')[0].trim());

        return `       <INVENTORYENTRIESOUT.LIST>
        <STOCKITEMNAME>${stockItemName}</STOCKITEMNAME>
        <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
        <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
        <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
        <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
        <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
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
         <GODOWNNAME>${sourceGodown}</GODOWNNAME>
         <BATCHNAME>Primary Batch</BATCHNAME>
         <INDENTNO>&#4; Not Applicable</INDENTNO>
         <ORDERNO>&#4; Not Applicable</ORDERNO>
         <TRACKINGNUMBER>&#4; Not Applicable</TRACKINGNUMBER>
         <DYNAMICCSTISCLEARED>No</DYNAMICCSTISCLEARED>
         <AMOUNT>${amountFormatted}</AMOUNT>
         <ACTUALQTY>${qtyFormatted}</ACTUALQTY>
         <BILLEDQTY>${qtyFormatted}</BILLEDQTY>
         <ADDITIONALDETAILS.LIST>        </ADDITIONALDETAILS.LIST>
         <VOUCHERCOMPONENTLIST.LIST>        </VOUCHERCOMPONENTLIST.LIST>
        </BATCHALLOCATIONS.LIST>
        <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>
        <RATEDETAILS.LIST>       </RATEDETAILS.LIST>
        <SUPPLEMENTARYDUTYHEADDETAILS.LIST>       </SUPPLEMENTARYDUTYHEADDETAILS.LIST>
        <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>
        <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>
        <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>
        <EXCISEALLOCATIONS.LIST>       </EXCISEALLOCATIONS.LIST>
        <EXPENSEALLOCATIONS.LIST>       </EXPENSEALLOCATIONS.LIST>
       </INVENTORYENTRIESOUT.LIST>`;
    }).join('\n');

    const remoteId = `81f73e2b-a3c5-4ff2-a56f-49d15ff7c0f6-${grnNumber.padStart(8, '0')}`;
    const vchKey = `81f73e2b-a3c5-4ff2-a56f-49d15ff7c0f6-0000b097:${grnNumber.padStart(8, '0')}`;

    return `<ENVELOPE>
 <HEADER>
  <TALLYREQUEST>Import Data</TALLYREQUEST>
 </HEADER>
 <BODY>
  <IMPORTDATA>
   <REQUESTDESC>
    <REPORTNAME>Vouchers</REPORTNAME>
    <STATICVARIABLES>
     <SVCURRENTCOMPANY>${escapeXML(COMPANY_NAME)}</SVCURRENTCOMPANY>
    </STATICVARIABLES>
   </REQUESTDESC>
   <REQUESTDATA>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <VOUCHER REMOTEID="${remoteId}" VCHKEY="${vchKey}" VCHTYPE="Stock Journal" ACTION="Create" OBJVIEW="Consumption Voucher View">
      <OLDAUDITENTRYIDS.LIST TYPE="Number">
       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
      </OLDAUDITENTRYIDS.LIST>
      <DATE>${docDateFormatted}</DATE>
      <VCHSTATUSDATE>${docDateFormatted}</VCHSTATUSDATE>
      <GUID>${remoteId}</GUID>
      <OBJECTUPDATEACTION>Alter</OBJECTUPDATEACTION>
      <GSTREGISTRATION>&#4; Not Applicable</GSTREGISTRATION>
      <VOUCHERTYPENAME>Stock Journal</VOUCHERTYPENAME>
      <VOUCHERNUMBER>${grnNumber}</VOUCHERNUMBER>
      <NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
      <CSTFORMISSUETYPE>&#4; Not Applicable</CSTFORMISSUETYPE>
      <CSTFORMRECVTYPE>&#4; Not Applicable</CSTFORMRECVTYPE>
      <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>
      <PERSISTEDVIEW>Consumption Voucher View</PERSISTEDVIEW>
      <VCHSTATUSTAXADJUSTMENT>Default</VCHSTATUSTAXADJUSTMENT>
      <VCHSTATUSVOUCHERTYPE>Stock Journal</VCHSTATUSVOUCHERTYPE>
      <VCHGSTCLASS>&#4; Not Applicable</VCHGSTCLASS>
      <VCHENTRYMODE>Use for Stock Journal</VCHENTRYMODE>
      <DESTINATIONGODOWN>${receivingPlant}</DESTINATIONGODOWN>
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
      <CHANGEVCHMODE>No</CHANGEVCHMODE>
      <RESETIRNQRCODE>No</RESETIRNQRCODE>
      <EWAYBILLDETAILS.LIST>      </EWAYBILLDETAILS.LIST>
      <EXCLUDEDTAXATIONS.LIST>      </EXCLUDEDTAXATIONS.LIST>
      <OLDAUDITENTRIES.LIST>      </OLDAUDITENTRIES.LIST>
      <ACCOUNTAUDITENTRIES.LIST>      </ACCOUNTAUDITENTRIES.LIST>
      <AUDITENTRIES.LIST>      </AUDITENTRIES.LIST>
      <DUTYHEADDETAILS.LIST>      </DUTYHEADDETAILS.LIST>
      <GSTADVADJDETAILS.LIST>      </GSTADVADJDETAILS.LIST>
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
${inventoryInXML}
${inventoryOutXML}
      <GST.LIST>      </GST.LIST>
      <STKJRNLADDLCOSTDETAILS.LIST>      </STKJRNLADDLCOSTDETAILS.LIST>
      <PAYROLLMODEOFPAYMENT.LIST>      </PAYROLLMODEOFPAYMENT.LIST>
      <ATTDRECORDS.LIST>      </ATTDRECORDS.LIST>
      <GSTEWAYCONSIGNORADDRESS.LIST>      </GSTEWAYCONSIGNORADDRESS.LIST>
      <GSTEWAYCONSIGNEEADDRESS.LIST>      </GSTEWAYCONSIGNEEADDRESS.LIST>
      <TEMPGSTRATEDETAILS.LIST>      </TEMPGSTRATEDETAILS.LIST>
      <TEMPGSTADVADJUSTED.LIST>      </TEMPGSTADVADJUSTED.LIST>
      <GSTBUYERADDRESS.LIST>      </GSTBUYERADDRESS.LIST>
      <GSTCONSIGNEEADDRESS.LIST>      </GSTCONSIGNEEADDRESS.LIST>
     </VOUCHER>
    </TALLYMESSAGE>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <COMPANY>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>81f73e2b-a3c5-4ff2-a56f-49d15ff7c0f6</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>2786887a-d92f-46c7-bf13-d8a373da8523</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
     </COMPANY>
    </TALLYMESSAGE>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <COMPANY>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>81f73e2b-a3c5-4ff2-a56f-49d15ff7c0f6</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
       <REMOTECMPSTATE>${cmpState}</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
      <REMOTECMPINFO.LIST MERGE="Yes">
       <NAME>2786887a-d92f-46c7-bf13-d8a373da8523</NAME>
       <REMOTECMPNAME>${escapeXML(COMPANY_NAME)}</REMOTECMPNAME>
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
    formatQuantity,
    generateTallyXML,
    generateGRNTallyXML,
    generateStockJournalTallyXML,
    generatePurchaseTallyXML,
    loadPOMaster
};
