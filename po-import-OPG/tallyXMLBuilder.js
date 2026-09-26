const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const config = require('./config');

let conditionTypeMap = null;
const COMPANY_NAME = config.COMPANY_NAME;

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
        const reservedWords = ['po', 'doc', 'document', 'purchasing', 'material', 'invoicing', 'vendor', 'reference', 'item', 'order', 'quantity', 'amount', 'unit', 'plant', 'stock', 'total'];
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
        const reservedWords = ['po', 'doc', 'document', 'purchasing', 'material', 'invoicing', 'vendor', 'reference', 'item', 'order', 'quantity', 'amount', 'unit', 'plant', 'stock', 'total'];
        if (code.length >= 3 && !reservedWords.includes(code)) {
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

    // Check if row has dedicated charge vendor mapping from vendor mapping file
    if (row._chargeVendors) {
        const code = getConditionCode(colName);
        if (code && row._chargeVendors[code]) {
            return row._chargeVendors[code];
        }
        const cleanCol = colName.toLowerCase().replace(/\s/g, '');
        for (const [k, v] of Object.entries(row._chargeVendors)) {
            if (k.toLowerCase().replace(/\s/g, '') === cleanCol) {
                return v;
            }
        }
    }

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
 * Format date for ORDERDUEDATE tag with attributes (JD and P)
 */
function formatOrderDueDate(dateValue) {
    let dt;
    if (!dateValue) {
        dt = new Date(2026, 3, 1);
    } else if (typeof dateValue === 'number' || (typeof dateValue === 'string' && dateValue.trim() !== '' && !isNaN(dateValue) && !dateValue.includes('.'))) {
        const serial = parseFloat(dateValue);
        dt = new Date(Math.round((serial - 25569) * 86400 * 1000));
    } else if (dateValue instanceof Date) {
        dt = dateValue;
    } else {
        const str = String(dateValue).trim();
        if (str.includes('.')) {
            const parts = str.split('.');
            if (parts.length === 3) {
                dt = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
            }
        }
        if (!dt || isNaN(dt.getTime())) {
            const parsed = new Date(str);
            if (!isNaN(parsed.getTime())) {
                dt = parsed;
            } else {
                dt = new Date(2026, 3, 1);
            }
        }
    }

    // Tally Julian Date epoch: 1-Jan-1900 = 1
    const epoch = new Date(1900, 0, 1);
    const diffDays = Math.floor((dt.getTime() - epoch.getTime()) / (86400 * 1000));
    const jd = diffDays + 1;

    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const day = dt.getDate();
    const month = months[dt.getMonth()];
    const yearStr = String(dt.getFullYear()).slice(-2);
    const formattedStr = `${day}-${month}-${yearStr}`;

    return `<ORDERDUEDATE JD="${jd}" P="${formattedStr}">${formattedStr}</ORDERDUEDATE>`;
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
    const rawDocType = poGroup.docType ||
        getRowValue(firstRow, 'PO -  Doc Type') ||
        getRowValue(firstRow, 'PO - Doc Type') ||
        getRowValue(firstRow, 'Purchasing Doc Type') ||
        getRowValue(firstRow, 'Purchasing Doc. Type') ||
        getRowValue(firstRow, 'Doc Type') ||
        getRowValue(firstRow, 'Document Type') ||
        'ZSPR';
    const docType = escapeXML(String(rawDocType).trim());
    const docDateFormatted = formatDate(getRowValue(firstRow, 'Document Date') || getRowValue(firstRow, 'Doc Date') || getRowValue(firstRow, 'PO Date') || getRowValue(firstRow, 'Posting Date'));

    const vendorCode = padVendor(getRowValue(firstRow, 'Vendor'));
    const rawVendorName = String(getRowValue(firstRow, 'Vendor Name') || (vendorCode ? vendorMap[vendorCode] || '' : '')).trim();
    const vendorName = escapeXML(rawVendorName || vendorCode);
    const partyName = escapeXML(vendorCode ? (rawVendorName ? `${vendorCode}-${rawVendorName}` : vendorCode) : rawVendorName);

    const street = String(getRowValue(firstRow, 'Street') || '').trim();
    const city = String(getRowValue(firstRow, 'City') || '').trim();
    const address = escapeXML(street && city ? `${street},${city}${city}` : (street || city));

    const postCode = escapeXML(getRowValue(firstRow, 'Post Code') ? String(getRowValue(firstRow, 'Post Code')).split('.')[0].trim() : '');
    const gstNo = escapeXML(String(getRowValue(firstRow, 'GST NO') || '').trim());
    const destRegion = String(getRowValue(firstRow, 'Destination region') || getRowValue(firstRow, 'Destination Region') || getRowValue(firstRow, 'Region') || '').trim();
    const isGst33OrBlank = destRegion ? (destRegion === '33' || destRegion.startsWith('33')) : (!gstNo || gstNo.startsWith('33'));
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

    const isZcol = docType === 'ZCOL';

    const itemsXML = activeItems.map(item => {
        const material = getRowValue(item, 'Material');
        const shortText = getRowValue(item, 'Purchase Order Text') ||
            getRowValue(item, 'Material Description') ||
            getRowValue(item, 'Purchase Order Line Item Text') ||
            getRowValue(item, 'Purchase Order - Short Text') ||
            getRowValue(item, 'Short Text') ||
            getRowValue(item, 'Text');
        const stockItemName = escapeXML(getStockItemName(material, shortText));

        const currency = String(getRowValue(item, 'Currency') || 'INR').trim().toUpperCase();
        const exRateVal = parseFloat(getRowValue(item, 'Exchange Rate')) || 1;
        const exchangeRate = (currency !== 'INR' && exRateVal > 0) ? exRateVal : 1;

        const qty = parseFloat(getRowValue(item, 'Order Quantity')) || 0;
        const unit = escapeXML(String(getRowValue(item, 'Order Unit') || 'Nos').trim());

        let rawPrice = parseFloat(getRowValue(item, 'Net Order Price')) || 0;
        let rawAmount = parseFloat(getRowValue(item, 'Net Order Value')) || (qty * rawPrice);

        // For ZCOL: directly use Effective Value for Item Amount and calculate Rate based on Effective Value
        if (isZcol) {
            const effectiveVal = parseFloat(getRowValue(item, 'Effective value') || getRowValue(item, 'Effective Value'));
            if (!isNaN(effectiveVal) && effectiveVal > 0) {
                rawAmount = effectiveVal;
                rawPrice = qty > 0 ? (effectiveVal / qty) : rawPrice;
            }
        }

        const price = rawPrice * exchangeRate;
        const amount = rawAmount * exchangeRate;

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
        const condCode = getConditionCode(col);

        // Check if any active item has charge values from the vendor mapping file for this condition type.
        // The mapping file provides INR-converted values (Condition value * Cond.exchange rate for non-INR).
        const hasVendorMappedValues = activeItems.some(item =>
            item._chargeValues && item._chargeValues[condCode] !== undefined
        );

        if (hasVendorMappedValues) {
            // Sum the INR-converted condition values from the mapping file across all items
            activeItems.forEach(item => {
                if (item._chargeValues && item._chargeValues[condCode] !== undefined) {
                    colSum += item._chargeValues[condCode];
                }
            });
        } else {
            // Fall back: sum the column values from the main PO excel (no exchange rate here;
            // exchange rate for PO line items is handled separately in the inventory section)
            activeItems.forEach(item => {
                const rawVal = getRowValue(item, col);
                if (rawVal !== undefined && rawVal !== null) {
                    const cleanVal = String(rawVal).replace(/,/g, '').trim();
                    colSum += parseFloat(cleanVal) || 0;
                }
            });
        }


        if (Math.abs(colSum) > 0.001) {
            const cleanColName = col.toLowerCase().replace(/\s/g, '');

            const exactKey = getExactKey(firstRow, col);
            const mappedName = condMap[condCode];

            const finalLedgerName = mappedName ? mappedName : exactKey.toString().replace(/\s+/g, ' ').trim();
            if (finalLedgerName.toLowerCase().replace(/\s/g, '') === 'grossprice') {
                return;
            }
            const normalizedName = escapeXML(finalLedgerName);

            // If ZCOL: ALL charges, cess, and taxes (JEXS, NAVS, etc.) are passed inside UDF fields with NO exceptions.
            // No external tax or charge ledgers are created outside for ZCOL.
            if (isZcol) {
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
                // Non-ZCOL POs: Split JEXS/NAVS into CGST/SGST/IGST or create external ledger entries
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

                activeLedgers.push({
                    name: normalizedName,
                    sum: colSum
                });
                totalTaxesAndCharges += colSum;
            }
        }
    });

    const totalVoucherAmount = totalNetValue + (isZcol ? 0 : totalTaxesAndCharges);

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
            const lower = file.toLowerCase();
            if (lower.endsWith('.xlsx') && !lower.includes('grn') && !lower.includes('invoice') && !lower.includes('fi data') && !lower.includes('sales') && !lower.includes('book')) {
                const filePath = path.join(parentDir, file);
                try {
                    const stats = fs.statSync(filePath);
                    if (stats.size > 15 * 1024 * 1024) continue; // Skip files > 15MB to prevent event loop block
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
    const firstRow = purchaseGroup.items[0];

    const rawVoucherNo = getRowValue(firstRow, 'Document Number') ||
        getRowValue(firstRow, 'Invoice No') ||
        getRowValue(firstRow, 'Invoice Number') ||
        purchaseGroup.poNumber ||
        '';
    const voucherNumber = escapeXML(String(rawVoucherNo).split('.')[0].trim());
    const reference = escapeXML(String(getRowValue(firstRow, 'Reference')).trim());
    const docDateFormatted = formatDate(getRowValue(firstRow, 'Document Date') || getRowValue(firstRow, 'Posting Date'));

    const firstRowDocType = String(
        getRowValue(firstRow, 'Purchasing Doc Type') ||
        getRowValue(firstRow, 'Purchasing Doc. Type') ||
        getRowValue(firstRow, 'PO - Doc Type') ||
        getRowValue(firstRow, 'Doc Type') ||
        getRowValue(firstRow, 'Purchase Order Type') ||
        ''
    ).trim();

    const headerDocType = firstRowDocType || 'ZSPR';
    const voucherTypeName = `Purchase ${headerDocType}`;

    const rawVendorCode = getRowValue(firstRow, 'Invoicing Party') || getRowValue(firstRow, 'Vendor');
    const vendorCode = padVendor(rawVendorCode);
    const rawVendorName = String(getRowValue(firstRow, 'Vendor Name') || (vendorCode ? (vendorMap[vendorCode] || vendors[vendorCode]?.vendorName || '') : '')).trim();
    const finalVendorName = vendorCode ? (rawVendorName ? `${vendorCode}-${rawVendorName}` : vendorCode) : (rawVendorName || 'Unknown Vendor');
    const vendorName = escapeXML(rawVendorName || vendorCode || 'Unknown Vendor');
    const partyName = escapeXML(finalVendorName);

    const street = String(vendors[vendorCode]?.street || getRowValue(firstRow, 'Street') || '').trim();
    const city = String(vendors[vendorCode]?.city || getRowValue(firstRow, 'City') || '').trim();
    const address = escapeXML(street && city ? `${street},${city}` : (street || city || ''));

    const postCode = escapeXML(vendors[vendorCode]?.postCode || getRowValue(firstRow, 'Post Code') || '');
    const gstNo = escapeXML(vendors[vendorCode]?.gstNo || getRowValue(firstRow, 'GST NO') || '');
    const destRegion = String(getRowValue(firstRow, 'Destination region') || getRowValue(firstRow, 'Destination Region') || getRowValue(firstRow, 'Region') || '').trim();
    const isGst33OrBlank = destRegion ? (destRegion === '33' || destRegion.startsWith('33')) : (!gstNo || gstNo.startsWith('33'));
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
        const shortText = getRowValue(item, 'Purchase Order Text') ||
            getRowValue(item, 'Material Description') ||
            getRowValue(item, 'Purchase Order Line Item Text') ||
            getRowValue(item, 'Purchase Order - Short Text') ||
            getRowValue(item, 'Text') ||
            getRowValue(item, 'Short Text');
        const stockItemName = escapeXML(getStockItemName(material, shortText));

        const currency = String(getRowValue(item, 'Currency') || 'INR').trim().toUpperCase();
        const exRateVal = parseFloat(getRowValue(item, 'Exchange Rate')) || 1;
        const exchangeRate = (currency !== 'INR' && exRateVal > 0) ? exRateVal : 1;

        const qty = parseFloat(getRowValue(item, 'Quantity') || getRowValue(item, 'Qty in OPUn')) || 0;
        const unit = escapeXML(String(getRowValue(item, 'Order Unit') || getRowValue(item, 'Order Price Unit') || getRowValue(item, 'Base Unit of Measure') || 'Nos').trim());
        let rawAmount = parseFloat(getRowValue(item, 'Amount') || getRowValue(item, 'Total Value')) || 0;
        const amount = rawAmount * exchangeRate;
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
            getRowValue(firstRow, 'Purchase Order Type') ||
            ''
        ).trim() || headerDocType;
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
     <VOUCHER VCHTYPE="${voucherTypeName}" ACTION="Create" OBJVIEW="Invoice Voucher View">
      <ADDRESS.LIST TYPE="String">
       <ADDRESS>${address}</ADDRESS>
      </ADDRESS.LIST>
      <OLDAUDITENTRYIDS.LIST TYPE="Number">
       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
      </OLDAUDITENTRYIDS.LIST>
      <DATE>${docDateFormatted}</DATE>
      <REFERENCEDATE>${docDateFormatted}</REFERENCEDATE>
      <VCHSTATUSDATE>${docDateFormatted}</VCHSTATUSDATE>
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
            const lower = file.toLowerCase();
            if (lower.endsWith('.xlsx') && !lower.includes('grn') && !lower.includes('fi data') && !lower.includes('sales') && !lower.includes('book')) {
                const filePath = path.join(parentDir, file);
                try {
                    const stats = fs.statSync(filePath);
                    if (stats.size > 15 * 1024 * 1024) continue; // Skip files > 15MB to prevent event loop block
                    const workbook = xlsx.readFile(filePath);
                    for (const sheetName of workbook.SheetNames) {
                        if (sheetName.toLowerCase().includes('detail') || sheetName.toLowerCase().includes('sheet')) {
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
    const destRegion = String(getRowValue(firstRow, 'Destination region') || getRowValue(firstRow, 'Destination Region') || getRowValue(firstRow, 'Region') || '').trim();
    const isGst33OrBlank = destRegion ? (destRegion === '33' || destRegion.startsWith('33')) : (!gstNo || gstNo.startsWith('33'));
    const regionName = escapeXML(cleanStateName(masterVendor.regionName || 'Tamil Nadu'));

    const cmpState = 'Tamil Nadu';
    const isLocal = regionName.toLowerCase().replace(/\s/g, '') === cmpState.toLowerCase().replace(/\s/g, '');

    // Extract Doc Type from GRN row
    const rawDocType = String(
        getRowValue(firstRow, 'Doc Type') ||
        getRowValue(firstRow, 'Doc. Type') ||
        getRowValue(firstRow, 'Document Type') ||
        getRowValue(firstRow, 'Purchase Order type') ||
        getRowValue(firstRow, 'PO Type') ||
        getRowValue(firstRow, 'Purchasing Doc Type') ||
        ''
    ).trim();
    const docType = rawDocType || 'ZSPR';

    // Determine Voucher Type: Receipt Note WE <DocType> (e.g. Receipt Note WE ZSPR, Receipt Note WE ZSTO)
    const voucherType = docType ? `Receipt Note ${eventType} ${docType}` : `Receipt Note ${eventType}`;

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
        const shortText = getRowValue(item, 'New Material Desc') || getRowValue(item, 'Material Description') || getRowValue(item, 'Purchase Order Line Item Text') || getRowValue(item, 'Purchase Order - Short Text') || getRowValue(item, 'Short Text') || getRowValue(item, 'Text');

        // Stock Item Name is the material column value (acting as an alias in Tally)
        const stockItemName = escapeXML(getStockItemName(material, shortText));

        const qty = parseFloat(getRowValue(item, 'Qty in Un. of Entry') || getRowValue(item, 'Order Quantity')) || 0;
        const unit = escapeXML(String(getRowValue(item, 'Unit of Entry') || getRowValue(item, 'Order Unit') || 'Nos').trim());

        // Amount in LC in Excel is tax-inclusive. We divide by 1.18 to get the net amount.
        // Amount in LC: sending direct Amount in LC for line item without 1.18 division / separate tax
        const amountLC = parseFloat(getRowValue(item, 'Amount in LC')) || 0;
        // const amount = amountLC / 1.18; // commented out
        const amount = amountLC;
        const price = qty > 0 ? (amount / qty) : 0;

        totalNetValue += amount;

        const rateFormatted = `${price.toFixed(2)}/${unit}`;
        const amountFormatted = `-${amount.toFixed(2)}`;
        const qtyFormatted = ` ${formatQuantity(qty)} ${unit}`;

        // For GRN, ledger name is Purchase <DocType>
        const itemDocType = String(
            getRowValue(item, 'Doc Type') ||
            getRowValue(item, 'Doc. Type') ||
            getRowValue(item, 'Document Type') ||
            getRowValue(item, 'Purchase Order type') ||
            getRowValue(item, 'PO Type') ||
            getRowValue(item, 'Purchasing Doc Type') ||
            docType
        ).trim();
        const ledgerName = `Purchase ${itemDocType}`;

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

    // Tax calculation (commented out to send direct Amount in LC without separate tax ledgers)
    const activeLedgers = [];
    let totalTaxesAndCharges = 0;

    /*
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
    */

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
      <NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
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

    // Extract event/trans type (e.g. WA, WE, WQ, WI)
    const eventType = String(getRowValue(firstRow, 'Trans./Event Type') || getRowValue(firstRow, 'Trans./Event TypeA') || 'WA').trim().toUpperCase();

    // Extract Doc Type from row; if blank, #N/A, or empty, default to ZSTO under its respective Trans./Event Type
    let rawDocType = String(
        getRowValue(firstRow, 'Doc Type') ||
        getRowValue(firstRow, 'Doc. Type') ||
        getRowValue(firstRow, 'Purchase Order type') ||
        getRowValue(firstRow, 'PO Type') ||
        getRowValue(firstRow, 'Purchasing Doc Type') ||
        ''
    ).trim();

    // If still empty, check 'Document Type' only if it's not the same as eventType (like WA or WL)
    if (!rawDocType) {
        const dt = String(getRowValue(firstRow, 'Document Type') || '').trim();
        if (dt && dt.toUpperCase() !== eventType) {
            rawDocType = dt;
        }
    }

    const isInvalidDocType = !rawDocType || rawDocType.toUpperCase() === '#N/A' || rawDocType.toUpperCase() === 'N/A' || rawDocType.toUpperCase() === 'NAN' || rawDocType.toUpperCase() === eventType;
    const effectiveDocType = isInvalidDocType ? 'ZSTO' : rawDocType;

    // Determine Voucher Type: Stock Journal <TransType> <DocType> (e.g. Stock Journal WA ZSTO)
    const voucherType = `Stock Journal ${eventType} ${effectiveDocType}`;

    const activeItems = grnGroup.items.filter(item => {
        const delInd = getRowValue(item, 'Deletion Indicator');
        return !(delInd && String(delInd).trim().toUpperCase() === 'L');
    });

    if (activeItems.length === 0) {
        return '';
    }

    // Separate items by 'Debit/Credit Ind.':
    // 'S' = Debit (Incoming / Production -> INVENTORYENTRIESIN)
    // 'H' = Credit (Outgoing / Consumption -> INVENTORYENTRIESOUT)
    const sItems = activeItems.filter(item => {
        const dc = String(getRowValue(item, 'Debit/Credit Ind.') || getRowValue(item, 'Debit/Credit') || '').trim().toUpperCase();
        return dc === 'S';
    });
    const hItems = activeItems.filter(item => {
        const dc = String(getRowValue(item, 'Debit/Credit Ind.') || getRowValue(item, 'Debit/Credit') || '').trim().toUpperCase();
        return dc === 'H';
    });

    let inItems = [];
    let outItems = [];

    if (sItems.length > 0 || hItems.length > 0) {
        // If only S exists, inItems has entries and outItems remains empty (pass as it is)
        // If only H exists, outItems has entries and inItems remains empty (pass as it is)
        // If both exist, both are populated
        inItems = sItems;
        outItems = hItems;
    } else {
        // Fallback if neither S nor H is specified: inspect quantity sign
        inItems = activeItems.filter(item => (parseFloat(getRowValue(item, 'Qty in Un. of Entry')) || 0) > 0);
        outItems = activeItems.filter(item => (parseFloat(getRowValue(item, 'Qty in Un. of Entry')) || 0) < 0);
        if (inItems.length === 0 && outItems.length === 0) {
            outItems = activeItems; // default to consumption if unspecified
        }
    }

    // Header destination godown (from the first incoming item)
    const firstInItem = inItems[0] || activeItems[0];
    const destinationGodown = escapeXML(String(getRowValue(firstInItem, 'Plant') || getRowValue(firstInItem, 'Receiving Plant') || '').split('.')[0].trim());
    const cmpState = 'Tamil Nadu';

    const inventoryInXML = inItems.map(item => {
        const material = getRowValue(item, 'Material');
        const shortText = getRowValue(item, 'New Material Desc') || getRowValue(item, 'Material Description') || getRowValue(item, 'Purchase Order Line Item Text') || getRowValue(item, 'Purchase Order - Short Text') || getRowValue(item, 'Short Text') || getRowValue(item, 'Text');
        const stockItemName = escapeXML(getStockItemName(material, shortText));

        const qty = Math.abs(parseFloat(getRowValue(item, 'Qty in Un. of Entry')) || 0);
        const unit = escapeXML(String(getRowValue(item, 'Unit of Entry') || 'Nos').trim());
        const amount = Math.abs(parseFloat(getRowValue(item, 'Amount in LC')) || 0);
        const price = qty > 0 ? (amount / qty) : 0;

        const rateFormatted = `${price.toFixed(2)}/${unit}`;
        const amountFormatted = `-${amount.toFixed(2)}`;
        const qtyFormatted = ` ${formatQuantity(qty)} ${unit}`;
        const inGodown = escapeXML(String(getRowValue(item, 'Plant') || destinationGodown).split('.')[0].trim());

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
         <GODOWNNAME>${inGodown}</GODOWNNAME>
         <BATCHNAME>Primary Batch</BATCHNAME>
         <DESTINATIONGODOWNNAME>${inGodown}</DESTINATIONGODOWNNAME>
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

    const inventoryOutXML = outItems.map(item => {
        const material = getRowValue(item, 'Material');
        const shortText = getRowValue(item, 'New Material Desc') || getRowValue(item, 'Material Description') || getRowValue(item, 'Purchase Order Line Item Text') || getRowValue(item, 'Purchase Order - Short Text') || getRowValue(item, 'Short Text') || getRowValue(item, 'Text');
        const stockItemName = escapeXML(getStockItemName(material, shortText));

        const qty = Math.abs(parseFloat(getRowValue(item, 'Qty in Un. of Entry')) || 0);
        const unit = escapeXML(String(getRowValue(item, 'Unit of Entry') || 'Nos').trim());
        const amount = Math.abs(parseFloat(getRowValue(item, 'Amount in LC')) || 0);
        const price = qty > 0 ? (amount / qty) : 0;

        const rateFormatted = `${price.toFixed(2)}/${unit}`;
        const amountFormatted = `${amount.toFixed(2)}`;
        const qtyFormatted = ` ${formatQuantity(qty)} ${unit}`;
        const outGodown = escapeXML(String(getRowValue(item, 'Plant') || '').split('.')[0].trim());

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
         <GODOWNNAME>${outGodown}</GODOWNNAME>
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
     <VOUCHER REMOTEID="${remoteId}" VCHKEY="${vchKey}" VCHTYPE="${voucherType}" ACTION="Create" OBJVIEW="Consumption Voucher View">
      <OLDAUDITENTRYIDS.LIST TYPE="Number">
       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
      </OLDAUDITENTRYIDS.LIST>
      <DATE>${docDateFormatted}</DATE>
      <VCHSTATUSDATE>${docDateFormatted}</VCHSTATUSDATE>
      <GUID>${remoteId}</GUID>
      <OBJECTUPDATEACTION>Alter</OBJECTUPDATEACTION>
      <GSTREGISTRATION>&#4; Not Applicable</GSTREGISTRATION>
      <VOUCHERTYPENAME>${voucherType}</VOUCHERTYPENAME>
      <VOUCHERNUMBER>${grnNumber}</VOUCHERNUMBER>
      <NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
      <CSTFORMISSUETYPE>&#4; Not Applicable</CSTFORMISSUETYPE>
      <CSTFORMRECVTYPE>&#4; Not Applicable</CSTFORMRECVTYPE>
      <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>
      <PERSISTEDVIEW>Consumption Voucher View</PERSISTEDVIEW>
      <VCHSTATUSTAXADJUSTMENT>Default</VCHSTATUSTAXADJUSTMENT>
      <VCHSTATUSVOUCHERTYPE>${voucherType}</VCHSTATUSVOUCHERTYPE>
      <VCHGSTCLASS>&#4; Not Applicable</VCHGSTCLASS>
      <VCHENTRYMODE>Use for Stock Journal</VCHENTRYMODE>
      <DESTINATIONGODOWN>${destinationGodown}</DESTINATIONGODOWN>
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

function generateSalesOrderTallyXML(poGroup) {
    const items = poGroup.items || [];
    const firstRow = items[0] || {};

    const rawSalesDoc = getRowValue(firstRow, 'Sales document') || poGroup.poNumber || '';
    const voucherNumber = String(rawSalesDoc).split('.')[0].trim();

    const salesDocType = String(getRowValue(firstRow, 'Sales Document Type') || getRowValue(firstRow, 'Sales Document Type Desc') || 'ZASH').trim();
    const voucherType = `Sales Order ${salesDocType}`;

    const rawParty = String(getRowValue(firstRow, 'Ship-to party') || getRowValue(firstRow, 'Party') || getRowValue(firstRow, 'Customer') || '').split('.')[0].trim();
    const partyLedger = rawParty ? rawParty.padStart(10, '0') : '';

    const rawDocDate = getRowValue(firstRow, 'Doc Date') || getRowValue(firstRow, 'Created on') || getRowValue(firstRow, 'Document Date');
    const docDateFormatted = formatDate(rawDocDate);

    let totalInventoryAmount = 0;
    let totalTaxAmount = 0;

    const inventoryEntriesXML = items.map(item => {
        const materialCode = String(getRowValue(item, 'Material') || getRowValue(item, 'Material entered') || '').split('.')[0].trim();
        const stockItemName = materialCode ? getStockItemName(materialCode) : '';

        const itemDocDate = getRowValue(item, 'Doc Date') || getRowValue(item, 'Created on') || getRowValue(item, 'Document Date') || rawDocDate;
        const itemOrderDueDateXML = formatOrderDueDate(itemDocDate);

        const qtyVal = getRowValue(item, 'Order Quantity') || getRowValue(item, 'Quantity') || 0;
        const qtyNum = parseFloat(String(qtyVal).trim()) || 0;
        const uom = String(getRowValue(item, 'Base Unit of Measure') || getRowValue(item, 'Sales unit') || getRowValue(item, 'Unit of measure') || 'MT').trim();
        const qtyFormatted = `${formatQuantity(qtyNum)} ${uom}`;

        // 1. Dynamic base price determination: check ZASH, YBPR, zpro, ZPRS in order
        const p1 = parseFloat(getRowValue(item, 'ZASH - Base Price') || 0) || 0;
        const p2 = parseFloat(getRowValue(item, 'YBPR- Basic Price') || getRowValue(item, 'YBPR - Basic Price') || 0) || 0;
        const p3 = parseFloat(getRowValue(item, 'zpro-base price') || getRowValue(item, 'zpro - base price') || 0) || 0;
        const p4 = parseFloat(getRowValue(item, 'ZPRS - base price') || getRowValue(item, 'ZPRS - Base Price') || 0) || 0;

        let basePrice = 0;
        if (p1 !== 0) basePrice = p1;
        else if (p2 !== 0) basePrice = p2;
        else if (p3 !== 0) basePrice = p3;
        else if (p4 !== 0) basePrice = p4;

        // 2. Check additional component columns
        const colYTRL = parseFloat(getRowValue(item, 'YTRL - Transmission') || 0) || 0;
        const colCESS = parseFloat(getRowValue(item, 'CESS') || 0) || 0;
        const colZASHQty = parseFloat(getRowValue(item, 'ZASH Qty * Base Price') || 0) || 0;
        const colYBPRQty = parseFloat(getRowValue(item, 'YBPR - Qty * Base Price') || 0) || 0;
        const colZProQty = parseFloat(getRowValue(item, 'zpro - Qty * BasePrice') || getRowValue(item, 'zpro - Qty * BasePrice ') || 0) || 0;
        const colZREL = parseFloat(getRowValue(item, 'ZREL') || 0) || 0;
        const colZOM2 = parseFloat(getRowValue(item, 'ZOM2 - Without Base amt') || getRowValue(item, 'ZOM2 - Without Base amt ') || 0) || 0;
        const colJWTH = parseFloat(getRowValue(item, 'JWTH') || 0) || 0;

        const sumAdditionalComponents = colYTRL + colCESS + colZASHQty + colYBPRQty + colZProQty + colZREL + colZOM2 + colJWTH;

        let lineAmount = 0;
        if (sumAdditionalComponents !== 0) {
            lineAmount = sumAdditionalComponents;
        } else {
            lineAmount = basePrice * qtyNum;
        }

        totalInventoryAmount += lineAmount;

        // Tax amount for this item
        const itemTax = parseFloat(getRowValue(item, 'TAX amount') || getRowValue(item, 'TAX Amount') || 0) || 0;
        totalTaxAmount += itemTax;

        const computedRate = qtyNum > 0 ? (lineAmount / qtyNum) : basePrice;
        const rateFormatted = `${computedRate.toFixed(2)}/${uom}`;
        const amountFormatted = lineAmount.toFixed(2);

        return `      <ALLINVENTORYENTRIES.LIST>
       <STOCKITEMNAME>${escapeXML(stockItemName)}</STOCKITEMNAME>
       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
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
        <GODOWNNAME>Main Location</GODOWNNAME>
        <BATCHNAME>Primary Batch</BATCHNAME>
        <INDENTNO>&#4; Not Applicable</INDENTNO>
        <ORDERNO>${voucherNumber}</ORDERNO>
        <TRACKINGNUMBER>&#4; Not Applicable</TRACKINGNUMBER>
        <DYNAMICCSTISCLEARED>No</DYNAMICCSTISCLEARED>
        <AMOUNT>${amountFormatted}</AMOUNT>
        <ACTUALQTY>${qtyFormatted}</ACTUALQTY>
        <BILLEDQTY>${qtyFormatted}</BILLEDQTY>
        ${itemOrderDueDateXML}
        <ADDITIONALDETAILS.LIST>        </ADDITIONALDETAILS.LIST>
        <VOUCHERCOMPONENTLIST.LIST>        </VOUCHERCOMPONENTLIST.LIST>
       </BATCHALLOCATIONS.LIST>
       <ACCOUNTINGALLOCATIONS.LIST>
        <OLDAUDITENTRYIDS.LIST TYPE="Number">
         <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
        </OLDAUDITENTRYIDS.LIST>
        <LEDGERNAME>${escapeXML(voucherType)}</LEDGERNAME>
        <GSTCLASS>&#4; Not Applicable</GSTCLASS>
        <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
        <LEDGERFROMITEM>No</LEDGERFROMITEM>
        <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
        <ISPARTYLEDGER>No</ISPARTYLEDGER>
        <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
        <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
        <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
        <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
        <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
        <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
        <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
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

    // Tax ledger calculations based on Destination region
    const destRegion = String(getRowValue(firstRow, 'Destination region') || '').trim();
    let taxLedgersXML = '';

    if (totalTaxAmount > 0) {
        if (destRegion === '33') {
            const halfTax = totalTaxAmount / 2;
            const halfTaxFormatted = halfTax.toFixed(2);
            taxLedgersXML = `      <LEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <LEDGERNAME>CGST</LEDGERNAME>
       <GSTCLASS>&#4; Not Applicable</GSTCLASS>
       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
       <LEDGERFROMITEM>No</LEDGERFROMITEM>
       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
       <ISPARTYLEDGER>No</ISPARTYLEDGER>
       <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
       <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
       <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
       <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
       <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
       <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
       <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
       <AMOUNT>${halfTaxFormatted}</AMOUNT>
      </LEDGERENTRIES.LIST>
      <LEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <LEDGERNAME>SGST</LEDGERNAME>
       <GSTCLASS>&#4; Not Applicable</GSTCLASS>
       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
       <LEDGERFROMITEM>No</LEDGERFROMITEM>
       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
       <ISPARTYLEDGER>No</ISPARTYLEDGER>
       <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
       <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
       <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
       <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
       <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
       <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
       <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
       <AMOUNT>${halfTaxFormatted}</AMOUNT>
      </LEDGERENTRIES.LIST>`;
        } else {
            const taxFormatted = totalTaxAmount.toFixed(2);
            taxLedgersXML = `      <LEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <LEDGERNAME>IGST</LEDGERNAME>
       <GSTCLASS>&#4; Not Applicable</GSTCLASS>
       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
       <LEDGERFROMITEM>No</LEDGERFROMITEM>
       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
       <ISPARTYLEDGER>No</ISPARTYLEDGER>
       <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
       <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
       <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
       <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
       <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
       <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
       <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
       <AMOUNT>${taxFormatted}</AMOUNT>
      </LEDGERENTRIES.LIST>`;
        }
    }

    const partyTotalAmount = totalInventoryAmount + totalTaxAmount;
    const partyAmountFormatted = (-partyTotalAmount).toFixed(2);
    const remoteId = `81f73e2b-a3c5-4ff2-a56f-49d15ff7c0f6-${voucherNumber.padStart(8, '0')}`;
    const vchKey = `81f73e2b-a3c5-4ff2-a56f-49d15ff7c0f6-0000b125:${voucherNumber.padStart(8, '0')}`;

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
     <VOUCHER REMOTEID="${remoteId}" VCHKEY="${vchKey}" VCHTYPE="${escapeXML(voucherType)}" ACTION="Create" OBJVIEW="Invoice Voucher View">
      <OLDAUDITENTRYIDS.LIST TYPE="Number">
       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
      </OLDAUDITENTRYIDS.LIST>
      <DATE>${docDateFormatted}</DATE>
      <VCHSTATUSDATE>${docDateFormatted}</VCHSTATUSDATE>
      <GUID>${remoteId}</GUID>
      <PARTYNAME>${escapeXML(partyLedger)}</PARTYNAME>
      <VOUCHERTYPENAME>${escapeXML(voucherType)}</VOUCHERTYPENAME>
      <PARTYLEDGERNAME>${escapeXML(partyLedger)}</PARTYLEDGERNAME>
      <VOUCHERNUMBER>${voucherNumber}</VOUCHERNUMBER>
      <BASICBUYERNAME>${escapeXML(partyLedger)}</BASICBUYERNAME>
      <REFERENCE>${voucherNumber}</REFERENCE>
      <NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
      <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
      <VCHSTATUSVOUCHERTYPE>${escapeXML(voucherType)}</VCHSTATUSVOUCHERTYPE>
      <DIFFACTUALQTY>No</DIFFACTUALQTY>
      <ISMSTFROMSYNC>No</ISMSTFROMSYNC>
      <ISDELETED>No</ISDELETED>
      <ASORIGINAL>No</ASORIGINAL>
      <AUDITED>No</AUDITED>
      <ISOPTIONAL>No</ISOPTIONAL>
      <EFFECTIVEDATE>${docDateFormatted}</EFFECTIVEDATE>
      <USETRACKINGNUMBER>No</USETRACKINGNUMBER>
      <ISINVOICE>No</ISINVOICE>
      <ISVATDUTYPAID>Yes</ISVATDUTYPAID>
${inventoryEntriesXML}
      <LEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <LEDGERNAME>${escapeXML(partyLedger)}</LEDGERNAME>
       <GSTCLASS>&#4; Not Applicable</GSTCLASS>
       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
       <LEDGERFROMITEM>No</LEDGERFROMITEM>
       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
       <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
       <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
       <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
       <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
       <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
       <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
       <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
       <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>
       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
       <AMOUNT>${partyAmountFormatted}</AMOUNT>
      </LEDGERENTRIES.LIST>
${taxLedgersXML}
     </VOUCHER>
    </TALLYMESSAGE>
   </REQUESTDATA>
  </IMPORTDATA>
 </BODY>
</ENVELOPE>`;
}

function generateFITallyXML(poGroup) {
    if (!poGroup || !poGroup.items || poGroup.items.length === 0) {
        throw new Error('No items in Financial Entry group');
    }

    const items = poGroup.items;
    const firstRow = items[0];

    const rawDocType = String(
        getRowValue(firstRow, 'Doc Type') ||
        getRowValue(firstRow, 'Doc type') ||
        getRowValue(firstRow, 'Document Type') ||
        'FI'
    ).trim();
    const voucherTypeName = escapeXML(`Journal ${rawDocType}`);
    const docNumber = String(poGroup.poNumber || getRowValue(firstRow, 'Document Number') || '').split('.')[0].trim();
    // const guid = `81f73e2b-a3c5-4ff2-a56f-50f15ff7c0f6-${docNumber.padStart(8, '0')}-${rawDocType}`;
    // const guid = `81f73e2b-a3c5-4ff2-a56f-57d15ff7c0f6-${docNumber.padStart(8, '0')}-${rawDocType}`;

    const postingDateRaw = getRowValue(firstRow, 'Posting Date') || getRowValue(firstRow, 'Doc Date');
    const dateFormatted = formatDate(postingDateRaw);

    let partyLedger = '';
    let partyAmount = 0;
    const ledgerEntries = [];

    items.forEach(row => {
        const ind = String(getRowValue(row, 'Debit/Credit Ind.') || getRowValue(row, 'Debit/Credit Ind') || 'S').trim().toUpperCase();
        const isCredit = ind === 'H';
        const isDeemedPositive = isCredit ? 'No' : 'Yes'; // S = Debit (Yes), H = Credit (No)

        const rawVendor = getRowValue(row, 'Vendor');
        const cleanVendor = rawVendor !== undefined && rawVendor !== null ? String(rawVendor).split('.')[0].trim() : '';
        const rawCustomer = getRowValue(row, 'Customer');
        const cleanCustomer = rawCustomer !== undefined && rawCustomer !== null ? String(rawCustomer).split('.')[0].trim() : '';
        const rawGL = getRowValue(row, 'G/L Account') || getRowValue(row, 'G/L Account_1');
        const cleanGL = rawGL !== undefined && rawGL !== null ? String(rawGL).split('.')[0].trim() : '';

        const accountType = String(getRowValue(row, 'Account Type') || '').trim().toUpperCase();
        let ledgerName = '';
        let isParty = false;

        if (accountType === 'K' || (cleanVendor && cleanVendor !== '0' && !cleanGL)) {
            // Vendor Entry: 10-digit padded vendor code only
            const paddedVendor = padVendor(cleanVendor);
            ledgerName = paddedVendor;
            isParty = true;
        } else if (accountType === 'D' || (cleanCustomer && cleanCustomer !== '0' && !cleanGL)) {
            // Customer Entry: 10-digit padded customer code only
            const paddedCustomer = padVendor(cleanCustomer);
            ledgerName = paddedCustomer;
            isParty = true;
        } else if (cleanGL) {
            // G/L Account Entry: raw G/L account code only (no suffix)
            ledgerName = cleanGL;
            isParty = false;
        } else {
            ledgerName = 'Suspense Ledger';
            isParty = false;
        }

        if (!partyLedger && isParty) {
            partyLedger = ledgerName;
        }

        const amtVal = parseFloat(getRowValue(row, 'Amount') || getRowValue(row, 'Amount in LC') || 0) || 0;
        // In Tally XML: Debit is negative (-), Credit is positive (+) in AMOUNT tag when ISDEEMEDPOSITIVE matches
        const tallyAmountFormatted = isCredit ? amtVal.toFixed(2) : (-amtVal).toFixed(2);

        if (isParty && isCredit) {
            partyAmount = amtVal;
        }

        ledgerEntries.push({
            ledgerName,
            isDeemedPositive,
            isPartyLedger: isParty ? 'Yes' : 'No',
            amountFormatted: tallyAmountFormatted
        });
    });

    if (!partyLedger && ledgerEntries.length > 0) {
        partyLedger = ledgerEntries[0].ledgerName;
    }

    const ledgerEntriesXML = ledgerEntries.map(entry => `       <ALLLEDGERENTRIES.LIST>
        <OLDAUDITENTRYIDS.LIST TYPE="Number">
         <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
        </OLDAUDITENTRYIDS.LIST>
        <LEDGERNAME>${escapeXML(entry.ledgerName)}</LEDGERNAME>
        <GSTCLASS>&#4; Not Applicable</GSTCLASS>
        <ISDEEMEDPOSITIVE>${entry.isDeemedPositive}</ISDEEMEDPOSITIVE>
        <LEDGERFROMITEM>No</LEDGERFROMITEM>
        <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
        <ISPARTYLEDGER>${entry.isPartyLedger}</ISPARTYLEDGER>
        <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
        <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
        <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
        <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
        <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
        <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
        <ISLASTDEEMEDPOSITIVE>${entry.isDeemedPositive}</ISLASTDEEMEDPOSITIVE>
        <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
        <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
        <AMOUNT>${entry.amountFormatted}</AMOUNT>
        <VATEXPAMOUNT>${entry.amountFormatted}</VATEXPAMOUNT>
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
       </ALLLEDGERENTRIES.LIST>`).join('\n');

    const rawRefKey = String(getRowValue(firstRow, 'Reference Key') || getRowValue(firstRow, 'ReferenceKey') || '').trim();
    let narrationVal = rawRefKey;
    if (rawRefKey.length > 4) {
        narrationVal = rawRefKey.substring(0, rawRefKey.length - 4);
    }
    const narrationXML = narrationVal ? `\n      <NARRATION>${escapeXML(narrationVal)}</NARRATION>` : '';

    let purchasingDocVal = '';
    for (const r of items) {
        const pDoc = getRowValue(r, 'Purchasing Document') || getRowValue(r, 'PurchasingDoc') || getRowValue(r, 'Purchase Order');
        if (pDoc !== undefined && pDoc !== null && String(pDoc).trim() !== '' && String(pDoc).trim() !== '0') {
            purchasingDocVal = String(pDoc).split('.')[0].trim();
            break;
        }
    }
    const referenceXML = purchasingDocVal ? `\n      <REFERENCE>${escapeXML(purchasingDocVal)}</REFERENCE>` : '';

    return `<ENVELOPE>
 <HEADER>
  <TALLYREQUEST>Import Data</TALLYREQUEST>
 </HEADER>
 <BODY>
  <IMPORTDATA>
   <REQUESTDESC>
    <REPORTNAME>Vouchers</REPORTNAME>
    <STATICVARIABLES>
     <SVCURRENTCOMPANY>${COMPANY_NAME}</SVCURRENTCOMPANY>
    </STATICVARIABLES>
   </REQUESTDESC>
   <REQUESTDATA>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <VOUCHER VCHTYPE="Journal" ACTION="Create" OBJVIEW="Accounting Voucher View">
      <OLDAUDITENTRYIDS.LIST TYPE="Number">
       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
      </OLDAUDITENTRYIDS.LIST>
      <DATE>${dateFormatted}</DATE>
      <VCHSTATUSDATE>${dateFormatted}</VCHSTATUSDATE>
      ${narrationXML}
      <OBJECTUPDATEACTION/>
      <GSTREGISTRATION TAXTYPE="GST" TAXREGISTRATION="">Tamil Nadu Registration</GSTREGISTRATION>
      <VOUCHERTYPENAME>${voucherTypeName}</VOUCHERTYPENAME>
      <PARTYLEDGERNAME>${escapeXML(partyLedger)}</PARTYLEDGERNAME>
      <VOUCHERNUMBER>${escapeXML(docNumber)}</VOUCHERNUMBER>${referenceXML}
      <NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
      <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>
      <VCHSTATUSVOUCHERTYPE>Journal</VCHSTATUSVOUCHERTYPE>
      <VCHENTRYMODE>As Voucher</VCHENTRYMODE>
      <EFFECTIVEDATE>${dateFormatted}</EFFECTIVEDATE>
${ledgerEntriesXML}
     </VOUCHER>
    </TALLYMESSAGE>
   </REQUESTDATA>
  </IMPORTDATA>
 </BODY>
</ENVELOPE>`;
}

/**
 * Generate Tally XML for Delivery Note (WL Trans./Event Type)
 * Voucher Type Name: "Delivery Note WL"
 */
function generateDeliveryNoteTallyXML(deliveryGroup, vendorMap = {}) {
    if (!deliveryGroup || !deliveryGroup.items || deliveryGroup.items.length === 0) {
        throw new Error('No items in Delivery Note group');
    }

    const items = deliveryGroup.items;
    const firstRow = items[0];

    const rawDocNum = deliveryGroup.poNumber || getRowValue(firstRow, 'Material Document') || '';
    const voucherNumber = escapeXML(String(rawDocNum).split('.')[0].trim());

    const voucherTypeName = 'Delivery Note WL';

    const rawPostingDate = getRowValue(firstRow, 'Posting Date') || getRowValue(firstRow, 'Document Date');
    const docDateFormatted = formatDate(rawPostingDate);

    // Customer Party ledger lookup - use "Goods recipient" column directly
    const rawGoodsRecipient = getRowValue(firstRow, 'Goods recipient') || getRowValue(firstRow, 'Customer') || '';
    const customerCode = padVendor(rawGoodsRecipient);
    const partyLedger = escapeXML(customerCode || 'Unknown Customer');
    const vendors = typeof loadVendorMaster === 'function' ? loadVendorMaster() : {};

    const street = String(vendors[customerCode]?.street || getRowValue(firstRow, 'Street') || '').trim();
    const city = String(vendors[customerCode]?.city || getRowValue(firstRow, 'City') || '').trim();
    const address = escapeXML(street && city ? `${street},,,${city}` : (street || city || ''));

    const postCode = escapeXML(vendors[customerCode]?.postCode || String(getRowValue(firstRow, 'Post Code') || '').trim());
    const gstNo = escapeXML(vendors[customerCode]?.gstNo || String(getRowValue(firstRow, 'GST NO') || '').trim());
    const regionName = escapeXML(cleanStateName(vendors[customerCode]?.regionName || getRowValue(firstRow, 'Region Name') || 'Tamil Nadu'));
    const cmpState = 'Tamil Nadu';

    let totalVoucherAmount = 0;

    const inventoryEntriesXML = items.map(item => {
        const rawMaterial = String(getRowValue(item, 'Material') || '').trim();
        const stockItemName = rawMaterial || (String(getRowValue(item, 'Material Desc') || '').trim() || 'FLYASH');

        const plant = escapeXML(String(getRowValue(item, 'Plant') || '1000').split('.')[0].trim());

        const qtyNum = parseFloat(getRowValue(item, 'Qty in Un. of Entry') || getRowValue(item, 'Quantity') || 0) || 0;
        const uom = escapeXML(String(getRowValue(item, 'Unit of Entry') || getRowValue(item, 'Base Unit of Measure') || 'MT').trim());
        const qtyFormatted = ` ${formatQuantity(qtyNum)} ${uom}`;

        // Amount in LC is the line amount
        const lineAmt = parseFloat(getRowValue(item, 'Amount in LC') || getRowValue(item, 'Amount') || 0) || 0;
        totalVoucherAmount += lineAmt;

        const rate = qtyNum > 0 ? (lineAmt / qtyNum) : 0;
        const rateFormatted = `${rate.toFixed(2)}/${uom}`;
        const amountFormatted = lineAmt.toFixed(2);

        // Sales / Stock Ledger Name: e.g. "Flyash Sales" if FLYASH, or "Coal Sales" if COAL, or generic Sales
        let salesLedger = 'Flyash Sales';
        if (rawMaterial.toUpperCase().includes('COAL')) {
            salesLedger = 'Coal Sales';
        }

        return `       <ALLINVENTORYENTRIES.LIST>
        <STOCKITEMNAME>${escapeXML(stockItemName)}</STOCKITEMNAME>
        <GSTOVRDNISREVCHARGEAPPL>&#4; Not Applicable</GSTOVRDNISREVCHARGEAPPL>
        <GSTOVRDNSTOREDNATURE/>
        <GSTRATEINFERAPPLICABILITY>As per Masters/Company</GSTRATEINFERAPPLICABILITY>
        <GSTHSNINFERAPPLICABILITY>As per Masters/Company</GSTHSNINFERAPPLICABILITY>
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
         <GODOWNNAME>${plant}</GODOWNNAME>
         <BATCHNAME>Primary Batch</BATCHNAME>
         <DESTINATIONGODOWNNAME>${plant}</DESTINATIONGODOWNNAME>
         <INDENTNO>&#4; Not Applicable</INDENTNO>
         <ORDERNO>&#4; Not Applicable</ORDERNO>
         <TRACKINGNUMBER>${voucherNumber}</TRACKINGNUMBER>
         <DYNAMICCSTISCLEARED>No</DYNAMICCSTISCLEARED>
         <AMOUNT>${amountFormatted}</AMOUNT>
         <ACTUALQTY>${qtyFormatted}</ACTUALQTY>
         <BILLEDQTY>${qtyFormatted}</BILLEDQTY>
         <ADDITIONALDETAILS.LIST>        </ADDITIONALDETAILS.LIST>
         <VOUCHERCOMPONENTLIST.LIST>        </VOUCHERCOMPONENTLIST.LIST>
        </BATCHALLOCATIONS.LIST>
        <ACCOUNTINGALLOCATIONS.LIST>
         <OLDAUDITENTRYIDS.LIST TYPE="Number">
          <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
         </OLDAUDITENTRYIDS.LIST>
         <LEDGERNAME>${escapeXML(salesLedger)}</LEDGERNAME>
         <GSTCLASS>&#4; Not Applicable</GSTCLASS>
         <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
         <LEDGERFROMITEM>No</LEDGERFROMITEM>
         <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
         <ISPARTYLEDGER>No</ISPARTYLEDGER>
         <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
         <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
         <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
         <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
         <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
         <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
         <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
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

    const totalPartyAmountFormatted = (-totalVoucherAmount).toFixed(2);

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
     <VOUCHER VCHTYPE="${escapeXML(voucherTypeName)}" ACTION="Create" OBJVIEW="Invoice Voucher View">
      <ADDRESS.LIST TYPE="String">
       <ADDRESS>${address}</ADDRESS>
      </ADDRESS.LIST>
      <BASICBUYERADDRESS.LIST TYPE="String">
       <BASICBUYERADDRESS>${address}</BASICBUYERADDRESS>
      </BASICBUYERADDRESS.LIST>
      <OLDAUDITENTRYIDS.LIST TYPE="Number">
       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
      </OLDAUDITENTRYIDS.LIST>
      <DATE>${docDateFormatted}</DATE>
      <VCHSTATUSDATE>${docDateFormatted}</VCHSTATUSDATE>
      <GSTREGISTRATIONTYPE>&#4; Unknown</GSTREGISTRATIONTYPE>
      <VATDEALERTYPE>&#4; Unknown</VATDEALERTYPE>
      <STATENAME>${regionName}</STATENAME>
      <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
      <PARTYGSTIN>${gstNo}</PARTYGSTIN>
      <PLACEOFSUPPLY>${regionName}</PLACEOFSUPPLY>
      <VOUCHERTYPENAME>${escapeXML(voucherTypeName)}</VOUCHERTYPENAME>
      <PARTYNAME>${partyLedger}</PARTYNAME>
      <GSTREGISTRATION TAXTYPE="GST" TAXREGISTRATION="">Tamil Nadu Registration</GSTREGISTRATION>
      <PARTYLEDGERNAME>${partyLedger}</PARTYLEDGERNAME>
      <VOUCHERNUMBER>${voucherNumber}</VOUCHERNUMBER>
      <BASICBUYERNAME>${partyLedger}</BASICBUYERNAME>
      <CMPGSTREGISTRATIONTYPE>Regular</CMPGSTREGISTRATIONTYPE>
      <CMPGSTSTATE>${cmpState}</CMPGSTSTATE>
      <NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
      <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
      <VCHSTATUSVOUCHERTYPE>${escapeXML(voucherTypeName)}</VCHSTATUSVOUCHERTYPE>
      <VCHSTATUSTAXUNIT>Tamil Nadu Registration</VCHSTATUSTAXUNIT>
      <VCHGSTCLASS>&#4; Not Applicable</VCHGSTCLASS>
      <DIFFACTUALQTY>No</DIFFACTUALQTY>
      <ISMSTFROMSYNC>No</ISMSTFROMSYNC>
      <ISDELETED>No</ISDELETED>
      <ASORIGINAL>No</ASORIGINAL>
      <AUDITED>No</AUDITED>
      <ISOPTIONAL>No</ISOPTIONAL>
      <EFFECTIVEDATE>${docDateFormatted}</EFFECTIVEDATE>
      <USETRACKINGNUMBER>No</USETRACKINGNUMBER>
      <ISINVOICE>No</ISINVOICE>
      <ISVATDUTYPAID>Yes</ISVATDUTYPAID>
${inventoryEntriesXML}
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
       <LEDGERNAME>${partyLedger}</LEDGERNAME>
       <GSTCLASS>&#4; Not Applicable</GSTCLASS>
       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
       <LEDGERFROMITEM>No</LEDGERFROMITEM>
       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
       <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
       <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
       <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
       <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
       <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
       <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
       <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
       <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>
       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
       <AMOUNT>${totalPartyAmountFormatted}</AMOUNT>
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
   </REQUESTDATA>
  </IMPORTDATA>
 </BODY>
</ENVELOPE>`;
}

/**
 * Generate Tally XML for Sales Invoice
 * Matches Sales_90015368.xml specification
 * Narration includes computed non-empty values for:
 * Profit centre, Volt  -  Draw Voltage, Grp2, Grp2-Desc, Grp3, Grp3  Desc, Std Start Date, Std End Date
 * Tax amount is split into CGST + SGST if Destination region is '33' (local), else IGST
 */
function generateSalesInvoiceTallyXML(salesGroup, vendorMap = {}) {
    if (!salesGroup || !salesGroup.items || salesGroup.items.length === 0) {
        throw new Error('No items in Sales Invoice group');
    }

    const items = salesGroup.items;
    const firstRow = items[0];

    const rawBillingDoc = salesGroup.poNumber || getRowValue(firstRow, 'Billing Document') || getRowValue(firstRow, 'Document Number') || '';
    const voucherNumber = escapeXML(String(rawBillingDoc).split('.')[0].trim());

    // Voucher Type Name: Sales {Billing Type} (e.g. Sales ZCOL)
    const billingType = String(getRowValue(firstRow, 'Billing Type') || 'ZCOL').trim();
    const voucherType = escapeXML(`Sales ${billingType}`);

    // Sales document reference
    const rawSalesDoc = getRowValue(firstRow, 'Sales document') || getRowValue(firstRow, 'Reference document') || voucherNumber;
    const referenceNumber = escapeXML(String(rawSalesDoc).split('.')[0].trim());

    // Billing Date
    const rawBillingDate = getRowValue(firstRow, 'Billing Date') || getRowValue(firstRow, 'Pricing date') || getRowValue(firstRow, 'Document Date');
    const docDateFormatted = formatDate(rawBillingDate);

    // Customer Party ledger lookup
    const rawSoldToParty = getRowValue(firstRow, 'Sold to Party') || getRowValue(firstRow, 'Customer') || '';
    const customerCode = padVendor(rawSoldToParty);
    const partyLedger = escapeXML(customerCode || 'Customer');

    const gstNo = escapeXML(String(getRowValue(firstRow, 'GST NO') || '').trim());
    const destRegion = String(getRowValue(firstRow, 'Destination region') || '33').trim();
    const isLocalTamilNadu = destRegion === '33' || destRegion === '';
    const stateName = isLocalTamilNadu ? 'Tamil Nadu' : cleanStateName(getRowValue(firstRow, 'Region Name') || 'Other State');
    const cmpState = 'Tamil Nadu';

/**
 * Format date for narration (DD.MM.YYYY).
 * Handles numeric Excel serial dates (e.g. 44722 -> 10.06.2022).
 * Returns null for zero, empty, or 00.01.1900 dummy dates so they are omitted from narration.
 */
function formatNarrationDate(val) {
    if (val === undefined || val === null) return null;
    const str = String(val).trim();
    if (str === '' || str === '0' || str.toUpperCase() === '#N/A' || str === '00.01.1900' || str === '00.00.0000' || str === '00-01-1900' || str === '1900-01-00') {
        return null;
    }

    // Excel numeric serial date (e.g. 44722 or '44722')
    if (typeof val === 'number' || (!isNaN(str) && !str.includes('.') && !str.includes('-') && !str.includes('/'))) {
        const serial = parseFloat(str);
        if (serial <= 60) return null; // 0 or Excel 1900 dummy date
        const wholeDays = Math.floor(serial);
        const date = new Date(Date.UTC(1899, 11, 30 + wholeDays));
        if (isNaN(date.getTime()) || date.getUTCFullYear() <= 1900) return null;
        const d = String(date.getUTCDate()).padStart(2, '0');
        const m = String(date.getUTCMonth() + 1).padStart(2, '0');
        const y = date.getUTCFullYear();
        return `${d}.${m}.${y}`;
    }

    // Date object
    if (val instanceof Date) {
        if (isNaN(val.getTime()) || val.getFullYear() <= 1900) return null;
        const d = String(val.getDate()).padStart(2, '0');
        const m = String(val.getMonth() + 1).padStart(2, '0');
        const y = val.getFullYear();
        return `${d}.${m}.${y}`;
    }

    // DD.MM.YYYY string
    if (str.includes('.')) {
        const parts = str.split('.');
        if (parts.length === 3) {
            const d = parts[0].padStart(2, '0');
            const m = parts[1].padStart(2, '0');
            const y = parts[2].trim();
            if (y === '1900' || y === '0000' || (d === '00' && m === '01') || (d === '00' && m === '00')) return null;
            return `${d}.${m}.${y}`;
        }
    }

    // YYYY-MM-DD or DD-MM-YYYY or with slashes
    if (str.includes('-') || str.includes('/')) {
        const sep = str.includes('-') ? '-' : '/';
        const parts = str.split(sep);
        if (parts.length === 3) {
            if (parts[0].length === 4) {
                // YYYY-MM-DD
                const y = parts[0];
                const m = parts[1].padStart(2, '0');
                const d = parts[2].padStart(2, '0');
                if (y === '1900' || y === '0000' || (d === '00' && m === '01') || (d === '00' && m === '00')) return null;
                return `${d}.${m}.${y}`;
            } else {
                // DD-MM-YYYY
                const d = parts[0].padStart(2, '0');
                const m = parts[1].padStart(2, '0');
                const y = parts[2];
                if (y === '1900' || y === '0000' || (d === '00' && m === '01') || (d === '00' && m === '00')) return null;
                return `${d}.${m}.${y}`;
            }
        }
    }

    return str;
}

    // Narration computed fields
    const narrationFields = [
        { col: 'Profit centre', label: 'Profit centre' },
        { col: 'Volt  -  Draw Voltage', label: 'Volt  -  Draw Voltage' },
        { col: 'Grp2', label: 'Grp2' },
        { col: 'Grp2-Desc', label: 'Grp2-Desc' },
        { col: 'Grp3', label: 'Grp3' },
        { col: 'Grp3  Desc', label: 'Grp3  Desc' },
        { col: 'Std Start Date', label: 'Std Start Date', isDate: true },
        { col: 'Std End Date', label: 'Std End Date', isDate: true },
        { col: 'Billing Date', label: 'Fixed Value Date', isDate: true, fallbackCol: 'Fixed Value Date' }
    ];

    const narrationParts = [];
    narrationFields.forEach(field => {
        const col = field.col;
        const label = field.label;
        let val = getRowValue(firstRow, col);
        if ((val === undefined || val === null || String(val).trim() === '' || String(val).trim() === '0') && field.fallbackCol) {
            val = getRowValue(firstRow, field.fallbackCol);
        }
        if ((val === undefined || val === null || String(val).trim() === '' || String(val).trim() === '0') && items.length > 1) {
            for (let i = 1; i < items.length; i++) {
                const itemVal = getRowValue(items[i], col) || (field.fallbackCol ? getRowValue(items[i], field.fallbackCol) : undefined);
                if (itemVal !== undefined && itemVal !== null && String(itemVal).trim() !== '' && String(itemVal).trim() !== '0') {
                    val = itemVal;
                    break;
                }
            }
        }

        if (field.isDate) {
            const formattedDate = formatNarrationDate(val);
            if (formattedDate) {
                narrationParts.push(`${label} : ${formattedDate}`);
            }
        } else {
            if (val !== undefined && val !== null && String(val).trim() !== '' && String(val).trim() !== '0' && String(val).trim().toUpperCase() !== '#N/A') {
                narrationParts.push(`${label} : ${String(val).trim()}`);
            }
        }
    });
    const narrationText = escapeXML(narrationParts.join(' | '));

    let totalNetValue = 0;
    let totalTaxAmount = 0;

    const inventoryEntriesXML = items.map(item => {
        const rawMaterial = String(getRowValue(item, 'Material') || getRowValue(item, 'Material entered') || '').trim();
        const shortText = String(getRowValue(item, 'Description') || getRowValue(item, 'Material Description') || getRowValue(item, 'Short Text') || '').trim();
        const stockItemName = escapeXML(getStockItemName(rawMaterial, shortText));

        const plant = escapeXML(String(getRowValue(item, 'Plant') || '1310').split('.')[0].trim());
        const batch = escapeXML(String(getRowValue(item, 'Batch') || 'Primary Batch').trim() || 'Primary Batch');

        const qtyNum = parseFloat(getRowValue(item, 'Billed Quantity') || getRowValue(item, 'Quantity') || 0) || 0;
        const uom = escapeXML(String(getRowValue(item, 'Sales unit') || getRowValue(item, 'Base Unit of Measure') || 'MT').trim());
        const qtyFormatted = ` ${formatQuantity(qtyNum)} ${uom}`;

        const rowNetVal = parseFloat(getRowValue(item, 'Net value') || 0) || 0;
        const rowTaxAmt = parseFloat(getRowValue(item, 'Tax amount') || getRowValue(item, 'Tax amount_1') || 0) || 0;
        totalNetValue += rowNetVal;
        totalTaxAmount += rowTaxAmt;

        // Assessable Inventory Amount = Net value - Tax amount
        let assessableAmount = rowNetVal - rowTaxAmt;
        if (assessableAmount <= 0 && rowNetVal > 0) {
            assessableAmount = rowNetVal;
        }

        const rate = qtyNum > 0 ? (assessableAmount / qtyNum) : 0;
        const rateFormatted = `${rate.toFixed(2)}/${uom}`;
        const amountFormatted = assessableAmount.toFixed(2);

        // Sales Ledger: e.g. "Sales ZCOL" from Billing Type
        const billingType = String(getRowValue(item, 'Billing Type') || getRowValue(firstRow, 'Billing Type') || 'ZCOL').trim();
        const salesLedger = `Sales ${billingType}`;

        return `       <ALLINVENTORYENTRIES.LIST>
        <STOCKITEMNAME>${stockItemName}</STOCKITEMNAME>
        <GSTOVRDNINELIGIBLEITC>&#4; Not Applicable</GSTOVRDNINELIGIBLEITC>
        <GSTOVRDNISREVCHARGEAPPL>&#4; Not Applicable</GSTOVRDNISREVCHARGEAPPL>
        <GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
        <GSTSOURCETYPE>Stock Item</GSTSOURCETYPE>
        <GSTITEMSOURCE>${stockItemName}</GSTITEMSOURCE>
        <HSNSOURCETYPE>Stock Item</HSNSOURCETYPE>
        <HSNITEMSOURCE>${stockItemName}</HSNITEMSOURCE>
        <GSTOVRDNSTOREDNATURE/>
        <GSTOVRDNTYPEOFSUPPLY>Capital Goods</GSTOVRDNTYPEOFSUPPLY>
        <GSTRATEINFERAPPLICABILITY>As per Masters/Company</GSTRATEINFERAPPLICABILITY>
        <GSTHSNINFERAPPLICABILITY>As per Masters/Company</GSTHSNINFERAPPLICABILITY>
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
         <GODOWNNAME>${plant}</GODOWNNAME>
         <BATCHNAME>${batch}</BATCHNAME>
         <DESTINATIONGODOWNNAME>${plant}</DESTINATIONGODOWNNAME>
         <INDENTNO>&#4; Not Applicable</INDENTNO>
         <ORDERNO>${referenceNumber}</ORDERNO>
         <TRACKINGNUMBER>${voucherNumber}</TRACKINGNUMBER>
         <DYNAMICCSTISCLEARED>No</DYNAMICCSTISCLEARED>
         <AMOUNT>${amountFormatted}</AMOUNT>
         <ACTUALQTY>${qtyFormatted}</ACTUALQTY>
         <BILLEDQTY>${qtyFormatted}</BILLEDQTY>
         <ADDITIONALDETAILS.LIST>        </ADDITIONALDETAILS.LIST>
         <VOUCHERCOMPONENTLIST.LIST>        </VOUCHERCOMPONENTLIST.LIST>
        </BATCHALLOCATIONS.LIST>
        <ACCOUNTINGALLOCATIONS.LIST>
         <OLDAUDITENTRYIDS.LIST TYPE="Number">
          <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
         </OLDAUDITENTRYIDS.LIST>
         <LEDGERNAME>${escapeXML(salesLedger)}</LEDGERNAME>
         <GSTCLASS>&#4; Not Applicable</GSTCLASS>
         <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
         <LEDGERFROMITEM>No</LEDGERFROMITEM>
         <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
         <ISPARTYLEDGER>No</ISPARTYLEDGER>
         <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
         <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
         <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
         <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
         <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
         <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
         <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
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
        <SUPPLEMENTARYDUTYHEADDETAILS.LIST>       </SUPPLEMENTARYDUTYHEADDETAILS.LIST>
        <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>
        <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>
        <EXCISEALLOCATIONS.LIST>       </EXCISEALLOCATIONS.LIST>
        <EXPENSEALLOCATIONS.LIST>       </EXPENSEALLOCATIONS.LIST>
       </ALLINVENTORYENTRIES.LIST>`;
    }).join('\n');

    // Tax calculation based on Destination region
    let taxLedgersXML = '';
    if (totalTaxAmount > 0) {
        if (isLocalTamilNadu) {
            // Local 33: Split tax amount equally for CGST and SGST
            const halfTax = totalTaxAmount / 2;
            const halfTaxFormatted = halfTax.toFixed(2);
            taxLedgersXML = `      <LEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <ROUNDTYPE>&#4; Not Applicable</ROUNDTYPE>
       <LEDGERNAME>CGST Tax Sales</LEDGERNAME>
       <GSTCLASS>&#4; Not Applicable</GSTCLASS>
       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
       <LEDGERFROMITEM>No</LEDGERFROMITEM>
       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
       <ISPARTYLEDGER>No</ISPARTYLEDGER>
       <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
       <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
       <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
       <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
       <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
       <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
       <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
       <AMOUNT>${halfTaxFormatted}</AMOUNT>
       <VATEXPAMOUNT>${halfTaxFormatted}</VATEXPAMOUNT>
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
      <LEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <ROUNDTYPE>&#4; Not Applicable</ROUNDTYPE>
       <LEDGERNAME>SGST Tax Sales</LEDGERNAME>
       <GSTCLASS>&#4; Not Applicable</GSTCLASS>
       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
       <LEDGERFROMITEM>No</LEDGERFROMITEM>
       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
       <ISPARTYLEDGER>No</ISPARTYLEDGER>
       <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
       <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
       <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
       <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
       <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
       <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
       <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
       <AMOUNT>${halfTaxFormatted}</AMOUNT>
       <VATEXPAMOUNT>${halfTaxFormatted}</VATEXPAMOUNT>
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
        } else {
            // Other destination region: Pass full tax amount as IGST Tax Sales
            const taxFormatted = totalTaxAmount.toFixed(2);
            taxLedgersXML = `      <LEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <ROUNDTYPE>&#4; Not Applicable</ROUNDTYPE>
       <LEDGERNAME>IGST Tax Sales</LEDGERNAME>
       <GSTCLASS>&#4; Not Applicable</GSTCLASS>
       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
       <LEDGERFROMITEM>No</LEDGERFROMITEM>
       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
       <ISPARTYLEDGER>No</ISPARTYLEDGER>
       <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
       <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
       <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
       <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
       <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
       <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
       <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
       <AMOUNT>${taxFormatted}</AMOUNT>
       <VATEXPAMOUNT>${taxFormatted}</VATEXPAMOUNT>
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
    }

    // Party entry amount is total net value (negative in Tally credit entry)
    const partyAmountFormatted = (-totalNetValue).toFixed(2);
    const remoteId = `1f4bca00-f679-4031-a055-5ca2bb211a68-${voucherNumber.padStart(8, '0')}`;
    const vchKey = `1f4bca00-f679-4031-a055-5ca2bb211a68-0000ae6b:${voucherNumber.padStart(8, '0')}`;

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
      <OLDAUDITENTRYIDS.LIST TYPE="Number">
       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
      </OLDAUDITENTRYIDS.LIST>
      <DATE>${docDateFormatted}</DATE>
      <REFERENCEDATE>${docDateFormatted}</REFERENCEDATE>
      <VCHSTATUSDATE>${docDateFormatted}</VCHSTATUSDATE>
      <GUID>${remoteId}</GUID>
      <GSTREGISTRATIONTYPE>&#4; Unknown</GSTREGISTRATIONTYPE>
      <VATDEALERTYPE>&#4; Unknown</VATDEALERTYPE>
      <STATENAME>${stateName}</STATENAME>
      <NARRATION>${narrationText}</NARRATION>
      <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
      <PLACEOFSUPPLY>${stateName}</PLACEOFSUPPLY>
      <VOUCHERTYPENAME>${voucherType}</VOUCHERTYPENAME>
      <PARTYNAME>${partyLedger}</PARTYNAME>
      <PARTYLEDGERNAME>${partyLedger}</PARTYLEDGERNAME>
      <VOUCHERNUMBER>${voucherNumber}</VOUCHERNUMBER>
      <BASICBUYERNAME>${partyLedger}</BASICBUYERNAME>
      <CMPGSTREGISTRATIONTYPE>Regular</CMPGSTREGISTRATIONTYPE>
      <REFERENCE>${referenceNumber}</REFERENCE>
      <PARTYMAILINGNAME>${partyLedger}</PARTYMAILINGNAME>
      <CONSIGNEEMAILINGNAME>${partyLedger}</CONSIGNEEMAILINGNAME>
      <CONSIGNEESTATENAME>${stateName}</CONSIGNEESTATENAME>
      <CMPGSTSTATE>${cmpState}</CMPGSTSTATE>
      <CONSIGNEECOUNTRYNAME>India</CONSIGNEECOUNTRYNAME>
      <BASICBASEPARTYNAME>${partyLedger}</BASICBASEPARTYNAME>
      <NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
      <CSTFORMISSUETYPE>&#4; Not Applicable</CSTFORMISSUETYPE>
      <CSTFORMRECVTYPE>&#4; Not Applicable</CSTFORMRECVTYPE>
      <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>
      <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
      <VCHSTATUSTAXADJUSTMENT>Default</VCHSTATUSTAXADJUSTMENT>
      <VCHSTATUSVOUCHERTYPE>${voucherType}</VCHSTATUSVOUCHERTYPE>
      <VCHGSTCLASS>&#4; Not Applicable</VCHGSTCLASS>
      <VCHENTRYMODE>Item Invoice</VCHENTRYMODE>
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
      <ISELIGIBLEFORITC>Yes</ISELIGIBLEFORITC>
      <USETRACKINGNUMBER>No</USETRACKINGNUMBER>
      <ISINVOICE>Yes</ISINVOICE>
      <ISVATDUTYPAID>Yes</ISVATDUTYPAID>
${inventoryEntriesXML}
      <CONTRITRANS.LIST>      </CONTRITRANS.LIST>
      <EWAYBILLERRORLIST.LIST>      </EWAYBILLERRORLIST.LIST>
      <IRNERRORLIST.LIST>      </IRNERRORLIST.LIST>
      <HARYANAVAT.LIST>      </HARYANAVAT.LIST>
      <SUPPLEMENTARYDUTYHEADDETAILS.LIST>      </SUPPLEMENTARYDUTYHEADDETAILS.LIST>
      <INVOICEDELNOTES.LIST>      </INVOICEDELNOTES.LIST>
      <INVOICEORDERLIST.LIST>
       <BASICORDERDATE>${docDateFormatted}</BASICORDERDATE>
       <BASICPURCHASEORDERNO>${referenceNumber}</BASICPURCHASEORDERNO>
      </INVOICEORDERLIST.LIST>
      <INVOICEINDENTLIST.LIST>      </INVOICEINDENTLIST.LIST>
      <ATTENDANCEENTRIES.LIST>      </ATTENDANCEENTRIES.LIST>
      <ORIGINVOICEDETAILS.LIST>      </ORIGINVOICEDETAILS.LIST>
      <INVOICEEXPORTLIST.LIST>      </INVOICEEXPORTLIST.LIST>
      <LEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <LEDGERNAME>${partyLedger}</LEDGERNAME>
       <GSTCLASS>&#4; Not Applicable</GSTCLASS>
       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
       <LEDGERFROMITEM>No</LEDGERFROMITEM>
       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
       <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
       <GSTOVERRIDDEN>No</GSTOVERRIDDEN>
       <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
       <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
       <STRDGSTISPARTYLEDGER>No</STRDGSTISPARTYLEDGER>
       <STRDGSTISDUTYLEDGER>No</STRDGSTISDUTYLEDGER>
       <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
       <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>
       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
       <AMOUNT>${partyAmountFormatted}</AMOUNT>
       <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>
       <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>
       <BILLALLOCATIONS.LIST>
        <NAME>${voucherNumber}</NAME>
        <BILLTYPE>New Ref</BILLTYPE>
        <TDSDEDUCTEEISSPECIALRATE>No</TDSDEDUCTEEISSPECIALRATE>
        <AMOUNT>${partyAmountFormatted}</AMOUNT>
        <INTERESTCOLLECTION.LIST>        </INTERESTCOLLECTION.LIST>
        <STBILLCATEGORIES.LIST>        </STBILLCATEGORIES.LIST>
       </BILLALLOCATIONS.LIST>
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
${taxLedgersXML}
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
       <NAME>1f4bca00-f679-4031-a055-5ca2bb211a68</NAME>
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
    generateSalesOrderTallyXML,
    generateFITallyXML,
    generateDeliveryNoteTallyXML,
    generateSalesInvoiceTallyXML,
    loadPOMaster
};

