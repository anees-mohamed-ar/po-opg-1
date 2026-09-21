const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const axios = require('axios');
const path = require('path');
const fs = require('fs');
const { generateTallyXML, generateGRNTallyXML, generatePurchaseTallyXML, generateSalesOrderTallyXML, generateFITallyXML, getRowValue, padVendor } = require('./tallyXMLBuilder');
const config = require('./config');

const app = express();
const PORT = config.SERVER_PORT;
const TALLY_URL = config.TALLY_URL;

app.use(cors());
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ limit: '100mb', extended: true }));

// Set up multer for file uploads
const upload = multer({ dest: 'uploads/' });

/**
 * Clean up uploaded file
 */
const cleanupFile = (filePath) => {
    if (filePath && fs.existsSync(filePath)) {
        try {
            fs.unlinkSync(filePath);
        } catch (err) {
            console.error('Error deleting temp file:', err);
        }
    }
};

/**
 * Route to check service health
 */
app.get('/health', (req, res) => {
    res.json({ status: 'ok', tallyUrl: TALLY_URL });
});

app.get('/api/po-xml', (req, res) => {
    try {
        const reqFile = req.query.file || 'Purchase Order_4600001048.xml';
        // Prevent path traversal
        const safeFile = path.basename(reqFile);
        const filePath = path.join(__dirname, '../', safeFile);
        if (!fs.existsSync(filePath)) {
            return res.status(404).json({ error: `XML file not found: ${safeFile}` });
        }
        let content = fs.readFileSync(filePath, 'utf16le');
        if (!content.trim().startsWith('<')) {
            content = fs.readFileSync(filePath, 'utf8');
        }
        // Sanitize invalid XML control character references like &#4;
        content = content.replace(/&#\d+;/g, '');
        res.type('application/xml').send(content);
    } catch (err) {
        res.status(500).json({ error: 'Failed to read XML file', details: err.message });
    }
});

app.get('/api/daybook-xml', (req, res) => {
    try {
        const filePath = path.join(__dirname, '../DayBook.xml');
        if (!fs.existsSync(filePath)) {
            return res.status(404).json({ error: 'DayBook XML file not found' });
        }
        let content = fs.readFileSync(filePath, 'utf16le');
        if (!content.trim().startsWith('<')) {
            content = fs.readFileSync(filePath, 'utf8');
        }
        // Sanitize invalid XML control character references like &#4;
        content = content.replace(/&#\d+;/g, '');
        res.type('application/xml').send(content);
    } catch (err) {
        res.status(500).json({ error: 'Failed to read DayBook XML file', details: err.message });
    }
});

let globalVendorMap = {};

/**
 * POST /api/upload
 * Takes an Excel sheet, parses the details, and returns grouped Purchase Orders.
 */
app.post('/api/upload', upload.any(), async (req, res) => {
    const mainFileObj = (req.files || []).find(f => f.fieldname === 'file') || (req.files && req.files[0]);
    if (!mainFileObj) {
        return res.status(400).json({ error: 'No file uploaded' });
    }

    const mappingFileObj = (req.files || []).find(f => f.fieldname === 'vendorMappingFile' || f.fieldname === 'vendorMapping' || f.fieldname === 'mappingFile');

    const filePath = mainFileObj.path;
    const mappingFilePath = mappingFileObj ? mappingFileObj.path : null;

    try {
        const importType = String(req.query.importType || req.body.importType || 'po').trim().toLowerCase();
        const isGRN = importType === 'grn';
        const isPurchase = importType === 'purchase';
        const isStockJournal = importType === 'stock_journal';
        const isSalesOrder = importType === 'sales_order';
        const isFI = importType === 'fi';

        // Parse optional vendor mapping file
        // Supports new format: PO NO, ItemNo, Condition type, Currency, Cond.exchange rate, Vendor, Condition value
        // vendorChargeMap: { [poNumber]: { [itemNumber]: { [condType]: { vendor, conditionValueINR } } } }
        const vendorChargeMap = {};
        if (mappingFilePath) {
            try {
                const mapWb = xlsx.readFile(mappingFilePath);
                cleanupFile(mappingFilePath);
                // Prefer the first sheet; 'IMATRIX-VENDOR -CONDITION TYPE' or 'Vendor with values' – both have same structure
                const mapSheetName = mapWb.SheetNames[0];
                const mapRows = xlsx.utils.sheet_to_json(mapWb.Sheets[mapSheetName]);
                mapRows.forEach(mr => {
                    // Support both old (PO NUMBER / PO Number / Purchasing Document) and new (PO NO) column names
                    const poNum = String(
                        getRowValue(mr, 'PO NO') || getRowValue(mr, 'PO NUMBER') ||
                        getRowValue(mr, 'PO Number') || getRowValue(mr, 'Purchasing Document') || ''
                    ).trim();
                    // Support both old (PO Line Item No / Line Item / Item) and new (ItemNo) column names
                    const lineItem = String(
                        getRowValue(mr, 'ItemNo') || getRowValue(mr, 'PO Line Item No') ||
                        getRowValue(mr, 'Line Item') || getRowValue(mr, 'Item') || ''
                    ).trim();
                    const condType = String(getRowValue(mr, 'Condition type') || getRowValue(mr, 'Condition Type') || '').trim().toUpperCase();
                    const vendor = String(getRowValue(mr, 'Vendor') || getRowValue(mr, 'Vendor Code') || '').trim();

                    // New: read condition value with currency conversion
                    const rawCondValue = parseFloat(
                        String(getRowValue(mr, 'Condition value') || getRowValue(mr, 'Condition Value') || '0').replace(/,/g, '')
                    ) || 0;
                    const currency = String(getRowValue(mr, 'Currency') || 'INR').trim().toUpperCase();
                    const rawExRate = parseFloat(
                        String(getRowValue(mr, 'Cond.exchange rate') || getRowValue(mr, 'Exchange Rate') || '1').replace(/,/g, '')
                    ) || 1;
                    // Convert to INR: if not INR, multiply by exchange rate
                    const conditionValueINR = (currency !== 'INR' && rawExRate > 0)
                        ? rawCondValue * rawExRate
                        : rawCondValue;

                    if (poNum && condType && vendor) {
                        const cleanPo = poNum.split('.')[0].trim();
                        const cleanItem = lineItem ? lineItem.split('.')[0].trim() : '';
                        if (!vendorChargeMap[cleanPo]) vendorChargeMap[cleanPo] = {};
                        const entry = { vendor, conditionValueINR };
                        if (cleanItem) {
                            if (!vendorChargeMap[cleanPo][cleanItem]) vendorChargeMap[cleanPo][cleanItem] = {};
                            vendorChargeMap[cleanPo][cleanItem][condType] = entry;
                        } else {
                            if (!vendorChargeMap[cleanPo]['*']) vendorChargeMap[cleanPo]['*'] = {};
                            vendorChargeMap[cleanPo]['*'][condType] = entry;
                        }
                    }
                });
                console.log(`Loaded vendor mapping file with entries for ${Object.keys(vendorChargeMap).length} POs.`);
            } catch (err) {
                console.error('Error reading vendor mapping file:', err);
            }
        }

        // Read the uploaded excel sheet
        const workbook = xlsx.readFile(filePath);

        // Find a sheet with 'detail', 'sheet1', 'sheet', or 'main' in the name
        let detailSheetName = null;
        for (const name of workbook.SheetNames) {
            const lowerName = name.toLowerCase();
            if (lowerName.includes('detail') || lowerName.includes('sheet1') || lowerName.includes('sheet') || lowerName.includes('main')) {
                detailSheetName = name;
                break;
            }
        }

        if (!detailSheetName && workbook.SheetNames.length > 0) {
            detailSheetName = workbook.SheetNames[0];
        }

        if (!detailSheetName) {
            cleanupFile(filePath);
            return res.status(400).json({
                error: 'Could not find a valid sheet in the workbook.',
                availableSheets: workbook.SheetNames
            });
        }

        const sheet = workbook.Sheets[detailSheetName];

        // Parse sheet to 2D array first to dynamically locate the header row
        const rawGrid = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        cleanupFile(filePath);

        // Find the index of the row containing the actual column headers
        let headerRowIndex = -1;
        const primaryHeaders = ['Purchasing Document', 'Document Number', 'Invoice No', 'Invoice Number', 'Material Document', 'Sales document'];
        const secondaryHeaders = ['Invoicing Party', 'Vendor'];

        // First pass: look for a row with primary document headers
        for (let i = 0; i < Math.min(rawGrid.length, 50); i++) {
            const row = rawGrid[i];
            if (row && Array.isArray(row)) {
                const hasPrimary = row.some(cell => primaryHeaders.includes(String(cell || '').trim()));
                if (hasPrimary) {
                    headerRowIndex = i;
                    break;
                }
            }
        }

        // Second pass fallback: look for secondary headers
        if (headerRowIndex === -1) {
            for (let i = 0; i < Math.min(rawGrid.length, 50); i++) {
                const row = rawGrid[i];
                if (row && Array.isArray(row)) {
                    const hasSecondary = row.some(cell => secondaryHeaders.includes(String(cell || '').trim()));
                    if (hasSecondary) {
                        headerRowIndex = i;
                        break;
                    }
                }
            }
        }

        if (headerRowIndex === -1) {
            headerRowIndex = 0; // Fallback
        }

        res.setHeader('Content-Type', 'application/json');
        res.setHeader('Transfer-Encoding', 'chunked');
        res.write(JSON.stringify({ type: 'progress', message: 'Starting row extraction...', rowsParsed: 0, totalRows: rawGrid.length }) + '\n');

        // Convert subsequent rows into objects using the detected headers
        const headerCount = {};
        const headers = rawGrid[headerRowIndex].map(h => {
            const clean = String(h || '').trim();
            if (!clean) return '';
            if (headerCount[clean] !== undefined) {
                headerCount[clean]++;
                return `${clean}_${headerCount[clean]}`;
            } else {
                headerCount[clean] = 0;
                return clean;
            }
        });
        const rawRows = [];
        let rowIndex = headerRowIndex + 1;
        const totalRows = rawGrid.length;

        function processBatch() {
            const batchSize = 2500;
            const endIdx = Math.min(rowIndex + batchSize, totalRows);
            for (; rowIndex < endIdx; rowIndex++) {
                const row = rawGrid[rowIndex];
                if (!row || row.length === 0) continue;

                const obj = {};
                headers.forEach((header, colIdx) => {
                    if (header) {
                        obj[header] = row[colIdx];
                    }
                });

                // Attach vendor mappings if available
                // _chargeVendors: { [condType]: vendorCode }  (string, for XML builder)
                // _chargeValues:  { [condType]: conditionValueINR } (number, INR-converted)
                if (Object.keys(vendorChargeMap).length > 0) {
                    const poNum = String(
                        getRowValue(obj, 'Purchasing Document') || getRowValue(obj, 'PO Number') ||
                        getRowValue(obj, 'Document Number') || ''
                    ).split('.')[0].trim();
                    const lineItem = String(getRowValue(obj, 'Item') || getRowValue(obj, 'Line Item') || '').split('.')[0].trim();
                    if (poNum && vendorChargeMap[poNum]) {
                        obj._chargeVendors = {};
                        obj._chargeValues = {};
                        // First apply wildcard (*) entries, then item-specific (overrides wildcard)
                        const applyEntries = (map) => {
                            for (const [condType, entry] of Object.entries(map)) {
                                if (typeof entry === 'object' && entry !== null && 'vendor' in entry) {
                                    // New format: { vendor, conditionValueINR }
                                    obj._chargeVendors[condType] = entry.vendor;
                                    // Add condition values (accumulate per item for multi-item POs)
                                    obj._chargeValues[condType] = (obj._chargeValues[condType] || 0) + entry.conditionValueINR;
                                } else {
                                    // Old format (plain string vendor code — backwards compat)
                                    obj._chargeVendors[condType] = entry;
                                }
                            }
                        };
                        if (vendorChargeMap[poNum]['*']) applyEntries(vendorChargeMap[poNum]['*']);
                        if (lineItem && vendorChargeMap[poNum][lineItem]) applyEntries(vendorChargeMap[poNum][lineItem]);
                    }
                }

                rawRows.push(obj);
            }

            res.write(JSON.stringify({
                type: 'progress',
                message: `Parsed ${rowIndex - (headerRowIndex + 1)} of ${totalRows - (headerRowIndex + 1)} rows`,
                rowsParsed: rowIndex - (headerRowIndex + 1),
                totalRows: totalRows - (headerRowIndex + 1)
            }) + '\n');

            if (rowIndex < totalRows) {
                setImmediate(processBatch);
            } else {
                finalizeUpload();
            }
        }

        function finalizeUpload() {
            // Build global vendor map (for PO only)
            if (!isGRN && !isPurchase && !isStockJournal && !isSalesOrder && !isFI) {
                globalVendorMap = {};
                rawRows.forEach(row => {
                    const vCode = getRowValue(row, 'Vendor');
                    const vName = getRowValue(row, 'Vendor Name');
                    if (vCode !== undefined && vCode !== null && vName !== undefined && vName !== null) {
                        const cleanCode = String(vCode).split('.')[0].trim();
                        globalVendorMap[cleanCode] = String(vName).trim();
                    }
                });
            }

            // Filter entries if applicable
            let filteredRows = rawRows;
            if (isGRN) {
                filteredRows = rawRows.filter(row => {
                    const eventType = String(getRowValue(row, 'Trans./Event Type') || '').trim().toUpperCase();
                    if (eventType === 'WE') {
                        const vendorVal = getRowValue(row, 'Vendor');
                        if (vendorVal === undefined || vendorVal === null || String(vendorVal).trim() === '') return false;

                        const poType = String(getRowValue(row, 'Purchase Order type') || '').trim().toUpperCase();
                        if (poType === 'ZSTO') return false;

                        return true;
                    }
                    return false;
                });
            } else if (isStockJournal) {
                filteredRows = rawRows.filter(row => {
                    const eventType = String(getRowValue(row, 'Trans./Event Type') || getRowValue(row, 'Trans./Event TypeA') || '').trim().toUpperCase();
                    if (eventType === 'WA') {
                        const plant = getRowValue(row, 'Receiving Plant') || getRowValue(row, 'Plant');
                        if (plant === undefined || plant === null || String(plant).trim() === '') return false;
                        return true;
                    }
                    return false;
                });
            }

            // Group rows by 'Sales document', 'Material Document', 'Document Number' / 'Invoice No', or 'Purchasing Document'
            const groups = {};
            filteredRows.forEach(row => {
                let groupNum;
                if (isSalesOrder) {
                    groupNum = getRowValue(row, 'Sales document') || getRowValue(row, 'Document Number');
                } else if (isGRN || isStockJournal) {
                    groupNum = getRowValue(row, 'Material Document') || getRowValue(row, 'GRN Number') || getRowValue(row, 'Purchasing Document');
                } else if (isPurchase || isFI) {
                    groupNum = getRowValue(row, 'Document Number') || getRowValue(row, 'Invoice No') || getRowValue(row, 'Invoice Number');
                } else {
                    groupNum = getRowValue(row, 'Purchasing Document') || getRowValue(row, 'PO Number') || getRowValue(row, 'Document Number');
                }

                if (groupNum !== undefined && groupNum !== null && String(groupNum).trim() !== '') {
                    const groupKey = String(groupNum).split('.')[0].trim();
                    if (groupKey && groupKey.toUpperCase() !== 'X' && groupKey.toLowerCase() !== 'undefined' && groupKey.toLowerCase() !== 'null') {
                        if (!groups[groupKey]) {
                            groups[groupKey] = {
                                poNumber: groupKey,
                                items: []
                            };
                        }
                        groups[groupKey].items.push(row);
                    }
                }
            });

            const ignoredPurchaseInvoices = [];

            const poList = [];
            Object.values(groups).forEach(poGroup => {
                const firstRow = poGroup.items[0];
                let docType = 'ZSPR';
                let vendorName = '';
                if (isSalesOrder) {
                    const salesDocType = String(getRowValue(firstRow, 'Sales Document Type') || 'ZASH').trim();
                    docType = `Sales Order ${salesDocType}`;
                    const partyCode = String(getRowValue(firstRow, 'Ship-to party') || getRowValue(firstRow, 'Party') || '').split('.')[0].trim();
                    vendorName = partyCode || 'Customer';
                } else if (isStockJournal) {
                    docType = 'WA';
                    const recPlant = String(getRowValue(firstRow, 'Receiving Plant') || getRowValue(firstRow, 'Plant') || '').trim();
                    vendorName = recPlant ? `Stock Transfer (${recPlant})` : 'Stock Transfer';
                } else if (isGRN) {
                    docType = String(getRowValue(firstRow, 'Trans./Event Type') || 'GRN').trim();
                    vendorName = String(getRowValue(firstRow, 'Vendor Description') || '').trim();
                } else if (isPurchase) {
                    const rawDocType = String(
                        getRowValue(firstRow, 'Purchasing Doc Type') ||
                        getRowValue(firstRow, 'Purchasing Doc. Type') ||
                        getRowValue(firstRow, 'PO - Doc Type') ||
                        getRowValue(firstRow, 'Doc Type') ||
                        getRowValue(firstRow, 'Purchase Order Type') ||
                        ''
                    ).trim();

                    if (rawDocType) {
                        docType = rawDocType;
                    } else {
                        docType = 'ZSPR';
                    }

                    const rawParty = getRowValue(firstRow, 'Invoicing Party') || getRowValue(firstRow, 'Vendor');
                    const rawVendorName = getRowValue(firstRow, 'Vendor Name') || getRowValue(firstRow, 'Vendor Description');
                    vendorName = rawVendorName ? String(rawVendorName).trim() : (rawParty ? String(rawParty).split('.')[0].trim() : 'Unknown Vendor');
                } else if (isFI) {
                    const rawDocType = String(
                        getRowValue(firstRow, 'Doc Type') ||
                        getRowValue(firstRow, 'Doc type') ||
                        getRowValue(firstRow, 'Document Type') ||
                        'FI'
                    ).trim();
                    docType = rawDocType;
                    // Vendor priority: Vendor column first, else Customer, else G/L Account
                    const vendorVal = getRowValue(firstRow, 'Vendor');
                    const customerVal = getRowValue(firstRow, 'Customer');
                    const glVal = getRowValue(firstRow, 'G/L Account') || getRowValue(firstRow, 'G/L Account_1');

                    if (vendorVal !== undefined && vendorVal !== null && String(vendorVal).trim() !== '' && String(vendorVal).trim() !== '0') {
                        vendorName = padVendor(vendorVal);
                    } else if (customerVal !== undefined && customerVal !== null && String(customerVal).trim() !== '' && String(customerVal).trim() !== '0') {
                        vendorName = padVendor(customerVal);
                    } else if (glVal) {
                        vendorName = String(glVal).split('.')[0].trim();
                    } else {
                        vendorName = 'Unknown Party';
                    }
                } else {
                    docType = String(getRowValue(firstRow, 'Doc Type') || getRowValue(firstRow, 'PO - Doc Type') || getRowValue(firstRow, 'Document Type') || 'ZSPR').trim();
                    vendorName = String(getRowValue(firstRow, 'Vendor Name') || '').trim();
                }

                /*
                if (isPurchase) {
                    // Check if any active stock line item (where Condition Type is empty) is missing Reference Document / GRN tracking number
                    const hasMissingRefDoc = poGroup.items.some(item => {
                        const delInd = getRowValue(item, 'Deletion Indicator');
                        if (delInd && String(delInd).trim().toUpperCase() === 'L') return false;
                        
                        // Ignore charge / fee rows (rows that have a Condition Type value)
                        const condType = getRowValue(item, 'Condition Type');
                        if (condType && String(condType).trim() !== '') return false;

                        const rawRefDoc = getRowValue(item, 'Reference Document') || getRowValue(item, 'Reference Document Item') || getRowValue(item, 'Ref Document') || getRowValue(item, 'GRN Number') || getRowValue(item, 'GRN');
                        if (rawRefDoc === undefined || rawRefDoc === null) return true;
                        const strVal = String(rawRefDoc).trim().toUpperCase();
                        return strVal === '' || strVal === '#N/A' || strVal === '#N/A!' || strVal === 'N/A' || strVal === 'NULL' || strVal === 'UNDEFINED';
                    });

                    if (hasMissingRefDoc) {
                        ignoredPurchaseInvoices.push({
                            poNumber: poGroup.poNumber,
                            docType,
                            vendorName,
                            itemCount: poGroup.items.length,
                            items: poGroup.items,
                            reason: 'Line item missing Reference Document (GRN Tracking Number)'
                        });
                        return; // Do not include in poList
                    }
                }
                */

                poList.push({
                    poNumber: poGroup.poNumber,
                    docType,
                    vendorName,
                    itemCount: poGroup.items.length,
                    items: poGroup.items
                });
            });

            if (poList.length === 0) {
                const label = isFI ? 'Financial Entry (FI)' : (isStockJournal ? 'Stock Journal' : (isGRN ? 'GRN' : (isPurchase ? 'Purchase Invoice' : 'purchase order')));
                res.write(JSON.stringify({ type: 'error', error: `No valid ${label} entries found in the sheet.` }) + '\n');
                return res.end();
            }

            const label = isFI ? 'Financial Entry (FI)' : (isStockJournal ? 'Stock Journal' : (isGRN ? 'GRN' : (isPurchase ? 'Purchase Invoice' : 'Purchase Order')));
            res.write(JSON.stringify({
                type: 'result',
                message: `Excel file processed successfully. Found ${poList.length} ${label} records.${ignoredPurchaseInvoices.length > 0 ? ` Ignored ${ignoredPurchaseInvoices.length} invoices with missing Reference Document.` : ''}`,
                poList,
                ignoredPurchaseInvoices
            }) + '\n');
            res.end();
        }

        // Start processing batch
        processBatch();

    } catch (error) {
        console.error('Error processing Excel file:', error);
        cleanupFile(filePath);
        if (!res.headersSent) {
            return res.status(500).json({ error: 'Internal server error while processing sheet', details: error.message });
        } else {
            res.write(JSON.stringify({ type: 'error', error: error.message }) + '\n');
            return res.end();
        }
    }
});

/**
 * POST /api/import
 * Takes selected Purchase Orders or GRNs, generates XML, and pushes to Tally.
 */
function parseTallyResponse(xmlStr) {
    if (!xmlStr || typeof xmlStr !== 'string') {
        return { created: 0, altered: 0, errors: 0, exceptions: 0, lineError: '', isSuccess: false };
    }
    const getVal = (tag) => {
        const match = xmlStr.match(new RegExp(`<${tag}>(.*?)</${tag}>`, 'i'));
        return match ? match[1].trim() : '';
    };

    const created = parseInt(getVal('CREATED') || '0', 10);
    const rawAltered = parseInt(getVal('ALTERED') || '0', 10);
    const altered = Math.floor(rawAltered / 2);
    const errors = parseInt(getVal('ERRORS') || '0', 10);
    const exceptions = parseInt(getVal('EXCEPTIONS') || '0', 10);
    const lineError = getVal('LINEERROR');

    const hasError = errors > 0 || exceptions > 0 || lineError !== '';
    const isSuccess = (created > 0 || rawAltered > 0) && !hasError;

    return {
        created,
        altered,
        errors,
        exceptions,
        lineError,
        isSuccess
    };
}

app.post('/api/import', async (req, res) => {
    const { selectedPOs, importType, skipBlankMaterial } = req.body;
    if (!selectedPOs || !Array.isArray(selectedPOs) || selectedPOs.length === 0) {
        return res.status(400).json({ error: 'No records selected for import' });
    }

    const isGRN = importType === 'grn';
    const isPurchase = importType === 'purchase';
    const isStockJournal = importType === 'stock_journal';
    const results = [];
    try {
        for (const poGroup of selectedPOs) {
            if (skipBlankMaterial) {
                poGroup.items = poGroup.items.filter(item => {
                    const mat = getRowValue(item, 'Material');
                    return mat !== undefined && mat !== null && String(mat).trim() !== '';
                });
            }

            const activeItems = poGroup.items.filter(item => {
                const delInd = getRowValue(item, 'Deletion Indicator');
                return !(delInd && String(delInd).trim().toUpperCase() === 'L');
            });

            if (activeItems.length === 0) {
                results.push({
                    poNumber: poGroup.poNumber,
                    vendorName: poGroup.vendorName || (poGroup.items[0] ? (isStockJournal ? 'Stock Transfer' : (isGRN ? poGroup.items[0]['Vendor Description'] : (isPurchase ? poGroup.items[0]['Invoicing Party'] : poGroup.items[0]['Vendor Name']))) : ''),
                    itemCount: 0,
                    status: 'skipped',
                    tallyResponse: `Skipped - Record has only deleted items (Deletion Indicator L)`,
                    tallyParsed: null,
                    error: null,
                    xmlGenerated: ''
                });
                continue;
            }

            const isSalesOrder = importType === 'sales_order';
            const isFI = importType === 'fi';
            const xmlPayload = isFI
                ? generateFITallyXML(poGroup)
                : (isSalesOrder
                    ? generateSalesOrderTallyXML(poGroup)
                    : (isStockJournal
                        ? generateStockJournalTallyXML(poGroup)
                        : (isGRN
                            ? generateGRNTallyXML(poGroup)
                            : (isPurchase
                                ? generatePurchaseTallyXML(poGroup, globalVendorMap)
                                : generateTallyXML(poGroup, globalVendorMap)))));
            let tallyResponse = null;
            let tallyParsed = null;
            let status = 'pending';
            let errorMsg = null;

            try {
                // Post XML to Tally server
                const response = await axios.post(TALLY_URL, xmlPayload, {
                    headers: {
                        'Content-Type': 'text/xml; charset=utf-8',
                    },
                    timeout: 120000 // 120 seconds (2 minutes) timeout for Tally voucher import
                });
                tallyResponse = response.data;
                tallyParsed = parseTallyResponse(tallyResponse);

                if (tallyParsed.exceptions > 0 || tallyParsed.errors > 0 || !tallyParsed.isSuccess) {
                    status = 'failed';
                    errorMsg = tallyParsed.lineError || `Tally import exception: exceptions=${tallyParsed.exceptions}, errors=${tallyParsed.errors}`;
                } else {
                    status = 'success';
                }
            } catch (err) {
                status = 'failed';
                errorMsg = err.message;
                if (err.response && err.response.data) {
                    tallyResponse = String(err.response.data);
                    tallyParsed = parseTallyResponse(tallyResponse);
                    if (tallyParsed.lineError) {
                        errorMsg = `${err.message} - ${tallyParsed.lineError}`;
                    }
                }
            }

            results.push({
                poNumber: poGroup.poNumber,
                vendorName: poGroup.vendorName || (poGroup.items[0] ? (isGRN ? poGroup.items[0]['Vendor Description'] : (isPurchase ? poGroup.items[0]['Invoicing Party'] : poGroup.items[0]['Vendor Name'])) : ''),
                itemCount: activeItems.length,
                status,
                tallyResponse,
                tallyParsed,
                error: errorMsg,
                xmlGenerated: xmlPayload
            });
        }

        return res.json({
            message: `Processed ${results.length} records`,
            results
        });

    } catch (error) {
        console.error('Error processing import:', error);
        return res.status(500).json({ error: 'Internal server error during import', details: error.message });
    }
});

app.post('/api/fetch-from-tally', async (req, res) => {
    try {
        const { fromDate, toDate } = req.body;
        if (!fromDate || !toDate) {
            return res.status(400).json({ error: 'fromDate and toDate are required.' });
        }

        // Build the XML request payload for exporting Voucher Register from Tally
        const xmlRequest = `
<ENVELOPE>
    <HEADER>
        <VERSION>1</VERSION>
        <TALLYREQUEST>Export</TALLYREQUEST>
        <TYPE>Data</TYPE>
        <ID>VOUCHER REGISTER</ID>
    </HEADER>
    <BODY>
        <DESC>
            <STATICVARIABLES>
                <SVFROMDATE TYPE="Date">${fromDate}</SVFROMDATE>
                <SVTODATE TYPE="Date">${toDate}</SVTODATE>
                <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
            </STATICVARIABLES>
        </DESC>
    </BODY>
</ENVELOPE>
`;

        console.log(`Sending XML request to Tally (${TALLY_URL}):`, xmlRequest);
        const response = await axios.post(TALLY_URL, xmlRequest, {
            headers: {
                'Content-Type': 'application/xml',
            },
            timeout: 60000 // Tally reports can be large, allow 60 seconds timeout
        });

        let content = response.data;
        // Clean up any invalid control character references like &#4;
        if (typeof content === 'string') {
            content = content.replace(/&#\d+;/g, '');
        }

        res.type('application/xml').send(content);
    } catch (err) {
        console.error('Error fetching data from Tally:', err.message);
        res.status(500).json({
            error: 'Failed to fetch data from Tally server.',
            details: err.message,
            tallyUrl: TALLY_URL
        });
    }
});

const { TEMPLATES_CONFIG, generateTemplateWorkbook, generateTemplateCSV } = require('./templatesConfig');

app.get('/api/templates/meta', (req, res) => {
    try {
        res.json(TEMPLATES_CONFIG);
    } catch (err) {
        res.status(500).json({ error: 'Failed to get templates metadata', details: err.message });
    }
});

app.get('/api/templates/download/:module', (req, res) => {
    try {
        const modKey = String(req.params.module || '').toLowerCase().trim();
        const conf = TEMPLATES_CONFIG[modKey];
        if (!conf) {
            return res.status(404).json({ error: `Template module '${modKey}' not found.` });
        }
        const buffer = generateTemplateWorkbook(modKey);
        const filename = `${conf.sheetName}.xlsx`;
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    } catch (err) {
        res.status(500).json({ error: 'Failed to generate template workbook', details: err.message });
    }
});

app.get('/api/templates/csv/:module', (req, res) => {
    try {
        const modKey = String(req.params.module || '').toLowerCase().trim();
        const conf = TEMPLATES_CONFIG[modKey];
        if (!conf) {
            return res.status(404).json({ error: `Template module '${modKey}' not found.` });
        }
        const csvContent = generateTemplateCSV(modKey);
        const filename = `${conf.sheetName}.csv`;
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Type', 'text/csv');
        res.send(csvContent);
    } catch (err) {
        res.status(500).json({ error: 'Failed to generate template CSV', details: err.message });
    }
});

// Create uploads folder if it doesn't exist
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir);
}

app.listen(PORT, () => {
    console.log(`Excel-to-Tally Backend Server listening on port ${PORT}`);
    console.log(`Targeting Tally Server at: ${TALLY_URL}`);
});
