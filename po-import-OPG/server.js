const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const axios = require('axios');
const path = require('path');
const fs = require('fs');
const { generateTallyXML, generateGRNTallyXML, generatePurchaseTallyXML, generateSalesOrderTallyXML, generateFITallyXML, getRowValue, loadPOMaster } = require('./tallyXMLBuilder');

const app = express();
const PORT = process.env.PORT || 5001;
const TALLY_URL = process.env.TALLY_URL || 'http://localhost:9321';

app.use(cors());
app.use(express.json());

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
app.post('/api/upload', upload.single('file'), async (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'No file uploaded' });
    }

    const filePath = req.file.path;
    try {
        const importType = String(req.query.importType || req.body.importType || 'po').trim().toLowerCase();
        const isGRN = importType === 'grn';
        const isPurchase = importType === 'purchase';
        const isStockJournal = importType === 'stock_journal';
        const isSalesOrder = importType === 'sales_order';
        const isFI = importType === 'fi';

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
        for (let i = 0; i < Math.min(rawGrid.length, 50); i++) {
            const row = rawGrid[i];
            if (row && Array.isArray(row)) {
                const hasHeader = row.some(cell => {
                    const str = String(cell || '').trim();
                    return str === 'Purchasing Document' || 
                           str === 'Document Number' || 
                           str === 'Invoice No' || 
                           str === 'Invoice Number' || 
                           str === 'Material Document' ||
                           str === 'Sales document' ||
                           str === 'Invoicing Party' ||
                           str === 'Vendor';
                });
                if (hasHeader) {
                    headerRowIndex = i;
                    break;
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
                        ''
                    ).trim();

                    if (rawDocType) {
                        docType = rawDocType;
                    } else {
                        const poNumber = String(getRowValue(firstRow, 'Purchasing Document') || getRowValue(firstRow, 'Purchase Order') || '').split('.')[0].trim();
                        const poMaster = loadPOMaster();
                        if (poNumber && poMaster[poNumber]) {
                            docType = poMaster[poNumber];
                        } else {
                            docType = 'ZSPR';
                        }
                    }

                    const rawParty = getRowValue(firstRow, 'Invoicing Party') || getRowValue(firstRow, 'Vendor');
                    const rawVendorName = getRowValue(firstRow, 'Vendor Name') || getRowValue(firstRow, 'Vendor Description');
                    vendorName = rawVendorName ? String(rawVendorName).trim() : (rawParty ? String(rawParty).split('.')[0].trim() : 'Unknown Vendor');
                } else if (isFI) {
                    const rawDocType = String(getRowValue(firstRow, 'Doc Type') || 'FI').trim();
                    docType = rawDocType;
                    // Vendor priority: Vendor column first, else G/L Account
                    const vendorVal = getRowValue(firstRow, 'Vendor');
                    const glVal = getRowValue(firstRow, 'G/L Account') || getRowValue(firstRow, 'G/L Account_1');
                    const partyCode = (vendorVal !== undefined && vendorVal !== null && String(vendorVal).trim() !== '' && String(vendorVal).trim() !== '0')
                        ? String(vendorVal).split('.')[0].trim()
                        : (glVal ? String(glVal).split('.')[0].trim() : 'Unknown Party');
                    const vendorDesc = getRowValue(firstRow, 'Vendor Name') || getRowValue(firstRow, 'Vendor Description') || getRowValue(firstRow, 'GST Partner') || getRowValue(firstRow, 'Line Item Desc') || getRowValue(firstRow, 'Text') || '';
                    vendorName = vendorDesc ? `${partyCode}-${String(vendorDesc).trim()}` : partyCode;
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
                    timeout: 5000 // 5 seconds timeout
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

app.get('/api/xml-files', (req, res) => {
    try {
        const parentPath = path.join(__dirname, '../');
        const files = fs.readdirSync(parentPath);
        const xmlFiles = files.filter(f => f.toLowerCase().endsWith('.xml'));
        res.json({ xmlFiles });
    } catch (err) {
        res.status(500).json({ error: 'Failed to list XML files', details: err.message });
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
