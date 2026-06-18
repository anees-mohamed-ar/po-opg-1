const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const axios = require('axios');
const path = require('path');
const fs = require('fs');
const { generateTallyXML, getRowValue } = require('./tallyXMLBuilder');

const app = express();
const PORT = process.env.PORT || 5000;
const TALLY_URL = process.env.TALLY_URL || 'http://localhost:9000';

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
        // Read the uploaded excel sheet
        const workbook = xlsx.readFile(filePath);
        
        // Find a sheet with 'detail' in the name
        let detailSheetName = null;
        for (const name of workbook.SheetNames) {
            if (name.toLowerCase().includes('detail')) {
                detailSheetName = name;
                break;
            }
        }

        if (!detailSheetName) {
            cleanupFile(filePath);
            return res.status(400).json({ 
                error: 'Could not find a sheet containing "detail" in its name.',
                availableSheets: workbook.SheetNames
            });
        }

        const sheet = workbook.Sheets[detailSheetName];
        // Parse sheet to JSON objects
        const rawRows = xlsx.utils.sheet_to_json(sheet);
        cleanupFile(filePath);

        // Build global vendor map
        globalVendorMap = {};
        rawRows.forEach(row => {
            const vCode = getRowValue(row, 'Vendor');
            const vName = getRowValue(row, 'Vendor Name');
            if (vCode !== undefined && vCode !== null && vName !== undefined && vName !== null) {
                const cleanCode = String(vCode).split('.')[0].trim();
                globalVendorMap[cleanCode] = String(vName).trim();
            }
        });

        // Group rows by 'Purchasing Document' (filtering out empty/invalid rows)
        const groups = {};
        rawRows.forEach(row => {
            // Skip rows marked with Deletion Indicator 'L'
            const delInd = getRowValue(row, 'Deletion Indicator');
            if (delInd && String(delInd).trim().toUpperCase() === 'L') {
                return;
            }

            const poNum = row['Purchasing Document'];
            if (poNum !== undefined && poNum !== null && String(poNum).trim() !== '') {
                const poKey = String(poNum).split('.')[0].trim();
                if (!groups[poKey]) {
                    groups[poKey] = {
                        poNumber: poKey,
                        items: []
                    };
                }
                groups[poKey].items.push(row);
            }
        });

        const poList = Object.values(groups).map(poGroup => {
            const firstRow = poGroup.items[0];
            const docType = String(getRowValue(firstRow, 'PO - Doc Type') || 'ZSPR').trim();
            const vendorName = String(getRowValue(firstRow, 'Vendor Name') || '').trim();
            return {
                poNumber: poGroup.poNumber,
                docType,
                vendorName,
                itemCount: poGroup.items.length,
                items: poGroup.items
            };
        });

        if (poList.length === 0) {
            return res.status(400).json({ error: 'No valid purchase orders found in the sheet.' });
        }

        return res.json({
            message: `Excel file processed successfully. Found ${poList.length} Purchase Orders.`,
            poList
        });

    } catch (error) {
        console.error('Error processing Excel file:', error);
        cleanupFile(filePath);
        return res.status(500).json({ error: 'Internal server error while processing sheet', details: error.message });
    }
});

/**
 * POST /api/import
 * Takes selected Purchase Orders, generates XML, and pushes to Tally.
 */
app.post('/api/import', async (req, res) => {
    const { selectedPOs } = req.body;
    if (!selectedPOs || !Array.isArray(selectedPOs) || selectedPOs.length === 0) {
        return res.status(400).json({ error: 'No purchase orders selected for import' });
    }

    const results = [];
    try {
        for (const poGroup of selectedPOs) {
            const xmlPayload = generateTallyXML(poGroup, globalVendorMap);
            let tallyResponse = null;
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
                status = 'success';
            } catch (err) {
                status = 'failed';
                errorMsg = err.message;
                if (err.response && err.response.data) {
                    tallyResponse = String(err.response.data);
                }
            }

            results.push({
                poNumber: poGroup.poNumber,
                vendorName: poGroup.vendorName || (poGroup.items[0] ? poGroup.items[0]['Vendor Name'] : ''),
                itemCount: poGroup.items.length,
                status,
                tallyResponse,
                error: errorMsg,
                xmlGenerated: xmlPayload
            });
        }

        return res.json({
            message: `Processed ${results.length} Purchase Orders`,
            results
        });

    } catch (error) {
        console.error('Error processing import:', error);
        return res.status(500).json({ error: 'Internal server error during import', details: error.message });
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
