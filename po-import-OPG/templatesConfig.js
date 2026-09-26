const xlsx = require('xlsx');

/**
 * Clean specification of columns required by the program for each module.
 * Only includes columns actually referenced/used in parsing and XML generation.
 */
const TEMPLATES_CONFIG = {
    po: {
        id: 'po',
        title: 'Purchase Order (PO)',
        sheetName: 'PO_Template',
        originalTotalCols: 132,
        description: 'Template for SAP Purchase Orders. Contains order details, vendor info, item lines, and applicable condition charges.',
        columns: [
            // --- Header / Grouping ---
            { name: 'Purchasing Document', required: true, sample: '4500003396', description: 'Unique PO Number (used as Voucher Number in Tally)' },
            { name: 'PO - Doc Type', required: true, sample: 'ZSPR', description: 'PO Document Type (e.g. ZSPR, ZCOL, ZCON) — determines voucher type and ledger name' },
            { name: 'Document Date', required: true, sample: '01.04.2023', description: 'PO creation date (DD.MM.YYYY)' },
            // --- Vendor Info ---
            { name: 'Vendor', required: true, sample: '11075', description: 'Vendor code (padded to 10 digits in Tally)' },
            { name: 'Vendor Name', required: true, sample: 'Industrial Aids & Packings', description: 'Vendor legal name' },
            { name: 'Street', required: true, sample: '123 Industrial Area', description: 'Vendor street address' },
            { name: 'City', required: true, sample: 'Chennai', description: 'Vendor city' },
            { name: 'Post Code', required: true, sample: '600001', description: 'Postal / PIN code' },
            { name: 'Region Name', required: true, sample: 'Tamil Nadu', description: 'Vendor state / region name (e.g. Tamil Nadu, Maharashtra)' },
            { name: 'GST NO', required: true, sample: '33AAAAA0000A1Z5', description: 'Vendor GSTIN' },
            { name: 'Destination region', required: true, sample: '33', description: 'Destination state code (e.g. 33 for Tamil Nadu) — determines CGST+SGST vs IGST' },
            // --- Line Item ---
            { name: 'Deletion Indicator', required: true, sample: '', description: "Set to 'L' if line item is deleted/cancelled; leave blank otherwise" },
            { name: 'Material', required: true, sample: '111963', description: 'Material / stock item code' },
            { name: 'Short Text', required: true, sample: 'M.S. HEX BOLT & NUT', description: 'Material description / line item text' },
            { name: 'Order Quantity', required: true, sample: 100, description: 'Ordered quantity' },
            { name: 'Order Unit', required: true, sample: 'NOS', description: 'Unit of measurement (NOS, KG, MT, etc.)' },
            { name: 'Net Order Price', required: true, sample: 52.50, description: 'Unit price of the item' },
            { name: 'Net Order Value', required: true, sample: 5250.00, description: 'Total item net amount (Qty x Price)' },
            { name: 'Effective value', required: false, sample: 5250.00, description: 'Effective value — used ONLY for ZCOL type orders; leave blank otherwise' },
            { name: 'Currency', required: true, sample: 'INR', description: 'Document currency (use INR for domestic)' },
            { name: 'Exchange Rate', required: false, sample: 1, description: 'Exchange rate to INR — use 1 for INR; fill for USD/EUR etc.' },
            // --- Condition / Charge Columns (include only columns that have values; all others can be blank) ---
            { name: 'JEXS', required: false, sample: 0, description: 'Excise / GST tax condition — splits into CGST+SGST (local) or IGST (inter-state) in Tally' },
            { name: 'NAVS', required: false, sample: 0, description: 'Non-deductible tax — splits into CGST+SGST or IGST in Tally' },
            { name: 'FRB1', required: false, sample: 0, description: 'Freight charge (value-based)' },
            { name: 'FRB2', required: false, sample: 0, description: 'Freight charge (alternate)' },
            { name: 'FRC2', required: false, sample: 0, description: 'Freight charge (quantity-based)' },
            { name: 'P001', required: false, sample: 0, description: 'Price condition P001' },
            { name: 'P101', required: false, sample: 0, description: 'Price condition P101' },
            { name: 'PB00', required: false, sample: 0, description: 'Gross base price condition' },
            { name: 'R003', required: false, sample: 0, description: 'Rebate condition' },
            { name: 'RA00', required: false, sample: 0, description: 'Rebate amount' },
            { name: 'RA01', required: false, sample: 0, description: 'Rebate amount 2' },
            { name: 'SKTO', required: false, sample: 0, description: 'Cash discount' },
            { name: 'WOTB', required: false, sample: 0, description: 'Without tax base' },
            { name: 'ZBCA', required: false, sample: 0, description: 'Bank charges' },
            { name: 'ZBCD', required: false, sample: 0, description: 'Bank charges (alternate)' },
            { name: 'ZBED', required: false, sample: 0, description: 'Basic excise duty' },
            { name: 'ZCEC', required: false, sample: 0, description: 'Central excise cess' },
            { name: 'ZCEQ', required: false, sample: 0, description: 'Central excise cess (qty-based)' },
            { name: 'ZCST', required: false, sample: 0, description: 'CST charge' },
            { name: 'ZDSD', required: false, sample: 0, description: 'Discount' },
            { name: 'ZEQP', required: false, sample: 0, description: 'Equipment charges' },
            { name: 'ZFRQ', required: false, sample: 0, description: 'Freight qty-based' },
            { name: 'ZFRV', required: false, sample: 0, description: 'Freight value-based' },
            { name: 'ZHAN', required: false, sample: 0, description: 'Handling charges' },
            { name: 'ZINS', required: false, sample: 0, description: 'Insurance charge' },
            { name: 'ZLAN', required: false, sample: 0, description: 'Landing charges' },
            { name: 'ZMFR', required: false, sample: 0, description: 'Manufacturing charges' },
            { name: 'ZMIS', required: false, sample: 0, description: 'Miscellaneous expenses' },
            { name: 'ZNE1', required: false, sample: 0, description: 'Surcharge 1' },
            { name: 'ZNE2', required: false, sample: 0, description: 'Surcharge 2' },
            { name: 'ZPAC', required: false, sample: 0, description: 'Packing & forwarding charges' },
            { name: 'ZPF%', required: false, sample: 0, description: 'Packing & forwarding (%)' },
            { name: 'ZPNF', required: false, sample: 0, description: 'P&F charges (value)' },
            { name: 'ZROY', required: false, sample: 0, description: 'Royalty charges' },
            { name: 'ZRTQ', required: false, sample: 0, description: 'Rate qty adjustment' },
            { name: 'ZRUQ', required: false, sample: 0, description: 'Rate unit qty' },
            { name: 'ZSIZ', required: false, sample: 0, description: 'Size charges' },
            { name: 'ZSTC', required: false, sample: 0, description: 'Service tax cess' },
            { name: 'ZSTP', required: false, sample: 0, description: 'Stamp charges' },
            { name: 'ZSTV', required: false, sample: 0, description: 'Service tax value' },
            { name: 'ZVAT', required: false, sample: 0, description: 'VAT' },
            { name: 'ZVIN', required: false, sample: 0, description: 'Vehicle insurance' },
            { name: 'ZWRF', required: false, sample: 0, description: 'Water and related fees' }
        ]
    },
    grn: {
        id: 'grn',
        title: 'Goods Receipt Note (GRN)',
        sheetName: 'GRN_Template',
        originalTotalCols: 28,
        description: 'Template for Goods Receipt Notes (WE movement type). Rows with Trans./Event Type = WE and a valid Vendor are imported as Receipt Notes in Tally.',
        columns: [
            { name: 'Trans./Event Type', required: true, sample: 'WE', description: "Must be 'WE' for Goods Receipt. Rows with 'WA' are treated as Stock Journal transfers and ignored here." },
            { name: 'Material Document', required: true, sample: '5000548667', description: 'GRN Material Document number — used as Voucher Number in Tally' },
            { name: 'Document Date', required: true, sample: '01.04.2023', description: 'GRN Document Date (DD.MM.YYYY)' },
            { name: 'Posting Date', required: true, sample: '01.04.2023', description: 'GRN Posting Date' },
            { name: 'Purchase Order', required: true, sample: '4500003396', description: 'Reference Purchase Order Number — linked in Tally as Reference' },
            { name: 'Purchase Order type', required: false, sample: 'ZSPR', description: 'PO / Doc Type (e.g. ZSPR, ZSTO) — used in Tally Voucher Type & Ledger name' },
            { name: 'Vendor', required: true, sample: '11075', description: 'Vendor code — rows without a Vendor are excluded; used to look up vendor details from master' },
            { name: 'Vendor Description', required: true, sample: 'Industrial Aids & Packings', description: 'Vendor name (fallback if not found in vendor master)' },
            { name: 'Deletion Indicator', required: true, sample: '', description: "Set to 'L' if line is cancelled; leave blank otherwise" },
            { name: 'Material', required: true, sample: '111963', description: 'Received material / stock item code' },
            { name: 'Material Description', required: true, sample: 'M.S. HEX BOLT & NUT', description: 'Material short text / description' },
            { name: 'Qty in Un. of Entry', required: true, sample: 100, description: 'Received quantity' },
            { name: 'Unit of Entry', required: true, sample: 'NOS', description: 'Unit of measurement' },
            { name: 'Amount in LC', required: true, sample: 5250.00, description: 'Total item amount in local currency (INR)' }
        ]
    },
    purchase: {
        id: 'purchase',
        title: 'Purchase Invoice',
        sheetName: 'Purchase_Template',
        originalTotalCols: 69,
        description: 'Template for Purchase Invoices. Stock item rows must have blank Condition Type; charge rows (freight, tax, etc.) must have a Condition Type code.',
        columns: [
            // --- Header / Grouping ---
            { name: 'Invoice No', required: true, sample: '5105646126', description: 'SAP Invoice Document Number / Invoice No — all rows with same number form one Tally Purchase voucher' },
            { name: 'Posting Date', required: true, sample: '01.04.2023', description: 'Invoice Posting Date (DD.MM.YYYY)' },
            { name: 'Doc Date', required: false, sample: '01.04.2023', description: 'Document Date (DD.MM.YYYY)' },
            { name: 'Reference', required: true, sample: 'INV-2023-0091', description: "Vendor's actual invoice / bill number" },
            { name: 'Purchasing Doc Type', required: true, sample: 'ZCON', description: 'PO Document Type (e.g. ZCON, ZSPR) — determines Purchase ledger name in Tally' },
            { name: 'Invoicing Party', required: true, sample: '60105', description: 'Vendor / Invoicing party code — used as Party Ledger in Tally (10-digit padded)' },
            { name: 'Vendor Name', required: true, sample: 'Precision Tools Corp', description: 'Vendor legal name' },
            { name: 'Destination region', required: true, sample: '33', description: 'Destination state code (e.g. 33 for Tamil Nadu) — determines CGST+SGST vs IGST' },
            // --- Line Item ---
            { name: 'Deletion Indicator', required: true, sample: '', description: "Set to 'L' if line is deleted; leave blank otherwise" },
            { name: 'Condition Type', required: true, sample: '', description: 'Leave BLANK for material/stock lines. Fill with condition code (e.g. FRB1, JEXS) for charge-only lines.' },
            { name: 'Material', required: true, sample: '202487', description: 'Material code — for stock item lines; leave blank for charge-only lines' },
            { name: 'Purchase Order - Short Text', required: true, sample: 'String light with 30 holder', description: 'Item description / line text' },
            { name: 'Purchasing Document', required: true, sample: '4510001624', description: 'Associated Purchase Order number — linked as Order No in batch allocation' },
            { name: 'Plant', required: true, sample: '1000', description: 'Receiving plant / godown code — used as Godown in Tally' },
            { name: 'Reference Document', required: true, sample: '5000548667', description: 'GRN number — used as Tracking Number to link invoice to receipt in Tally' },
            { name: 'Quantity', required: true, sample: 4, description: 'Invoiced quantity' },
            { name: 'Order Unit', required: true, sample: 'NOS', description: 'Unit of measure' },
            { name: 'Amount', required: true, sample: 19200.00, description: 'Line item amount (net)' },
            { name: 'Value-Added Tax Amt', required: false, sample: 3456.00, description: 'VAT / GST tax amount — used to determine CGST/SGST/IGST ledger amounts in Tally' },
            { name: 'Currency', required: true, sample: 'INR', description: 'Document currency' },
            { name: 'Exchange Rate', required: false, sample: 1, description: 'Exchange rate to INR; use 1 for INR transactions' }
        ]
    },
    sales_order: {
        id: 'sales_order',
        title: 'Sales Order',
        sheetName: 'SalesOrder_Template',
        originalTotalCols: 81,
        description: 'Template for Sales Orders. Includes customer code, material, price components, and tax amounts.',
        columns: [
            // --- Header / Grouping ---
            { name: 'Sales document', required: true, sample: '3792', description: 'Sales Order Number — used as Voucher Number in Tally' },
            { name: 'Sales Document Type', required: true, sample: 'ZASH', description: 'Sales Doc Type code (e.g. ZASH, ZCOL) — determines voucher type name' },
            { name: 'Doc Date', required: true, sample: '01.04.2023', description: 'Sales document date (DD.MM.YYYY)' },
            { name: 'Ship-to party', required: true, sample: '90022480', description: 'Customer / Buyer code — used as Party Ledger in Tally (10-digit padded)' },
            { name: 'Destination region', required: true, sample: '33', description: 'Customer state code (e.g. 33 for Tamil Nadu) — determines CGST+SGST vs IGST' },
            // --- Line Item ---
            { name: 'Material', required: true, sample: 'COAL-I', description: 'Sold material / stock item code' },
            { name: 'Order Quantity', required: true, sample: 31.84, description: 'Sales quantity' },
            { name: 'Base Unit of Measure', required: true, sample: 'MT', description: 'Unit of measurement (MT, NOS, KG, etc.)' },
            // --- Price Components (fill only the applicable column; others can be blank) ---
            { name: 'ZASH - Base Price', required: false, sample: 14780.00, description: 'Base price per unit — for ZASH type orders' },
            { name: 'ZASH Qty * Base Price', required: false, sample: 470595.20, description: 'Total base price (Qty x ZASH rate) — used as line amount if filled' },
            { name: 'YBPR - Basic Price', required: false, sample: 0, description: 'Base price per unit — for YBPR type orders' },
            { name: 'YBPR - Qty * Base Price', required: false, sample: 0, description: 'Total base price (Qty x YBPR rate)' },
            { name: 'zpro - base price', required: false, sample: 0, description: 'Base price per unit — for ZPRO type orders' },
            { name: 'zpro - Qty * BasePrice', required: false, sample: 0, description: 'Total base price (Qty x ZPRO rate)' },
            { name: 'ZPRS - Base Price', required: false, sample: 0, description: 'Base price per unit — for ZPRS type orders' },
            // --- Additional Charge Components ---
            { name: 'YTRL - Transmission', required: false, sample: 0, description: 'Transmission / wheeling charges added to line amount' },
            { name: 'CESS', required: false, sample: 400.00, description: 'Cess charges added to line amount' },
            { name: 'ZREL', required: false, sample: 0, description: 'Railway / freight charges added to line amount' },
            { name: 'ZOM2 - Without Base amt', required: false, sample: 0, description: 'Handling / other charges (excl. base price) added to line amount' },
            { name: 'JWTH', required: false, sample: 0, description: 'Water charges / additional surcharge added to line amount' },
            // --- Tax ---
            { name: 'TAX amount', required: false, sample: 22892.96, description: 'Total GST / Tax Amount — split into CGST+SGST or IGST based on Destination region' }
        ]
    },
    fi: {
        id: 'fi',
        title: 'Financial Entry (FI)',
        sheetName: 'FI_Template',
        originalTotalCols: 51,
        description: 'Template for Financial Journal Entries (Vendor, Customer, and G/L Account postings). Each row = one ledger line. All rows with the same Document Number form one Tally Journal voucher.',
        columns: [
            { name: 'Document Number', required: true, sample: '90000000', description: 'SAP Financial Document Number — all rows with same Document Number = one Journal voucher in Tally' },
            { name: 'Doc Type', required: true, sample: 'KR', description: "Document type code — appended to voucher type name (e.g. 'Journal KR', 'Journal SA', 'Journal RV')" },
            { name: 'Posting Date', required: true, sample: '01.04.2023', description: 'Voucher posting date (DD.MM.YYYY)' },
            { name: 'Account Type', required: true, sample: 'K', description: "'K' = Vendor, 'D' = Customer, leave blank for G/L Account entries" },
            { name: 'Vendor', required: false, sample: '60105', description: 'Vendor code — required when Account Type is K; leave blank otherwise' },
            { name: 'Customer', required: false, sample: '', description: 'Customer code — required when Account Type is D; leave blank otherwise' },
            { name: 'G/L Account', required: false, sample: '400010', description: 'G/L Account code — required when Account Type is blank; leave blank for Vendor/Customer rows' },
            { name: 'Debit/Credit Ind.', required: true, sample: 'S', description: "'S' = Debit (positive in Tally), 'H' = Credit (negative in Tally)" },
            { name: 'Amount', required: true, sample: 15400.00, description: 'Absolute line amount (always enter as positive — Debit/Credit Ind. controls the sign)' },
            { name: 'Reference Key', required: false, sample: 'REF-2023-AP01', description: 'Used as Narration text in the Tally Journal voucher; leave blank if not applicable' },
            { name: 'Purchasing Document', required: false, sample: '4510001624', description: 'Linked PO Number — appears as Reference in Tally voucher; leave blank if not applicable' }
        ]
    },
    stock_journal: {
        id: 'stock_journal',
        title: 'Stock Journal (WA)',
        sheetName: 'StockJournal_Template',
        originalTotalCols: 28,
        description: 'Template for Stock Transfer / Inter-godown transfers (WA movement type). Imported as Stock Journal vouchers in Tally with source (OUT) and destination (IN) godowns.',
        columns: [
            { name: 'Trans./Event Type', required: true, sample: 'WA', description: "Must be 'WA' for stock transfer / inter-godown movement" },
            { name: 'Material Document', required: true, sample: '4900001234', description: 'Transfer Material Document Number — used as Voucher Number in Tally' },
            { name: 'Document Date', required: true, sample: '01.04.2023', description: 'Transfer document date (DD.MM.YYYY)' },
            { name: 'Posting Date', required: true, sample: '01.04.2023', description: 'Posting date of the transfer' },
            { name: 'Plant', required: true, sample: '1000', description: 'Source issuing plant / godown code (OUT side in Tally Stock Journal)' },
            { name: 'Receiving Plant', required: true, sample: '2000', description: 'Destination receiving plant / godown code (IN side in Tally Stock Journal)' },
            { name: 'Deletion Indicator', required: true, sample: '', description: "Set to 'L' if line is cancelled; leave blank otherwise" },
            { name: 'Material', required: true, sample: '111963', description: 'Material / stock item code being transferred' },
            { name: 'Material Description', required: true, sample: 'M.S. HEX BOLT & NUT', description: 'Material description' },
            { name: 'Qty in Un. of Entry', required: true, sample: 50, description: 'Quantity transferred (only positive-quantity rows are processed)' },
            { name: 'Unit of Entry', required: true, sample: 'NOS', description: 'Unit of measurement' },
            { name: 'Amount in LC', required: true, sample: 2625.00, description: 'Valuation amount in local currency (INR)' }
        ]
    },
    delivery_note: {
        id: 'delivery_note',
        title: 'Delivery Note (WL)',
        sheetName: 'DeliveryNote_Template',
        originalTotalCols: 30,
        description: 'Template for Goods Issue / Delivery Notes (WL Trans./Event Type). Imported as Delivery Note vouchers in Tally.',
        columns: [
            { name: 'Trans./Event Type', required: true, sample: 'WL', description: "Must be 'WL' for delivery note entries" },
            { name: 'Material Document', required: true, sample: '4900630047', description: 'Delivery Note Material Document Number — used as Voucher Number in Tally' },
            { name: 'Document Date', required: false, sample: '01.04.2021', description: 'Document date (DD.MM.YYYY)' },
            { name: 'Posting Date', required: true, sample: '01.04.2021', description: 'Posting date of the delivery note' },
            { name: 'Customer', required: true, sample: '379', description: 'Customer account code — used as Party Ledger (padded to 10 digits)' },
            { name: 'Goods recipient', required: false, sample: '0000000379', description: 'Goods recipient customer code' },
            { name: 'Material', required: true, sample: 'FLYASH', description: 'Material / stock item code' },
            { name: 'Material Desc', required: false, sample: 'FLYASH', description: 'Material description' },
            { name: 'Plant', required: true, sample: '1000', description: 'Issuing plant / godown code in Tally' },
            { name: 'Qty in Un. of Entry', required: true, sample: 22.21, description: 'Billed / delivered quantity' },
            { name: 'Unit of Entry', required: true, sample: 'MT', description: 'Unit of measure (e.g. MT, NOS)' },
            { name: 'Amount in LC', required: true, sample: 2221.00, description: 'Total item amount in local currency (INR)' },
            { name: 'Reference', required: false, sample: '0080040913', description: 'Outbound Delivery reference number' },
            { name: 'Destination region', required: false, sample: '33', description: 'State code (e.g. 33 for Tamil Nadu)' }
        ]
    },
    sales_invoice: {
        id: 'sales_invoice',
        title: 'Sales Invoice',
        sheetName: 'SalesInvoice_Template',
        originalTotalCols: 122,
        description: 'Template for Sales Invoices. Imported as Sales vouchers in Tally with automatic CGST+SGST/IGST tax calculation and computed narration fields.',
        columns: [
            { name: 'Billing Document', required: true, sample: '90015368', description: 'Unique Billing Document Number (used as Voucher Number in Tally)' },
            { name: 'Billing Type', required: true, sample: 'ZCOL', description: 'Billing Type (e.g. ZCOL, ZOM) — used for Sales ledger (e.g. Sales ZCOL)' },
            { name: 'Billing Date', required: true, sample: '01.04.2022', description: 'Billing Date (DD.MM.YYYY or Excel date)' },
            { name: 'Sales document', required: false, sample: '3745', description: 'Sales document / reference order number' },
            { name: 'Sold to Party', required: true, sample: '400017', description: 'Customer account code (padded to 10 digits)' },
            { name: 'Search Term', required: false, sample: 'CHENNAI', description: 'Customer name / search term for party ledger' },
            { name: 'GST NO', required: false, sample: '33AAECC0681N1ZL', description: 'Customer GSTIN' },
            { name: 'Destination region', required: true, sample: '33', description: 'State region code: 33 splits Tax Amount into CGST & SGST, otherwise IGST' },
            { name: 'Material', required: true, sample: 'COAL-I', description: 'Stock item / material code' },
            { name: 'Description', required: false, sample: 'Coal - Import (MV Flag VI)', description: 'Material description' },
            { name: 'Plant', required: false, sample: '1310', description: 'Plant / godown code' },
            { name: 'Batch', required: false, sample: '0000000529', description: 'Batch number' },
            { name: 'Billed Quantity', required: true, sample: 28.58, description: 'Billed quantity' },
            { name: 'Sales unit', required: true, sample: 'MT', description: 'Unit of measure (MT, LS, etc.)' },
            { name: 'Net value', required: true, sample: 265691.11, description: 'Total invoice item amount (Assessable Value = Net value - Tax amount)' },
            { name: 'Tax amount', required: false, sample: 12575.20, description: 'Tax amount to split or assign as IGST' },
            { name: 'Profit centre', required: false, sample: '1300', description: 'Profit centre (included in narration if provided)' },
            { name: 'Volt  -  Draw Voltage', required: false, sample: '110', description: 'Draw voltage (included in narration if provided)' },
            { name: 'Grp2', required: false, sample: '110', description: 'Group 2 code (included in narration if provided)' },
            { name: 'Grp2-Desc', required: false, sample: 'grp2 desc', description: 'Group 2 description (included in narration if provided)' },
            { name: 'Grp3', required: false, sample: 'CH2', description: 'Group 3 code (included in narration if provided)' },
            { name: 'Grp3  Desc', required: false, sample: 'Chennai EDC North', description: 'Group 3 description (included in narration if provided)' },
            { name: 'Std Start Date', required: false, sample: '', description: 'Standard start date (included in narration if provided)' },
            { name: 'Std End Date', required: false, sample: '', description: 'Standard end date (included in narration if provided)' }
        ]
    }
};

/**
 * Generate an Excel workbook buffer with headers and a sample row.
 */
function generateTemplateWorkbook(moduleKey) {
    const config = TEMPLATES_CONFIG[moduleKey];
    if (!config) throw new Error(`Unknown module key: ${moduleKey}`);

    const headers = config.columns.map(c => c.name);
    const sampleRow = config.columns.map(c => c.sample);

    const aoa = [headers, sampleRow];
    const ws = xlsx.utils.aoa_to_sheet(aoa);

    // Auto-fit column widths based on header length
    const colWidths = config.columns.map(c => ({
        wch: Math.max(c.name.length, String(c.sample || '').length, 12) + 2
    }));
    ws['!cols'] = colWidths;

    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, config.sheetName);

    return xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

/**
 * Generate CSV text with headers and a sample row.
 */
function generateTemplateCSV(moduleKey) {
    const config = TEMPLATES_CONFIG[moduleKey];
    if (!config) throw new Error(`Unknown module key: ${moduleKey}`);

    const escapeCsv = (val) => {
        if (val === undefined || val === null) return '';
        const str = String(val);
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
            return `"${str.replace(/"/g, '""')}"`;
        }
        return str;
    };

    const headerLine = config.columns.map(c => escapeCsv(c.name)).join(',');
    const sampleLine = config.columns.map(c => escapeCsv(c.sample)).join(',');

    return `${headerLine}\r\n${sampleLine}\r\n`;
}

module.exports = {
    TEMPLATES_CONFIG,
    generateTemplateWorkbook,
    generateTemplateCSV
};
