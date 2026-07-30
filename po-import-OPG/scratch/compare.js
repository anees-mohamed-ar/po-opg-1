const fs = require('fs');
const path = require('path');

const originalPath = path.join(__dirname, '../../Purchase_5105646126.xml');
const generatedPath = path.join(__dirname, 'test_purchase_output.xml');

if (fs.existsSync(originalPath) && fs.existsSync(generatedPath)) {
    let orig = fs.readFileSync(originalPath, 'utf8');
    if (!orig.trim().startsWith('<')) orig = fs.readFileSync(originalPath, 'utf16le');
    
    let gen = fs.readFileSync(generatedPath, 'utf8');
    
    // Compare basic fields
    const getField = (content, regex) => {
        const m = content.match(regex);
        return m ? m[1] : 'NOT FOUND';
    };
    
    console.log('--- BASIC HEADER FIELD COMPARISON ---');
    console.log('Original VOUCHERNUMBER:', getField(orig, /<VOUCHERNUMBER>(.*?)<\/VOUCHERNUMBER>/));
    console.log('Generated VOUCHERNUMBER:', getField(gen, /<VOUCHERNUMBER>(.*?)<\/VOUCHERNUMBER>/));
    
    console.log('Original PARTYLEDGERNAME:', getField(orig, /<PARTYLEDGERNAME>(.*?)<\/PARTYLEDGERNAME>/));
    console.log('Generated PARTYLEDGERNAME:', getField(gen, /<PARTYLEDGERNAME>(.*?)<\/PARTYLEDGERNAME>/));
    
    console.log('Original REFERENCE:', getField(orig, /<REFERENCE>(.*?)<\/REFERENCE>/));
    console.log('Generated REFERENCE:', getField(gen, /<REFERENCE>(.*?)<\/REFERENCE>/));
    
    console.log('Original DATE:', getField(orig, /<DATE>(.*?)<\/DATE>/));
    console.log('Generated DATE:', getField(gen, /<DATE>(.*?)<\/DATE>/));

    // Compare Inventory entries
    const getInvEntries = (content) => {
        const list = [];
        const regex = /<ALLINVENTORYENTRIES\.LIST>([\s\S]*?)<\/ALLINVENTORYENTRIES\.LIST>/g;
        let match;
        while ((match = regex.exec(content)) !== null) {
            const entry = match[1];
            const stockName = entry.match(/<STOCKITEMNAME>(.*?)<\/STOCKITEMNAME>/)?.[1] || 'N/A';
            const rate = entry.match(/<RATE>(.*?)<\/RATE>/)?.[1] || 'N/A';
            const qty = entry.match(/<BILLEDQTY>(.*?)<\/BILLEDQTY>/)?.[1] || 'N/A';
            const amount = entry.match(/<AMOUNT>(.*?)<\/AMOUNT>/)?.[1] || 'N/A';
            const ledger = entry.match(/<LEDGERNAME>(.*?)<\/LEDGERNAME>/)?.[1] || 'N/A';
            list.push({ stockName, rate, qty, amount, ledger });
        }
        return list;
    };
    
    const origInv = getInvEntries(orig);
    const genInv = getInvEntries(gen);
    
    console.log(`\n--- INVENTORY ENTRIES COMPARISON ---`);
    console.log(`Original entries: ${origInv.length}, Generated entries: ${genInv.length}`);
    
    console.log('\n--- GENERATED DETAILS: ---');
    genInv.forEach((e, idx) => {
        console.log(`Entry #${idx+1}: Stock=${e.stockName}, Qty=${e.qty}, Rate=${e.rate}, Amount=${e.amount}, Ledger=${e.ledger}`);
    });
    
    // Compare Party Ledgers
    const getPartyLedgerAmount = (content) => {
        const regex = /<LEDGERENTRIES\.LIST>([\s\S]*?)<\/LEDGERENTRIES\.LIST>/g;
        let match;
        while ((match = regex.exec(content)) !== null) {
            const entry = match[1];
            const isParty = entry.includes('<ISPARTYLEDGER>Yes</ISPARTYLEDGER>');
            if (isParty) {
                return entry.match(/<AMOUNT>(.*?)<\/AMOUNT>/)?.[1];
            }
        }
        return 'NOT FOUND';
    };
    
    console.log('\n--- PARTY LEDGER AMOUNT COMPARISON ---');
    console.log('Original Party Ledger Amount:', getPartyLedgerAmount(orig));
    console.log('Generated Party Ledger Amount:', getPartyLedgerAmount(gen));
}
