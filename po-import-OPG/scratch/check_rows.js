const { generateStockJournalTallyXML } = require('../tallyXMLBuilder');

const mockRows = [
    {
        'Material Document': '4900646535',
        'Material Doc.Item': '1',
        'Trans./Event Type': 'WA',
        'Doc Type': 'ZSTO',
        'Material': 'COAL-I',
        'Material Description': 'Coal - Import',
        'Plant': '1010',
        'Receiving Plant': '1000',
        'Debit/Credit Ind.': 'H',
        'Qty in Un. of Entry': 36.23,
        'Unit of Entry': 'MT',
        'Amount in LC': 310112.50,
        'Document Date': 44866
    },
    {
        'Material Document': '4900646535',
        'Material Doc.Item': '2',
        'Trans./Event Type': 'WA',
        'Doc Type': 'ZSTO',
        'Material': 'COAL-I',
        'Material Description': 'Coal - Import',
        'Plant': '1000',
        'Receiving Plant': '1010',
        'Debit/Credit Ind.': 'S',
        'Qty in Un. of Entry': 36.23,
        'Unit of Entry': 'MT',
        'Amount in LC': 310112.50,
        'Document Date': 44866
    }
];

const xml = generateStockJournalTallyXML({ items: mockRows });

const inCount = (xml.match(/<INVENTORYENTRIESIN\.LIST>/g) || []).length;
const outCount = (xml.match(/<INVENTORYENTRIESOUT\.LIST>/g) || []).length;
const inGodown = xml.match(/<INVENTORYENTRIESIN\.LIST>[\s\S]*?<GODOWNNAME>(.*?)<\/GODOWNNAME>/)?.[1];
const inAmount = xml.match(/<INVENTORYENTRIESIN\.LIST>[\s\S]*?<AMOUNT>(.*?)<\/AMOUNT>/)?.[1];
const outGodown = xml.match(/<INVENTORYENTRIESOUT\.LIST>[\s\S]*?<GODOWNNAME>(.*?)<\/GODOWNNAME>/)?.[1];
const outAmount = xml.match(/<INVENTORYENTRIESOUT\.LIST>[\s\S]*?<AMOUNT>(.*?)<\/AMOUNT>/)?.[1];
const destGodown = xml.match(/<DESTINATIONGODOWN>(.*?)<\/DESTINATIONGODOWN>/)?.[1];

console.log('inCount:', inCount, 'outCount:', outCount);
console.log('DESTINATIONGODOWN:', destGodown);
console.log('IN Godown:', inGodown, 'IN Amount:', inAmount);
console.log('OUT Godown:', outGodown, 'OUT Amount:', outAmount);




