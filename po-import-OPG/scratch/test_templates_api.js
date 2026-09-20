const http = require('http');

http.get('http://localhost:5001/api/templates/meta', res => {
    let data = '';
    res.on('data', chunk => data += chunk);
    res.on('end', () => {
        try {
            const json = JSON.parse(data);
            console.log('Template modules successfully fetched:', Object.keys(json));
            for (const k of Object.keys(json)) {
                console.log(`  - ${k}: ${json[k].columns.length} columns (Original had: ${json[k].originalTotalCols})`);
            }
        } catch (e) {
            console.error('Parse error:', e, data);
        }
    });
}).on('error', err => console.error('Error:', err.message));
