/**
 * Test für Python Bridge
 */
const pythonBridge = require('./python/python_bridge.js');

async function test() {
    console.log('Python Path:', pythonBridge.getPythonPath());
    
    // Test listSheets
    console.log('\nTesting listSheets...');
    const sheets = await pythonBridge.listSheets('/Users/nojan/Desktop/Testdatei.xlsx');
    console.log('Sheets:', JSON.stringify(sheets, null, 2));
    
    // Test readSheet
    console.log('\nTesting readSheet...');
    const result = await pythonBridge.readSheet('/Users/nojan/Desktop/Testdatei.xlsx', 'DEFENCE&SPACE 2026 Baseline');
    
    console.log('Success:', result.success);
    console.log('Rows:', result.rowCount);
    console.log('Columns:', result.columnCount);
    console.log('Headers (first 5):', result.headers.slice(0, 5));
    console.log('Cell Styles count:', Object.keys(result.cellStyles).length);
    console.log('Cell Fonts count:', Object.keys(result.cellFonts).length);
    console.log('Hidden Rows count:', result.hiddenRows.length);
    console.log('Default Font:', result.defaultFont);
}

test().catch(e => console.error('Error:', e.message));
