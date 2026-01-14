const { readSheetWithExcelJS } = require('./exceljs-reader');

async function testExtraction() {
    console.log('=== TESTE EXCELJS-READER ===\n');
    
    // Zuerst Sheet-Name herausfinden
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/test-styles.xlsx');
    const sheetName = wb.worksheets[0].name;
    console.log('Sheet-Name:', sheetName);
    
    const result = await readSheetWithExcelJS('/Users/nojan/Desktop/test-styles.xlsx', sheetName);
    
    console.log('\n=== RESULT OBJECT ===');
    console.log('Keys:', Object.keys(result));
    console.log('\nHeaders:', result.headers);
    console.log('\nRows:', result.rows);
    console.log('\nAnzahl Zeilen:', result.rows?.length || 'undefined');
    console.log('Anzahl extrahierte Styles:', Object.keys(result.cellStyles).length);
    console.log('Anzahl Formeln:', Object.keys(result.cellFormulas).length);
    console.log('Anzahl Hyperlinks:', Object.keys(result.cellHyperlinks).length);
    console.log('Anzahl RichText:', Object.keys(result.richTextCells).length);
    
    console.log('\n=== EXTRAHIERTE STYLES ===');
    Object.entries(result.cellStyles).forEach(([key, style]) => {
        console.log(`${key}:`, JSON.stringify(style));
    });
    
    console.log('\n=== ZELLEN MIT WERTEN (erste 10) ===');
    result.rows.slice(0, 10).forEach((row, idx) => {
        console.log(`Zeile ${idx}:`, row);
    });
}

testExtraction().catch(console.error);
