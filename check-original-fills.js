const ExcelJS = require('exceljs');

async function check() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/test-styles-exceljs.xlsx');
    
    const ws = wb.worksheets[0];
    
    console.log('=== ORIGINAL Zeile 4 - Fills prüfen ===');
    for (let col = 1; col <= 9; col++) {
        const cell = ws.getCell(4, col);
        const fill = cell.fill;
        
        console.log('Spalte ' + col + ':');
        console.log('  Wert: ' + cell.value);
        console.log('  Fill:', JSON.stringify(fill));
    }
}

check().catch(e => console.error(e));
