const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();

wb.xlsx.readFile('/Users/nojan/Desktop/test-styles Kopie.xlsx').then(() => {
    const ws = wb.getWorksheet('Style Tests');
    
    console.log('=== Zeile 4 KOMPLETT (alle Properties) ===');
    const row4 = ws.getRow(4);
    for (let col = 1; col <= 8; col++) {
        const cell = row4.getCell(col);
        console.log('\nCol', col, ':', cell.value);
        console.log('  fill:', JSON.stringify(cell.fill));
        
        // Prüfe Theme Colors
        if (cell.fill && cell.fill.fgColor) {
            console.log('  fgColor keys:', Object.keys(cell.fill.fgColor));
            if (cell.fill.fgColor.theme !== undefined) {
                console.log('  THEME COLOR:', cell.fill.fgColor.theme, 'tint:', cell.fill.fgColor.tint);
            }
            if (cell.fill.fgColor.indexed !== undefined) {
                console.log('  INDEXED COLOR:', cell.fill.fgColor.indexed);
            }
        }
    }
    
    // Auch Workbook Theme prüfen
    console.log('\n=== Workbook Theme ===');
    if (wb.theme) {
        console.log('Theme vorhanden:', Object.keys(wb.theme));
    } else {
        console.log('Kein Theme gefunden');
    }
}).catch(e => console.error(e));
