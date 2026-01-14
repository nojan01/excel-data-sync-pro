const ExcelJS = require('exceljs');

async function test() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER Kopie 2.xlsx');
    
    const ws = workbook.getWorksheet(1);
    console.log('Sheet:', ws.name);
    console.log('AutoFilter:', ws.autoFilter);
    console.log('Tables:', Object.keys(ws.tables || {}));
    
    // Prüfe Header-Zeile (Zeile 1) - erste 5 Zellen
    console.log('\n--- Header-Zeile (Zeile 1) ---');
    for (let col = 1; col <= 5; col++) {
        const cell = ws.getCell(1, col);
        const hasFill = cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern === 'solid';
        console.log('Zelle ' + col + ':', {
            value: String(cell.value).substring(0, 20),
            hasFill: hasFill,
            fillColor: hasFill ? cell.fill.fgColor?.argb : 'none'
        });
    }
    
    // Test Export
    console.log('\n--- Export Test ---');
    const exportPath = '/Users/nojan/Desktop/Test_Fill_Export.xlsx';
    await workbook.xlsx.writeFile(exportPath);
    console.log('Exportiert nach:', exportPath);
    
    // Wieder einlesen
    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(exportPath);
    const ws2 = workbook2.getWorksheet(1);
    
    console.log('\n--- Nach Export - Header-Zeile ---');
    for (let col = 1; col <= 5; col++) {
        const cell = ws2.getCell(1, col);
        const hasFill = cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern === 'solid';
        console.log('Zelle ' + col + ':', {
            hasFill: hasFill,
            fillColor: hasFill ? cell.fill.fgColor?.argb : 'none'
        });
    }
}

test().catch(console.error);
