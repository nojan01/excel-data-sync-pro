const ExcelJS = require('exceljs');

async function testExport() {
    // Quell-Datei lesen
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/AutoFilter_Test.xlsx');
    
    const ws = workbook.getWorksheet(1);
    console.log('=== VOR EXPORT ===');
    console.log('Sheet:', ws.name);
    console.log('AutoFilter:', ws.autoFilter);
    
    for (let col = 1; col <= 4; col++) {
        const cell = ws.getCell(1, col);
        console.log('Zelle ' + String.fromCharCode(64 + col) + '1:', {
            value: cell.value,
            fill: JSON.stringify(cell.fill)
        });
    }
    
    // Export simulieren - einfach speichern
    const exportPath = '/Users/nojan/Desktop/Export_Test_Fill.xlsx';
    await workbook.xlsx.writeFile(exportPath);
    console.log('\nExportiert nach:', exportPath);
    
    // Exportierte Datei wieder lesen
    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(exportPath);
    
    const ws2 = workbook2.getWorksheet(1);
    console.log('\n=== NACH EXPORT (erneut gelesen) ===');
    console.log('AutoFilter:', ws2.autoFilter);
    
    for (let col = 1; col <= 4; col++) {
        const cell = ws2.getCell(1, col);
        console.log('Zelle ' + String.fromCharCode(64 + col) + '1:', {
            value: cell.value,
            fill: JSON.stringify(cell.fill)
        });
    }
}

testExport().catch(console.error);
