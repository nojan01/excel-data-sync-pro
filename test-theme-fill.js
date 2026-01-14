const ExcelJS = require('exceljs');

async function test() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER Kopie 2.xlsx');
    
    const ws = workbook.getWorksheet(1);
    
    // Prüfe Header-Zeile detailliert
    console.log('--- Header-Zeile Fill Details ---');
    for (let col = 1; col <= 5; col++) {
        const cell = ws.getCell(1, col);
        console.log('Zelle ' + col + ':');
        console.log('  fill:', JSON.stringify(cell.fill, null, 2));
    }
    
    // Prüfe Theme
    console.log('\n--- Workbook Theme ---');
    console.log('theme:', workbook.model?.themes ? 'vorhanden' : 'nicht vorhanden');
}

test().catch(console.error);
