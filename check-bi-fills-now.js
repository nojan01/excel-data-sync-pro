const ExcelJS = require('exceljs');

async function checkBIFills() {
    const filePath = '/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];
    
    console.log('=== PRÜFE FILLS IN SPALTE BI (61) ===\n');
    
    // Prüfe die ersten 30 Zeilen in BI
    let fillCount = 0;
    for (let row = 1; row <= 30; row++) {
        const cell = sheet.getCell(row, 61); // BI = 61
        const fill = cell.fill;
        if (fill && fill.type === 'pattern' && fill.pattern !== 'none') {
            fillCount++;
            console.log(`BI${row}: Fill = ${JSON.stringify(fill)}`);
        }
    }
    
    console.log(`\nZellen mit Fill in BI (Zeile 1-30): ${fillCount}`);
    
    // Vergleiche mit BH (60)
    console.log('\n=== VERGLEICHE MIT SPALTE BH (60) ===\n');
    let bhFillCount = 0;
    for (let row = 1; row <= 30; row++) {
        const cell = sheet.getCell(row, 60); // BH = 60
        const fill = cell.fill;
        if (fill && fill.type === 'pattern' && fill.pattern !== 'none') {
            bhFillCount++;
            console.log(`BH${row}: Fill = ${JSON.stringify(fill)}`);
        }
    }
    console.log(`\nZellen mit Fill in BH (Zeile 1-30): ${bhFillCount}`);
    
    // Prüfe auch column.style
    console.log('\n=== SPALTEN-STYLE ===');
    const colBI = sheet.getColumn(61);
    const colBH = sheet.getColumn(60);
    console.log('BI column style:', JSON.stringify(colBI.style));
    console.log('BH column style:', JSON.stringify(colBH.style));
    
    // Wie viele Spalten hat die Datei wirklich?
    console.log('\n=== DATEI-INFO ===');
    console.log('columnCount:', sheet.columnCount);
    console.log('Letzte Spalte:', sheet.getColumn(sheet.columnCount).letter);
}

checkBIFills().catch(console.error);
