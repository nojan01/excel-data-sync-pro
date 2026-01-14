const ExcelJS = require('exceljs');
const path = require('path');

async function testSpliceOnExistingFile() {
    console.log('=== Test: spliceColumns auf existierender Datei ===\n');
    
    // Erstelle erst eine Test-Datei mit vielen Zeilen
    const wb1 = new ExcelJS.Workbook();
    const ws1 = wb1.addWorksheet('Test');
    
    // 100 Zeilen mit Styles
    for (let row = 1; row <= 100; row++) {
        for (let col = 1; col <= 5; col++) {
            const cell = ws1.getCell(row, col);
            cell.value = 'R' + row + 'C' + col;
            // Verschiedene Farben je nach Spalte
            const colors = ['FFFF0000', 'FF00FF00', 'FF0000FF', 'FFFFFF00', 'FFFF00FF'];
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors[col-1] } };
        }
    }
    
    // Einige leere Zeilen mit nur Styles (simuliert reale Daten)
    for (let row = 101; row <= 110; row++) {
        for (let col = 1; col <= 5; col++) {
            const cell = ws1.getCell(row, col);
            // Kein Wert, nur Style
            const colors = ['FFFF0000', 'FF00FF00', 'FF0000FF', 'FFFFFF00', 'FFFF00FF'];
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors[col-1] } };
        }
    }
    
    const testFile = path.join(__dirname, 'test-splice-existing.xlsx');
    await wb1.xlsx.writeFile(testFile);
    console.log('Test-Datei erstellt mit 110 Zeilen, 5 Spalten\n');
    
    // Jetzt laden wie der Writer es tut
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(testFile);
    const ws2 = wb2.getWorksheet('Test');
    
    console.log('VOR spliceColumns:');
    // Pruefe einige Zeilen
    [1, 50, 100, 105].forEach(row => {
        console.log('  Zeile ' + row + ':');
        for (let col = 1; col <= 4; col++) {
            const cell = ws2.getCell(row, col);
            console.log('    Spalte ' + col + ': "' + (cell.value || '-') + '", Fill=' + (cell.fill?.fgColor?.argb || 'keine'));
        }
    });
    
    // Spalte 2 loeschen (die gruene)
    console.log('\nFuehre spliceColumns(2, 1) aus...\n');
    ws2.spliceColumns(2, 1);
    
    console.log('NACH spliceColumns:');
    [1, 50, 100, 105].forEach(row => {
        console.log('  Zeile ' + row + ':');
        for (let col = 1; col <= 3; col++) {
            const cell = ws2.getCell(row, col);
            console.log('    Spalte ' + col + ': "' + (cell.value || '-') + '", Fill=' + (cell.fill?.fgColor?.argb || 'keine'));
        }
    });
    
    // Speichern
    const outputFile = path.join(__dirname, 'test-splice-existing-output.xlsx');
    await wb2.xlsx.writeFile(outputFile);
    console.log('\nGespeichert:', outputFile);
    
    // Neu laden und pruefen
    const wb3 = new ExcelJS.Workbook();
    await wb3.xlsx.readFile(outputFile);
    const ws3 = wb3.getWorksheet('Test');
    
    console.log('\nNACH Speichern und Neu-Laden:');
    [1, 50, 100, 105].forEach(row => {
        console.log('  Zeile ' + row + ':');
        for (let col = 1; col <= 3; col++) {
            const cell = ws3.getCell(row, col);
            console.log('    Spalte ' + col + ': "' + (cell.value || '-') + '", Fill=' + (cell.fill?.fgColor?.argb || 'keine'));
        }
    });
    
    console.log('\n=== Test abgeschlossen ===');
}

testSpliceOnExistingFile().catch(console.error);
