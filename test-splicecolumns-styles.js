const ExcelJS = require('exceljs');
const path = require('path');

async function testSpliceColumnsStyles() {
    console.log('=== Test: spliceColumns und Style-Erhaltung ===\n');
    
    // Erstelle eine Test-Workbook
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Test');
    
    // Header
    ws.getCell(1, 1).value = 'A';
    ws.getCell(1, 2).value = 'B - TO DELETE';
    ws.getCell(1, 3).value = 'C';
    ws.getCell(1, 4).value = 'D';
    
    // Zeile 2: Werte + Styles
    ws.getCell(2, 1).value = 'Wert A';
    ws.getCell(2, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } }; // Rot
    
    ws.getCell(2, 2).value = 'Wert B (loeschen)';
    ws.getCell(2, 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } }; // Gruen
    
    ws.getCell(2, 3).value = 'Wert C';
    ws.getCell(2, 3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0000FF' } }; // Blau
    
    ws.getCell(2, 4).value = 'Wert D';
    ws.getCell(2, 4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Gelb
    
    // Zeile 3: Leere Zellen MIT Styles
    ws.getCell(3, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } }; // Rot
    ws.getCell(3, 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } }; // Gruen
    ws.getCell(3, 3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0000FF' } }; // Blau
    ws.getCell(3, 4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Gelb
    
    console.log('VOR spliceColumns:');
    for (let col = 1; col <= 4; col++) {
        const cell2 = ws.getCell(2, col);
        const cell3 = ws.getCell(3, col);
        console.log('  Spalte ' + col + ': Wert="' + (cell2.value || '(leer)') + '", Fill2=' + JSON.stringify(cell2.fill?.fgColor) + ', Fill3=' + JSON.stringify(cell3.fill?.fgColor));
    }
    
    // Spalte 2 loeschen
    console.log('\nFuehre spliceColumns(2, 1) aus...\n');
    ws.spliceColumns(2, 1);
    
    console.log('NACH spliceColumns:');
    for (let col = 1; col <= 3; col++) {
        const cell2 = ws.getCell(2, col);
        const cell3 = ws.getCell(3, col);
        console.log('  Spalte ' + col + ': Wert="' + (cell2.value || '(leer)') + '", Fill2=' + JSON.stringify(cell2.fill?.fgColor) + ', Fill3=' + JSON.stringify(cell3.fill?.fgColor));
    }
    
    // Speichern und neu laden um zu testen
    const testPath = path.join(__dirname, 'test-splice-output.xlsx');
    await workbook.xlsx.writeFile(testPath);
    console.log('\nDatei gespeichert:', testPath);
    
    // Neu laden und pruefen
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(testPath);
    const ws2 = wb2.getWorksheet('Test');
    
    console.log('\nNACH Speichern und Neu-Laden:');
    for (let col = 1; col <= 3; col++) {
        const cell2 = ws2.getCell(2, col);
        const cell3 = ws2.getCell(3, col);
        console.log('  Spalte ' + col + ': Wert="' + (cell2.value || '(leer)') + '", Fill2=' + JSON.stringify(cell2.fill?.fgColor) + ', Fill3=' + JSON.stringify(cell3.fill?.fgColor));
    }
    
    console.log('\n=== Test abgeschlossen ===');
}

testSpliceColumnsStyles().catch(console.error);
