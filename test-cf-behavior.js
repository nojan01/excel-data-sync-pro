const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function testCFBehavior() {
    // Erstelle eine Test-Datei mit CF auf Spalte A
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Test');
    
    // Füge Daten hinzu
    ws.getCell('A1').value = 'Spalte A';
    ws.getCell('B1').value = 'Spalte B';
    ws.getCell('C1').value = 'Spalte C';
    for (let i = 2; i <= 10; i++) {
        ws.getCell(`A${i}`).value = `A${i}`;
        ws.getCell(`B${i}`).value = `B${i}`;
        ws.getCell(`C${i}`).value = `C${i}`;
    }
    
    // CF auf Spalte A (A2:A10)
    ws.conditionalFormattings = [
        {
            ref: 'A2:A10',
            rules: [{ type: 'expression', formulae: ['TRUE'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } } }]
        },
        {
            ref: 'B2:B10',
            rules: [{ type: 'expression', formulae: ['TRUE'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF00FF00' } } } }]
        }
    ];
    
    console.log('=== VOR spliceColumns ===');
    console.log('CF:', JSON.stringify(ws.conditionalFormattings, null, 2));
    
    // Lösche Spalte A
    ws.spliceColumns(1, 1);
    
    console.log('\n=== NACH spliceColumns ===');
    console.log('CF:', JSON.stringify(ws.conditionalFormattings, null, 2));
    
    // Was steht jetzt in Spalte A?
    console.log('\nSpalte A nach splice:', ws.getCell('A1').value); // Sollte "Spalte B" sein
    console.log('Spalte B nach splice:', ws.getCell('B1').value); // Sollte "Spalte C" sein
    
    // Speichern
    const testPath = '/Users/nojan/Desktop/test-cf-splice.xlsx';
    await workbook.xlsx.writeFile(testPath);
    console.log('\nGespeichert:', testPath);
    
    // Jetzt erneut öffnen und CF prüfen
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(testPath);
    const ws2 = wb2.getWorksheet('Test');
    
    console.log('\n=== NACH RELOAD ===');
    console.log('CF:', JSON.stringify(ws2.conditionalFormattings, null, 2));
}

testCFBehavior().catch(console.error);
