// Test: Prüfe bedingte Formatierung nach spliceColumns
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

async function testCFAfterSplice() {
    // Erstelle Test-Workbook mit bedingter Formatierung
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Test');
    
    // Füge einige Daten hinzu
    ws.getCell('A1').value = 'A';
    ws.getCell('B1').value = 'B';
    ws.getCell('C1').value = 'C';
    ws.getCell('D1').value = 'D';
    ws.getCell('E1').value = 'E';
    
    // Füge bedingte Formatierung hinzu
    ws.addConditionalFormatting({
        ref: 'D1:E10',
        rules: [
            {
                type: 'expression',
                formulae: ['$D1>0'],
                style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } }
            }
        ]
    });
    
    console.log('=== VOR spliceColumns ===');
    console.log('Column Count:', ws.columnCount);
    console.log('CF:', JSON.stringify(ws.conditionalFormattings, null, 2));
    
    // Speichere und lade neu, um wie im echten Szenario zu arbeiten
    const tempPath = '/tmp/test-cf-splice.xlsx';
    await wb.xlsx.writeFile(tempPath);
    
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(tempPath);
    const ws2 = wb2.getWorksheet('Test');
    
    console.log('\n=== NACH LADEN ===');
    console.log('Column Count:', ws2.columnCount);
    console.log('CF:', JSON.stringify(ws2.conditionalFormattings, null, 2));
    
    // Spalte C (3) löschen
    console.log('\n=== SPALTE C (3) LÖSCHEN ===');
    ws2.spliceColumns(3, 1);
    
    console.log('\n=== NACH spliceColumns ===');
    console.log('Column Count:', ws2.columnCount);
    console.log('Zellen Zeile 1:', ws2.getCell('A1').value, ws2.getCell('B1').value, ws2.getCell('C1').value, ws2.getCell('D1').value);
    console.log('CF:', JSON.stringify(ws2.conditionalFormattings, null, 2));
    
    // Was ExcelJS macht: Die Zellwerte werden verschoben, aber was passiert mit CF?
    console.log('\n=== ANALYSE ===');
    if (ws2.conditionalFormattings && ws2.conditionalFormattings.length > 0) {
        const cf = ws2.conditionalFormattings[0];
        console.log('CF ref nach splice:', cf.ref);
        console.log('Erwarteter Wert: C1:D10 (verschoben von D1:E10)');
        console.log('IST KORREKT:', cf.ref === 'C1:D10' ? 'JA' : 'NEIN - PROBLEM GEFUNDEN!');
    }
    
    fs.unlinkSync(tempPath);
}

testCFAfterSplice().catch(console.error);
