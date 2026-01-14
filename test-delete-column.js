// Test: Spalte löschen beim Speichern
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function testDeleteColumn() {
    console.log('=== Test: Spalte löschen ===\n');
    
    const testFile = path.join(__dirname, 'test-delete-col.xlsx');
    const outputFile = path.join(__dirname, 'test-delete-col-saved.xlsx');
    
    // 1. Erstelle Test-Datei mit 5 Spalten
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Test');
    
    ws.getRow(1).values = ['Spalte A', 'Spalte B', 'Spalte C', 'Spalte D', 'Spalte E'];
    for (let i = 2; i <= 5; i++) {
        ws.getRow(i).values = [`A${i}`, `B${i}`, `C${i}`, `D${i}`, `E${i}`];
    }
    
    await wb.xlsx.writeFile(testFile);
    console.log('✓ Test-Datei erstellt mit 5 Spalten');
    
    // 2. Lade Datei und simuliere "Spalte B löschen"
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(testFile);
    const ws2 = wb2.getWorksheet('Test');
    
    console.log('\n--- VOR dem Löschen ---');
    console.log('Spaltenanzahl:', ws2.columnCount);
    console.log('Zeile 1:', ws2.getRow(1).values);
    
    // Simuliere was unser Frontend macht: neue Header und Daten ohne Spalte B
    const newHeaders = ['Spalte A', 'Spalte C', 'Spalte D', 'Spalte E'];
    const newData = [
        ['A2', 'C2', 'D2', 'E2'],
        ['A3', 'C3', 'D3', 'E3'],
        ['A4', 'C4', 'D4', 'E4'],
        ['A5', 'C5', 'D5', 'E5']
    ];
    
    // Header überschreiben
    newHeaders.forEach((header, colIndex) => {
        ws2.getCell(1, colIndex + 1).value = header;
    });
    
    // Daten überschreiben
    newData.forEach((row, rowIndex) => {
        row.forEach((value, colIndex) => {
            ws2.getCell(rowIndex + 2, colIndex + 1).value = value;
        });
    });
    
    // Überschüssige Spalten leeren
    const originalColumnCount = ws2.columnCount;
    console.log('\nOriginal Spaltenanzahl:', originalColumnCount);
    console.log('Neue Spaltenanzahl:', newHeaders.length);
    
    if (originalColumnCount > newHeaders.length) {
        const rowCount = ws2.rowCount;
        console.log('Lösche Spalten', newHeaders.length + 1, 'bis', originalColumnCount);
        
        for (let rowIdx = 1; rowIdx <= rowCount; rowIdx++) {
            for (let colIdx = newHeaders.length + 1; colIdx <= originalColumnCount; colIdx++) {
                const cell = ws2.getCell(rowIdx, colIdx);
                cell.value = null;
                cell.style = {};
            }
        }
    }
    
    await wb2.xlsx.writeFile(outputFile);
    console.log('\n✓ Datei gespeichert');
    
    // 3. Verifiziere
    const wb3 = new ExcelJS.Workbook();
    await wb3.xlsx.readFile(outputFile);
    const ws3 = wb3.getWorksheet('Test');
    
    console.log('\n--- NACH dem Speichern ---');
    console.log('Spaltenanzahl:', ws3.columnCount);
    console.log('Zeile 1:', ws3.getRow(1).values);
    console.log('Zeile 2:', ws3.getRow(2).values);
    
    // Prüfe ob Spalte 5 wirklich leer ist
    console.log('\nSpalte 5 Werte:');
    for (let i = 1; i <= 5; i++) {
        const cell = ws3.getCell(i, 5);
        console.log(`  Zeile ${i}: "${cell.value}"`);
    }
    
    // Aufräumen
    fs.unlinkSync(testFile);
    fs.unlinkSync(outputFile);
    
    console.log('\n=== Test Ende ===');
}

testDeleteColumn().catch(console.error);
