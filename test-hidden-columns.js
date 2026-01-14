// Test: Versteckte Spalten beim Speichern
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function testHiddenColumns() {
    console.log('=== Test: Versteckte Spalten ===\n');
    
    // 1. Erstelle eine Test-Datei mit versteckten Spalten
    const testFile = path.join(__dirname, 'test-hidden-cols.xlsx');
    const outputFile = path.join(__dirname, 'test-hidden-cols-saved.xlsx');
    
    // Erstelle Workbook mit versteckten Spalten
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Test');
    
    // Header hinzufügen
    ws.getRow(1).values = ['Spalte A', 'Spalte B (versteckt)', 'Spalte C', 'Spalte D (versteckt)', 'Spalte E'];
    
    // Daten hinzufügen
    for (let i = 2; i <= 10; i++) {
        ws.getRow(i).values = [`A${i}`, `B${i}`, `C${i}`, `D${i}`, `E${i}`];
    }
    
    // Spalten B und D verstecken (Index 2 und 4)
    ws.getColumn(2).hidden = true;
    ws.getColumn(4).hidden = true;
    
    await wb.xlsx.writeFile(testFile);
    console.log('✓ Test-Datei erstellt:', testFile);
    
    // 2. Prüfe ob Spalten versteckt sind
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(testFile);
    const ws2 = wb2.getWorksheet('Test');
    
    console.log('\n--- Spalten-Status nach Erstellung ---');
    for (let i = 1; i <= 5; i++) {
        const col = ws2.getColumn(i);
        console.log(`Spalte ${i}: hidden=${col.hidden}`);
    }
    
    // 3. Simuliere das Speichern mit changedCells (wie unser Tool es macht)
    // Ändere eine Zelle
    ws2.getCell(2, 1).value = 'GEÄNDERT';
    
    // Jetzt die versteckten Spalten setzen wie unser Code es macht
    const hiddenColumns = [1, 3]; // Spalte B (Index 1) und D (Index 3) - 0-basiert
    const hiddenSet = new Set(hiddenColumns);
    const columnCount = ws2.columnCount;
    
    console.log('\n--- Setze Spaltensichtbarkeit ---');
    console.log('hiddenColumns (0-basiert):', hiddenColumns);
    console.log('columnCount:', columnCount);
    
    for (let colIdx = 0; colIdx < columnCount; colIdx++) {
        const col = ws2.getColumn(colIdx + 1);
        const shouldBeHidden = hiddenSet.has(colIdx);
        console.log(`Spalte ${colIdx + 1}: vorher hidden=${col.hidden}, setze auf hidden=${shouldBeHidden}`);
        col.hidden = shouldBeHidden;
    }
    
    await wb2.xlsx.writeFile(outputFile);
    console.log('\n✓ Datei gespeichert:', outputFile);
    
    // 4. Verifiziere das Ergebnis
    const wb3 = new ExcelJS.Workbook();
    await wb3.xlsx.readFile(outputFile);
    const ws3 = wb3.getWorksheet('Test');
    
    console.log('\n--- Spalten-Status nach Speichern ---');
    let success = true;
    for (let i = 1; i <= 5; i++) {
        const col = ws3.getColumn(i);
        const expected = (i === 2 || i === 4); // Spalten B und D (1-basiert: 2 und 4)
        const status = col.hidden === expected ? '✓' : '✗';
        if (col.hidden !== expected) success = false;
        console.log(`${status} Spalte ${i}: hidden=${col.hidden} (erwartet: ${expected})`);
    }
    
    // Aufräumen
    fs.unlinkSync(testFile);
    fs.unlinkSync(outputFile);
    
    console.log('\n' + (success ? '✓ TEST BESTANDEN' : '✗ TEST FEHLGESCHLAGEN'));
}

testHiddenColumns().catch(console.error);
