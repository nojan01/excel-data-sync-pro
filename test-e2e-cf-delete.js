// E2E Test: Simuliere den kompletten Speichervorgang und prüfe CF
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Importiere die Writer-Funktionen
const { exportMultipleSheetsWithExcelJS } = require('./exceljs-writer');

const sourceFile = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
const testOutputFile = '/tmp/test-e2e-cf-delete.xlsx';

async function testE2ECFDelete() {
    console.log('=== E2E Test: CF nach Spaltenlöschung ===\n');
    
    // 1. Lade die Originaldatei und prüfe CF
    console.log('1. Lade Originaldatei...');
    const wb1 = new ExcelJS.Workbook();
    await wb1.xlsx.readFile(sourceFile);
    const ws1 = wb1.getWorksheet(1);
    console.log('   Original CF[0].ref:', ws1.conditionalFormattings[0]?.ref);
    
    // 2. Simuliere Frontend-Daten wie sie beim Speichern gesendet werden
    console.log('\n2. Simuliere Frontend-Daten...');
    
    // Headers (ohne die gelöschte Spalte E)
    const originalHeaders = [];
    for (let col = 1; col <= ws1.columnCount; col++) {
        const cell = ws1.getCell(1, col);
        originalHeaders.push(cell.value || '');
    }
    
    // Entferne Spalte 5 (Index 4) - simuliert Spaltenlöschung
    const deletedColumnIndex = 4; // 0-basiert, entspricht Spalte E (5)
    const headersAfterDelete = [...originalHeaders];
    headersAfterDelete.splice(deletedColumnIndex, 1);
    
    console.log('   Original Spaltenanzahl:', originalHeaders.length);
    console.log('   Nach Löschen:', headersAfterDelete.length);
    console.log('   Gelöschte Spalte (0-basiert):', deletedColumnIndex);
    
    // Daten vorbereiten (vereinfacht - nur erste 10 Zeilen)
    const dataRows = [];
    for (let row = 2; row <= Math.min(11, ws1.rowCount); row++) {
        const rowData = [];
        for (let col = 1; col <= ws1.columnCount; col++) {
            if (col !== deletedColumnIndex + 1) { // Überspringe gelöschte Spalte
                const cell = ws1.getCell(row, col);
                rowData.push(cell.value || '');
            }
        }
        dataRows.push(rowData);
    }
    
    // 3. Rufe exportMultipleSheetsWithExcelJS auf
    console.log('\n3. Rufe exportMultipleSheetsWithExcelJS auf...');
    
    const sheets = [{
        sheetName: ws1.name,
        headers: headersAfterDelete,
        data: dataRows,
        fullRewrite: true,
        structuralChange: true,
        deletedColumnIndex: deletedColumnIndex,
        cellStyles: {},
        autoFilterRange: 'A1:BH2404' // Angepasster Bereich
    }];
    
    const result = await exportMultipleSheetsWithExcelJS(sourceFile, testOutputFile, sheets, {});
    
    console.log('   Ergebnis:', result.success ? 'Erfolgreich' : 'Fehler: ' + result.error);
    
    // 4. Prüfe die gespeicherte Datei
    console.log('\n4. Prüfe gespeicherte Datei...');
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(testOutputFile);
    const ws2 = wb2.getWorksheet(1);
    
    console.log('   CF-Regeln:', ws2.conditionalFormattings?.length || 0);
    if (ws2.conditionalFormattings && ws2.conditionalFormattings.length > 0) {
        console.log('   CF[0].ref:', ws2.conditionalFormattings[0]?.ref);
        console.log('   CF[1].ref:', ws2.conditionalFormattings[1]?.ref);
        console.log('   CF[2].ref:', ws2.conditionalFormattings[2]?.ref);
    }
    
    // 5. Validierung
    console.log('\n5. Validierung...');
    const originalRef = 'AN2135:AY2404';
    const expectedRef = 'AM2135:AX2404'; // Nach Löschen von Spalte E sollten alle Refs um 1 nach links
    const actualRef = ws2.conditionalFormattings[0]?.ref;
    
    console.log('   Original:  ', originalRef);
    console.log('   Erwartet:  ', expectedRef);
    console.log('   Tatsächlich:', actualRef);
    console.log('   KORREKT:', actualRef === expectedRef ? 'JA ✓' : 'NEIN ✗ - PROBLEM!');
    
    // Aufräumen
    if (fs.existsSync(testOutputFile)) {
        fs.unlinkSync(testOutputFile);
    }
}

testE2ECFDelete().catch(console.error);
