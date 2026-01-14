/**
 * E2E Test: Exakte Simulation des exceljs-writer Verhaltens
 * Testet ob die Kombination aus _rows-Umordnung + Daten-Schreiben funktioniert
 */

const ExcelJS = require('exceljs');

async function test() {
    const inputPath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    console.log('Lade Datei:', inputPath);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputPath);
    
    const ws = workbook.worksheets[0];
    
    // Simuliere Frontend-Daten: Hole die ersten 15 Zeilen
    const headers = [];
    ws.getRow(1).eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber <= 10) headers.push(cell.value);
    });
    
    const data = [];
    for (let row = 2; row <= 15; row++) {
        const rowData = [];
        for (let col = 1; col <= 10; col++) {
            rowData.push(ws.getCell(row, col).value);
        }
        data.push(rowData);
    }
    
    console.log('\n=== ORIGINAL DATEN ===');
    console.log('Zeile 2 (idx 0):', data[0].slice(0, 3));
    console.log('Zeile 10 (idx 8):', data[8].slice(0, 3));
    console.log('\n=== ORIGINAL STYLES ===');
    console.log('Zeile 2, Spalte B Font:', JSON.stringify(ws.getCell(2, 2).font));
    console.log('Zeile 10, Spalte B Font:', JSON.stringify(ws.getCell(10, 2).font));
    
    // Simuliere Row-Move: Tausche Zeile 2 (idx 0) und Zeile 10 (idx 8) im Frontend
    // Nach dem Tausch: data[0] hat jetzt Original-Zeile 10 Daten, data[8] hat Original-Zeile 2 Daten
    const temp = data[0];
    data[0] = data[8];
    data[8] = temp;
    
    // rowMapping: rowMapping[newPos] = originalDataIndex (0-basiert)
    // Nach Tausch: Position 0 hat Original-Zeile 8, Position 8 hat Original-Zeile 0
    const rowMapping = [8, 1, 2, 3, 4, 5, 6, 7, 0, 9, 10, 11, 12, 13]; // Einfacher Tausch
    
    console.log('\n=== NACH FRONTEND-TAUSCH ===');
    console.log('data[0] (sollte Original-Zeile 10 sein):', data[0].slice(0, 3));
    console.log('data[8] (sollte Original-Zeile 2 sein):', data[8].slice(0, 3));
    console.log('rowMapping[0]:', rowMapping[0], '(erwartet: 8 = Original-Zeile 10)');
    console.log('rowMapping[8]:', rowMapping[8], '(erwartet: 0 = Original-Zeile 2)');
    
    // === WRITER-LOGIK START ===
    console.log('\n=== WRITER-LOGIK ===');
    
    // Schritt 1: _rows umordnen
    const headerRow = ws._rows[0];
    const newRows = [headerRow];
    
    for (let newDataIdx = 0; newDataIdx < rowMapping.length; newDataIdx++) {
        const originalDataIdx = rowMapping[newDataIdx];
        const originalRowsIdx = originalDataIdx + 1; // +1 weil Header bei 0
        const row = ws._rows[originalRowsIdx];
        
        if (row) {
            row._number = newDataIdx + 2;
            newRows.push(row);
        } else {
            newRows.push(undefined);
        }
    }
    
    ws._rows = newRows;
    
    console.log('_rows umgeordnet');
    console.log('Excel-Zeile 2 (nach Umordnung) Value Spalte B:', ws.getCell(2, 2).value);
    console.log('Excel-Zeile 2 (nach Umordnung) Font Spalte B:', JSON.stringify(ws.getCell(2, 2).font));
    
    // Schritt 2: Daten schreiben (wie im Writer)
    data.forEach((row, rowIndex) => {
        row.forEach((value, colIndex) => {
            const cell = ws.getCell(rowIndex + 2, colIndex + 1);
            cell.value = value === null || value === undefined ? '' : value;
        });
    });
    
    console.log('\nNach Daten-Schreiben:');
    console.log('Excel-Zeile 2 Value Spalte B:', ws.getCell(2, 2).value);
    console.log('Excel-Zeile 2 Font Spalte B:', JSON.stringify(ws.getCell(2, 2).font));
    console.log('Excel-Zeile 10 Value Spalte B:', ws.getCell(10, 2).value);
    console.log('Excel-Zeile 10 Font Spalte B:', JSON.stringify(ws.getCell(10, 2).font));
    
    // Speichern
    const outputPath = '/tmp/test-e2e-row-move.xlsx';
    await workbook.xlsx.writeFile(outputPath);
    console.log('\nGespeichert:', outputPath);
    
    // Reload und prüfen
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outputPath);
    const ws2 = wb2.worksheets[0];
    
    console.log('\n=== NACH RELOAD ===');
    console.log('Zeile 2, Spalte A:', ws2.getCell(2, 1).value);
    console.log('Zeile 2, Spalte B Value:', ws2.getCell(2, 2).value);
    console.log('Zeile 2, Spalte B Font:', JSON.stringify(ws2.getCell(2, 2).font));
    console.log('\nZeile 10, Spalte A:', ws2.getCell(10, 1).value);
    console.log('Zeile 10, Spalte B Value:', ws2.getCell(10, 2).value);
    console.log('Zeile 10, Spalte B Font:', JSON.stringify(ws2.getCell(10, 2).font));
    
    // VERIFIKATION
    console.log('\n=== VERIFIKATION ===');
    const font2 = ws2.getCell(2, 2).font;
    const font10 = ws2.getCell(10, 2).font;
    
    // Nach dem Tausch sollte:
    // - Zeile 2 den Wert UND Style von Original-Zeile 10 haben (bold, italic, blue)
    // - Zeile 10 den Wert UND Style von Original-Zeile 2 haben (normal)
    
    const isZeile2BoldItalic = font2 && font2.bold === true && font2.italic === true;
    const isZeile10Normal = !font10 || (!font10.bold && !font10.italic);
    
    if (isZeile2BoldItalic && isZeile10Normal) {
        console.log('✅ ERFOLG: Styles wurden korrekt mit den Zeilen verschoben!');
        console.log('  Zeile 2 hat jetzt bold/italic (wie Original-Zeile 10)');
        console.log('  Zeile 10 ist jetzt normal (wie Original-Zeile 2)');
    } else {
        console.log('❌ FEHLER: Styles nicht korrekt verschoben');
        console.log('  Zeile 2 bold/italic:', isZeile2BoldItalic);
        console.log('  Zeile 10 normal:', isZeile10Normal);
    }
}

test().catch(console.error);
