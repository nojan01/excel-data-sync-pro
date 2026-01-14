/**
 * VOLLSTÄNDIGER E2E TEST
 * Simuliert exakt was die App macht:
 * 1. Lade Excel
 * 2. Simuliere Frontend Row-Move (Zeile 2 ↔ Zeile 10)
 * 3. Simuliere Export mit rowMapping
 * 4. Prüfe ob Styles korrekt verschoben wurden
 */

const ExcelJS = require('exceljs');

async function test() {
    const inputPath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    console.log('=== VOLLSTÄNDIGER E2E TEST ===\n');
    console.log('Simuliert: Row-Move von Zeile 10 nach Position 1 (vor Zeile 2)');
    console.log('Erwartet: Zeile 10 wird zu neuer Zeile 2, alte Zeile 2 wird zu Zeile 3\n');
    
    // SCHRITT 1: Lade Excel
    console.log('1. Lade Excel...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputPath);
    const ws = workbook.worksheets[0];
    
    // Simuliere Frontend-Daten
    const frontendData = [];
    for (let row = 2; row <= 15; row++) {
        const rowData = [];
        for (let col = 1; col <= 10; col++) {
            rowData.push(ws.getCell(row, col).value);
        }
        frontendData.push(rowData);
    }
    
    console.log('   Original frontendData[0] (Zeile 2):', frontendData[0].slice(0, 3));
    console.log('   Original frontendData[8] (Zeile 10):', frontendData[8].slice(0, 3));
    console.log('   Original Zeile 10, Spalte B Font:', JSON.stringify(ws.getCell(10, 2).font).substring(0, 80));
    
    // SCHRITT 2: Simuliere Frontend Row-Move
    // Verschiebe Zeile 10 (Index 8) an Position 1 (vor Zeile 2)
    console.log('\n2. Simuliere Frontend Row-Move...');
    
    // rowMapping startet als [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
    // Nach Row-Move: Zeile 10 (Index 8) kommt an Position 0
    // Neue Reihenfolge: [8, 0, 1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13]
    const rowMapping = [8, 0, 1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13];
    
    // Frontend-Daten auch in neue Reihenfolge bringen
    const movedRow = frontendData.splice(8, 1)[0]; // Entferne Index 8
    frontendData.splice(0, 0, movedRow); // Füge an Position 0 ein
    
    console.log('   rowMapping:', rowMapping.slice(0, 10), '...');
    console.log('   Nach Move frontendData[0] (sollte Original-Zeile 10 sein):', frontendData[0].slice(0, 3));
    console.log('   Nach Move frontendData[1] (sollte Original-Zeile 2 sein):', frontendData[1].slice(0, 3));
    
    // SCHRITT 3: Simuliere Writer-Logik
    console.log('\n3. Simuliere exceljs-writer Logik...');
    
    // 3a: Zeilen umordnen (wie im Writer)
    const headerRow = ws._rows[0];
    const newRows = [headerRow];
    
    for (let newDataIdx = 0; newDataIdx < rowMapping.length; newDataIdx++) {
        const originalDataIdx = rowMapping[newDataIdx];
        const originalRowsIdx = originalDataIdx + 1;
        const row = ws._rows[originalRowsIdx];
        
        if (row) {
            row._number = newDataIdx + 2;
            newRows.push(row);
        } else {
            newRows.push(undefined);
        }
    }
    
    ws._rows = newRows;
    
    console.log('   _rows umgeordnet');
    console.log('   Zeile 2 nach Umordnung - Value:', ws.getCell(2, 2).value);
    console.log('   Zeile 2 nach Umordnung - Font:', JSON.stringify(ws.getCell(2, 2).font).substring(0, 80));
    
    // 3b: Daten schreiben (wie im Writer)
    frontendData.forEach((row, rowIndex) => {
        row.forEach((value, colIndex) => {
            const cell = ws.getCell(rowIndex + 2, colIndex + 1);
            cell.value = value === null || value === undefined ? '' : value;
        });
    });
    
    console.log('   Daten geschrieben');
    console.log('   Zeile 2 nach Schreiben - Value:', ws.getCell(2, 2).value);
    console.log('   Zeile 2 nach Schreiben - Font:', JSON.stringify(ws.getCell(2, 2).font).substring(0, 80));
    
    // SCHRITT 4: Speichern und neu laden
    console.log('\n4. Speichere und lade neu...');
    const outputPath = '/tmp/test-full-e2e-rowmove.xlsx';
    await workbook.xlsx.writeFile(outputPath);
    
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outputPath);
    const ws2 = wb2.worksheets[0];
    
    // SCHRITT 5: Verifikation
    console.log('\n5. VERIFIKATION:');
    console.log('   Zeile 2 (sollte Original-Zeile 10 sein):');
    console.log('     Spalte A:', ws2.getCell(2, 1).value, '(erwartet: 468)');
    console.log('     Spalte B:', ws2.getCell(2, 2).value, '(erwartet: CASSIDIAN)');
    console.log('     Font:', JSON.stringify(ws2.getCell(2, 2).font));
    
    console.log('   Zeile 3 (sollte Original-Zeile 2 sein):');
    console.log('     Spalte A:', ws2.getCell(3, 1).value, '(erwartet: 369)');
    console.log('     Spalte B:', ws2.getCell(3, 2).value, '(erwartet: null oder leer)');
    console.log('     Font:', JSON.stringify(ws2.getCell(3, 2).font));
    
    // Finale Bewertung
    const font2 = ws2.getCell(2, 2).font;
    const value2 = ws2.getCell(2, 1).value;
    
    console.log('\n=== ERGEBNIS ===');
    if (font2 && font2.bold === true && value2 === 468) {
        console.log('✅ ERFOLG: Row-Move mit Styles funktioniert korrekt!');
        console.log('   Wert UND Style von Zeile 10 sind jetzt in Zeile 2');
    } else {
        console.log('❌ FEHLER: Etwas stimmt nicht');
        console.log('   Wert korrekt:', value2 === 468);
        console.log('   Style korrekt (bold):', font2?.bold === true);
    }
    
    console.log('\nAusgabedatei:', outputPath);
}

test().catch(console.error);
