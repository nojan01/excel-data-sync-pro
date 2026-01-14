/**
 * Test: Row Move mit SICHTBAREN Styles
 * Findet Zellen mit echten Fills und testet ob diese korrekt verschoben werden
 */

const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

async function test() {
    // Pfad zur Originaldatei
    const inputPath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    console.log('Lade Datei:', inputPath);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputPath);
    
    const ws = workbook.worksheets[0];
    console.log('Worksheet:', ws.name);
    console.log('Zeilen:', ws.rowCount);
    
    // SCHRITT 1: Finde Zellen mit echten Fills (nicht "none")
    console.log('\n=== SUCHE NACH ZELLEN MIT FILLS ===');
    
    const styledCells = [];
    for (let rowNum = 2; rowNum <= Math.min(50, ws.rowCount); rowNum++) {
        const row = ws.getRow(rowNum);
        for (let colNum = 1; colNum <= Math.min(10, ws.columnCount); colNum++) {
            const cell = row.getCell(colNum);
            if (cell.fill && cell.fill.pattern && cell.fill.pattern !== 'none') {
                styledCells.push({
                    row: rowNum,
                    col: colNum,
                    value: cell.value,
                    fill: cell.fill
                });
            }
        }
    }
    
    console.log('Gefundene Zellen mit Fills:', styledCells.length);
    
    if (styledCells.length === 0) {
        console.log('\nKeine Zellen mit Fills in den ersten 50 Zeilen gefunden.');
        console.log('Suche nach anderen Styles (Font)...');
        
        // Alternative: Suche nach Font-Styles
        for (let rowNum = 2; rowNum <= Math.min(50, ws.rowCount); rowNum++) {
            const row = ws.getRow(rowNum);
            for (let colNum = 1; colNum <= Math.min(10, ws.columnCount); colNum++) {
                const cell = row.getCell(colNum);
                if (cell.font && (cell.font.bold || cell.font.color)) {
                    console.log(`  Row ${rowNum}, Col ${colNum}: Font =`, JSON.stringify(cell.font));
                    if (styledCells.length < 5) {
                        styledCells.push({
                            row: rowNum,
                            col: colNum,
                            value: cell.value,
                            font: cell.font,
                            fill: cell.fill
                        });
                    }
                }
            }
        }
    }
    
    // Zeige gefundene Styles
    styledCells.slice(0, 10).forEach(c => {
        console.log(`  Row ${c.row}, Col ${c.col}: Value="${c.value}"`, 
                    c.fill ? `Fill=${JSON.stringify(c.fill)}` : '',
                    c.font ? `Font=${JSON.stringify(c.font)}` : '');
    });
    
    // SCHRITT 2: Simuliere Row-Move
    console.log('\n=== ROW MOVE TEST ===');
    
    // Tausche Zeile 2 und Zeile 3 (oder erste zwei gefundene styled rows)
    const rowA = 2;
    const rowB = 3;
    
    console.log(`Vor Tausch:`);
    const cellA_before = ws.getCell(rowA, 1);
    const cellB_before = ws.getCell(rowB, 1);
    console.log(`  Zeile ${rowA}: Value="${cellA_before.value}", Fill=${JSON.stringify(cellA_before.fill)}, Font=${JSON.stringify(cellA_before.font)}`);
    console.log(`  Zeile ${rowB}: Value="${cellB_before.value}", Fill=${JSON.stringify(cellB_before.fill)}, Font=${JSON.stringify(cellB_before.font)}`);
    
    // Row-Objekte direkt tauschen
    const rowObjA = ws._rows[rowA - 1]; // 0-based
    const rowObjB = ws._rows[rowB - 1];
    
    if (!rowObjA || !rowObjB) {
        console.log('FEHLER: Row-Objekte nicht gefunden');
        return;
    }
    
    // Tausche
    ws._rows[rowA - 1] = rowObjB;
    ws._rows[rowB - 1] = rowObjA;
    
    // Aktualisiere _number
    ws._rows[rowA - 1]._number = rowA;
    ws._rows[rowB - 1]._number = rowB;
    
    console.log(`\nNach Tausch (nur Row-Objekte getauscht):`);
    const cellA_after = ws.getCell(rowA, 1);
    const cellB_after = ws.getCell(rowB, 1);
    console.log(`  Zeile ${rowA}: Value="${cellA_after.value}", Fill=${JSON.stringify(cellA_after.fill)}, Font=${JSON.stringify(cellA_after.font)}`);
    console.log(`  Zeile ${rowB}: Value="${cellB_after.value}", Fill=${JSON.stringify(cellB_after.fill)}, Font=${JSON.stringify(cellB_after.font)}`);
    
    // SCHRITT 3: Speichern und prüfen
    const outputPath = '/tmp/test-row-styled.xlsx';
    await workbook.xlsx.writeFile(outputPath);
    console.log('\nGespeichert:', outputPath);
    
    // Neu laden und prüfen
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outputPath);
    const ws2 = wb2.worksheets[0];
    
    console.log('\n=== NACH RELOAD ===');
    const cellA_reload = ws2.getCell(rowA, 1);
    const cellB_reload = ws2.getCell(rowB, 1);
    console.log(`  Zeile ${rowA}: Value="${cellA_reload.value}", Fill=${JSON.stringify(cellA_reload.fill)}, Font=${JSON.stringify(cellA_reload.font)}`);
    console.log(`  Zeile ${rowB}: Value="${cellB_reload.value}", Fill=${JSON.stringify(cellB_reload.fill)}, Font=${JSON.stringify(cellB_reload.font)}`);
    
    // VERIFIKATION
    console.log('\n=== VERIFIKATION ===');
    const valueSwapped = cellA_reload.value === cellB_before.value && cellB_reload.value === cellA_before.value;
    const styleSwapped = JSON.stringify(cellA_reload.fill) === JSON.stringify(cellB_before.fill);
    
    console.log(`Values getauscht: ${valueSwapped ? '✓ JA' : '✗ NEIN'}`);
    console.log(`  Erwartet: Zeile ${rowA} hat jetzt "${cellB_before.value}", Zeile ${rowB} hat "${cellA_before.value}"`);
    console.log(`  Tatsächlich: Zeile ${rowA} hat "${cellA_reload.value}", Zeile ${rowB} hat "${cellB_reload.value}"`);
    
    console.log(`\nStyles getauscht: ${styleSwapped ? '✓ JA' : '✗ NEIN'}`);
    console.log(`  Erwartet: Zeile ${rowA} hat Fill von vorher Zeile ${rowB}`);
}

test().catch(console.error);
