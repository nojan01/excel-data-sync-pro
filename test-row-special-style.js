/**
 * Test: Row Move mit Row 10 (hat speziellen Font-Style)
 * Row 10, Col 2 hat: bold, italic, blue color, underline
 */

const ExcelJS = require('exceljs');

async function test() {
    const inputPath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    console.log('Lade Datei:', inputPath);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputPath);
    
    const ws = workbook.worksheets[0];
    
    // Zeile 10 hat den speziellen Style (Row 10, Col 2: bold, italic, blue, underline)
    const specialRow = 10;
    const normalRow = 2;
    
    console.log('\n=== VORHER ===');
    console.log(`Zeile ${normalRow}, Spalte B:`);
    const cellNormal = ws.getCell(normalRow, 2);
    console.log(`  Value: "${cellNormal.value}"`);
    console.log(`  Font: ${JSON.stringify(cellNormal.font)}`);
    
    console.log(`\nZeile ${specialRow}, Spalte B:`);
    const cellSpecial = ws.getCell(specialRow, 2);
    console.log(`  Value: "${cellSpecial.value}"`);
    console.log(`  Font: ${JSON.stringify(cellSpecial.font)}`);
    
    // Speichere die Font-Infos für Vergleich
    const fontNormalBefore = JSON.stringify(cellNormal.font);
    const fontSpecialBefore = JSON.stringify(cellSpecial.font);
    const valueNormalBefore = cellNormal.value;
    const valueSpecialBefore = cellSpecial.value;
    
    // TAUSCHE Row 2 und Row 10 via _rows
    console.log('\n=== TAUSCHE ZEILEN ===');
    const rowObj2 = ws._rows[normalRow - 1];
    const rowObj10 = ws._rows[specialRow - 1];
    
    ws._rows[normalRow - 1] = rowObj10;
    ws._rows[specialRow - 1] = rowObj2;
    
    ws._rows[normalRow - 1]._number = normalRow;
    ws._rows[specialRow - 1]._number = specialRow;
    
    // Speichern
    const outputPath = '/tmp/test-row-special-style.xlsx';
    await workbook.xlsx.writeFile(outputPath);
    console.log('Gespeichert:', outputPath);
    
    // Neu laden und prüfen
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outputPath);
    const ws2 = wb2.worksheets[0];
    
    console.log('\n=== NACHHER (nach Reload) ===');
    console.log(`Zeile ${normalRow}, Spalte B:`);
    const cellNormalAfter = ws2.getCell(normalRow, 2);
    console.log(`  Value: "${cellNormalAfter.value}"`);
    console.log(`  Font: ${JSON.stringify(cellNormalAfter.font)}`);
    
    console.log(`\nZeile ${specialRow}, Spalte B:`);
    const cellSpecialAfter = ws2.getCell(specialRow, 2);
    console.log(`  Value: "${cellSpecialAfter.value}"`);
    console.log(`  Font: ${JSON.stringify(cellSpecialAfter.font)}`);
    
    // VERIFIKATION
    console.log('\n=== VERIFIKATION ===');
    
    // Zeile 2 sollte jetzt den Wert und Font von Zeile 10 haben
    const valueSwapped = 
        cellNormalAfter.value === valueSpecialBefore &&
        cellSpecialAfter.value === valueNormalBefore;
    
    // Font-Vergleich (ignoriere Reihenfolge der Properties)
    const font2After = cellNormalAfter.font;
    const font10After = cellSpecialAfter.font;
    
    const styleSwapped = 
        font2After && font2After.bold === true && font2After.italic === true &&
        (!font10After || (!font10After.bold && !font10After.italic));
    
    console.log(`\nWerte korrekt getauscht: ${valueSwapped ? '✓ JA' : '✗ NEIN'}`);
    console.log(`  Zeile ${normalRow} sollte "${valueSpecialBefore}" haben, hat: "${cellNormalAfter.value}"`);
    console.log(`  Zeile ${specialRow} sollte "${valueNormalBefore}" haben, hat: "${cellSpecialAfter.value}"`);
    
    console.log(`\nStyles korrekt getauscht: ${styleSwapped ? '✓ JA' : '✗ NEIN'}`);
    console.log(`  Zeile ${normalRow} sollte bold/italic haben: ${font2After?.bold ? 'JA' : 'NEIN'}`);
    console.log(`  Zeile ${specialRow} sollte normal sein: ${!font10After?.bold ? 'JA' : 'NEIN'}`);
    
    console.log('\n=== FAZIT ===');
    if (valueSwapped && styleSwapped) {
        console.log('✅ Row-Swap funktioniert korrekt! Werte UND Styles werden getauscht.');
    } else {
        console.log('❌ Problem: ' + (!valueSwapped ? 'Werte nicht getauscht. ' : '') + (!styleSwapped ? 'Styles nicht getauscht.' : ''));
    }
}

test().catch(console.error);
