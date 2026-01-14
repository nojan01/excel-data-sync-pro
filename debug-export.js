// Debug: Was ist WIRKLICH in den cellStyles vom Frontend?
// Wir müssen die Export-Datei analysieren, um zu verstehen was passiert ist

const ExcelJS = require('exceljs');

async function debug() {
    const exportPath = '/Users/nojan/Desktop/Export_test-styles-exceljs.xlsx';
    const originalPath = '/Users/nojan/Desktop/test-styles-exceljs.xlsx';
    
    const wb1 = new ExcelJS.Workbook();
    const wb2 = new ExcelJS.Workbook();
    
    await wb1.xlsx.readFile(originalPath);
    await wb2.xlsx.readFile(exportPath);
    
    const ws1 = wb1.worksheets[0];
    const ws2 = wb2.worksheets[0];
    
    console.log('=== ZEILE 4 ANALYSE ===');
    console.log('\n--- ORIGINAL (Zeile 4) ---');
    
    // Original Zeile 4, alle Spalten
    for (let col = 1; col <= 9; col++) {
        const cell = ws1.getCell(4, col);
        const fill = cell.fill;
        let fillStr = 'keine';
        if (fill && fill.type === 'pattern' && fill.fgColor) {
            fillStr = fill.fgColor.argb || fill.fgColor.theme || JSON.stringify(fill.fgColor);
        }
        console.log('  Spalte ' + col + ' (Idx ' + (col-1) + '): Wert="' + cell.value + '", Fill=' + fillStr);
    }
    
    console.log('\n--- EXPORT (Zeile 4) ---');
    
    // Export Zeile 4, alle Spalten
    for (let col = 1; col <= 8; col++) {
        const cell = ws2.getCell(4, col);
        const fill = cell.fill;
        let fillStr = 'keine';
        if (fill && fill.type === 'pattern' && fill.fgColor) {
            fillStr = fill.fgColor.argb || fill.fgColor.theme || JSON.stringify(fill.fgColor);
        }
        console.log('  Spalte ' + col + ' (Idx ' + (col-1) + '): Wert="' + cell.value + '", Fill=' + fillStr);
    }
    
    // Zeige das erwartete Mapping
    console.log('\n=== ERWARTETES MAPPING (nach Spalte A löschen) ===');
    console.log('Original Spalte 2 (B) -> Export Spalte 1 (A)');
    console.log('Original Spalte 3 (C) -> Export Spalte 2 (B)');
    console.log('... usw.');
    
    // Prüfe ob die Werte korrekt sind (Werte werden von spliceColumns korrekt verschoben)
    console.log('\n=== WERT-PRÜFUNG ===');
    for (let origCol = 2; origCol <= 9; origCol++) {
        const expCol = origCol - 1;
        const origCell = ws1.getCell(4, origCol);
        const expCell = ws2.getCell(4, expCol);
        
        const match = origCell.value === expCell.value ? '✓' : '✗';
        console.log('  Orig Spalte ' + origCol + ' ("' + origCell.value + '") -> Export Spalte ' + expCol + ' ("' + expCell.value + '") ' + match);
    }
    
    // Prüfe Fills 
    console.log('\n=== FILL-PRÜFUNG (DAS PROBLEM) ===');
    console.log('spliceColumns verschiebt die Zellen, aber applyMissingFills überschreibt sie dann');
    console.log('');
    
    // Was hat ExcelJS nach spliceColumns für Fills?
    // Das wissen wir nicht direkt, aber wir können die Export-Datei analysieren
    
    // Die Frage ist: Was steht in cellStyles[3-1]?
    // cellStyles kommt vom Frontend und enthält die ORIGINAL-Fills
    // Wenn cellStyles[3-1] = Grün, dann sollte adjustedCellStyles[3-0] = Grün
    // Aber applyMissingFills setzt cell(4, 1) = Grün, was KORREKT ist!
    
    // ABER: Der Test zeigt cell(4, 1) = BLAU, nicht Grün!
    // Das bedeutet entweder:
    // 1. cellStyles[3-1] enthält Blau, nicht Grün
    // 2. Oder es gibt einen anderen Bug
    
    console.log('Die Farben im Export sind:');
    for (let col = 1; col <= 8; col++) {
        const cell = ws2.getCell(4, col);
        const fill = cell.fill;
        let color = 'keine';
        if (fill && fill.type === 'pattern' && fill.fgColor && fill.fgColor.argb) {
            const hex = fill.fgColor.argb.substring(2);
            if (hex === '0000FF') color = 'BLAU';
            else if (hex === '00FF00') color = 'GRÜN';
            else if (hex === 'FF0000') color = 'ROT';
            else if (hex === 'FFFF00') color = 'GELB';
            else if (hex === 'FFA500') color = 'ORANGE';
            else if (hex === '800080') color = 'LILA';
            else if (hex === '00FFFF') color = 'CYAN';
            else if (hex === 'FF69B4') color = 'PINK';
            else color = hex;
        }
        console.log('  Export Spalte ' + col + ': ' + color);
    }
    
    console.log('\nDas SOLLTE sein (Original B->A, C->B, usw.):');
    console.log('  Export Spalte 1: GRÜN (war Original B)');
    console.log('  Export Spalte 2: BLAU (war Original C)');
    console.log('  Export Spalte 3: GELB (war Original D)');
    console.log('  Export Spalte 4: ORANGE (war Original E)');
    console.log('  Export Spalte 5: LILA (war Original F)');
    console.log('  Export Spalte 6: CYAN (war Original G)');
    console.log('  Export Spalte 7: PINK (war Original H)');
    console.log('  Export Spalte 8: keine (war Original I)');
    
    console.log('\n=== HYPOTHESE ===');
    console.log('Wenn Export Spalte 1 = BLAU, dann wurde cellStyles[3-2] (Original C, Blau) zu [3-1] verschoben');
    console.log('Aber cellStyles[3-1] (Original B, Grün) wurde zu [3-0] verschoben');
    console.log('');
    console.log('FEHLER: applyMissingFills überschreibt existierende Fills NICHT!');
    console.log('Aber spliceColumns HAT die Fills bereits korrekt verschoben!');
    console.log('Dann fügt applyMissingFills ZUSÄTZLICHE Fills hinzu an den FALSCHEN Stellen!');
}

debug().catch(e => console.error(e));
