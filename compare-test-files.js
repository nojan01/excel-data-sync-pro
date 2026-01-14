const ExcelJS = require('exceljs');

async function compare() {
    const wb1 = new ExcelJS.Workbook();
    const wb2 = new ExcelJS.Workbook();
    
    await wb1.xlsx.readFile('/Users/nojan/Desktop/test-styles-exceljs.xlsx');
    await wb2.xlsx.readFile('/Users/nojan/Desktop/Export_test-styles-exceljs.xlsx');
    
    const ws1 = wb1.worksheets[0];
    const ws2 = wb2.worksheets[0];
    
    console.log('=== DATEI-VERGLEICH ===');
    console.log('Original:', ws1.name, '- Zeilen:', ws1.rowCount, '- Spalten:', ws1.columnCount);
    console.log('Export:', ws2.name, '- Zeilen:', ws2.rowCount, '- Spalten:', ws2.columnCount);
    
    // Prüfe Merged Cells
    console.log('\n=== MERGED CELLS ===');
    console.log('Original:', ws1.model.merges || []);
    console.log('Export:', ws2.model.merges || []);
    console.log('\nErwartet nach Spalte A löschen:');
    console.log('  A1:H1 -> A1:G1');
    console.log('  G20:I22 -> F20:H22');
    console.log('  A22:C22 -> A22:B22');
    console.log('  D22:F22 -> C22:E22');
    
    // Spalte A wurde gelöscht, also Original B = Export A
    console.log('\n=== FILL-VERGLEICH (Zeile 4) ===');
    console.log('Original sollte: A=Rot, B=Grün, C=Blau, D=Gelb...');
    console.log('Export sollte: A=Grün (war B), B=Blau (war C), etc.');
    
    // Original row 4
    console.log('\n--- ORIGINAL ---');
    for (let col = 1; col <= 9; col++) {
        const cell1 = ws1.getCell(4, col);
        const fill1 = cell1.fill;
        let color1 = '-';
        if (fill1 && fill1.type === 'pattern' && fill1.fgColor) {
            color1 = fill1.fgColor.argb || fill1.fgColor.theme || 'x';
        }
        console.log('  Orig Spalte ' + col + ': Fill=' + color1 + ', Wert=' + cell1.value);
    }
    
    // Export row 4
    console.log('\n--- EXPORT (nach Spalte A löschen) ---');
    for (let col = 1; col <= 8; col++) {
        const cell2 = ws2.getCell(4, col);
        const fill2 = cell2.fill;
        let color2 = '-';
        if (fill2 && fill2.type === 'pattern' && fill2.fgColor) {
            color2 = fill2.fgColor.argb || fill2.fgColor.theme || 'x';
        }
        console.log('  Export Spalte ' + col + ': Fill=' + color2 + ', Wert=' + cell2.value);
    }
    
    // Detaillierter Vergleich - Original B vs Export A
    console.log('\n=== KRITISCHER VERGLEICH ===');
    console.log('Original Spalte N+1 sollte = Export Spalte N (da A gelöscht):');
    
    let errors = [];
    let correct = 0;
    for (let row = 1; row <= 28; row++) {
        for (let col = 2; col <= 9; col++) { // Original ab Spalte 2
            const origCell = ws1.getCell(row, col);
            const exportCell = ws2.getCell(row, col - 1); // Export ist um 1 verschoben
            
            const origFill = origCell.fill;
            const exportFill = exportCell.fill;
            
            let origColor = null;
            let exportColor = null;
            
            if (origFill && origFill.type === 'pattern' && origFill.fgColor) {
                origColor = origFill.fgColor.argb || String(origFill.fgColor.theme);
            }
            if (exportFill && exportFill.type === 'pattern' && exportFill.fgColor) {
                exportColor = exportFill.fgColor.argb || String(exportFill.fgColor.theme);
            }
            
            if (origColor !== exportColor) {
                // Prüfe ob beide leer sind (das ist OK)
                if (origColor === null && exportColor === null) {
                    correct++;
                } else {
                    errors.push({
                        row: row,
                        origCol: col,
                        exportCol: col - 1,
                        origColor: origColor,
                        exportColor: exportColor,
                        origValue: origCell.value,
                        exportValue: exportCell.value
                    });
                }
            } else {
                correct++;
            }
        }
    }
    
    console.log('Korrekt:', correct, '- Fehler:', errors.length);
    
    if (errors.length > 0) {
        console.log('\nFEHLER (erste 20):');
        errors.slice(0, 20).forEach(e => {
            console.log('  Zeile ' + e.row + ': Orig Spalte ' + e.origCol + ' (' + e.origColor + ') != Export Spalte ' + e.exportCol + ' (' + e.exportColor + ')');
            console.log('    Werte: "' + e.origValue + '" vs "' + e.exportValue + '"');
        });
    } else {
        console.log('\n✅ ALLE FILLS KORREKT!');
    }
}

compare().catch(e => console.error(e));
