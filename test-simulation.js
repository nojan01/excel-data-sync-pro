const ExcelJS = require('exceljs');
const path = require('path');

// Hilfsfunktion: Spalten-Buchstabe zu Nummer
function colLetterToNumber(letters) {
    let result = 0;
    for (let i = 0; i < letters.length; i++) {
        result = result * 26 + (letters.charCodeAt(i) - 64);
    }
    return result;
}

// Simuliert den kompletten Flow wie im exceljs-writer.js
async function simulateExport() {
    const originalPath = '/Users/nojan/Desktop/test-styles-exceljs.xlsx';
    const testOutputPath = '/Users/nojan/Desktop/Test_Simulation.xlsx';
    
    console.log('=== SIMULATION DES EXPORTS ===\n');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(originalPath);
    const ws = workbook.getWorksheet(1);
    
    // Simuliere cellStyles wie vom Frontend nach Spalte A löschen
    // Die rosa Fills sind für G22, H22, I22 -> nach Anpassung: F22, G22, H22
    // cellStyles Format: "rowIdx-colIdx" mit 0-basiert
    // F22 = Zeile 22, Spalte 6 -> rowIdx=21, colIdx=5 -> "21-5"
    // G22 = Zeile 22, Spalte 7 -> rowIdx=21, colIdx=6 -> "21-6"
    // H22 = Zeile 22, Spalte 8 -> rowIdx=21, colIdx=7 -> "21-7"
    const cellStyles = {
        '21-5': { fill: '#FFE0E0' },  // F22 - rosa (vom Frontend angepasst von G22)
        '21-6': { fill: '#FFE0E0' },  // G22 - rosa (vom Frontend angepasst von H22)
        '21-7': { fill: '#FFE0E0' },  // H22 - rosa (vom Frontend angepasst von I22)
    };
    
    console.log('cellStyles vor Verarbeitung:', cellStyles);
    
    // 1. Prepare Merged Cells
    console.log('\n=== 1. prepareMergedCellsForColumnDelete ===');
    const oldMerges = [...ws.model.merges];
    console.log('Alte Merges:', oldMerges);
    
    const deletedColExcel = 1; // Spalte A
    const newMerges = [];
    
    oldMerges.forEach(mergeRange => {
        const parts = mergeRange.split(':');
        if (parts.length !== 2) return;
        
        const startMatch = parts[0].match(/^([A-Z]+)(\d+)$/);
        const endMatch = parts[1].match(/^([A-Z]+)(\d+)$/);
        if (!startMatch || !endMatch) return;
        
        const startCol = colLetterToNumber(startMatch[1]);
        const startRow = parseInt(startMatch[2]);
        const endCol = colLetterToNumber(endMatch[1]);
        const endRow = parseInt(endMatch[2]);
        
        // Speichere Master-Wert und Fill
        const masterCell = ws.getCell(startRow, startCol);
        const masterValue = masterCell.value;
        const masterFill = masterCell.fill;
        
        let newStartCol = startCol;
        let newEndCol = endCol;
        
        if (endCol < deletedColExcel) {
            // Merge komplett vor gelöschter Spalte - bleibt gleich
        } else if (startCol > deletedColExcel) {
            // Merge komplett nach gelöschter Spalte - verschiebe um 1
            newStartCol--;
            newEndCol--;
        } else if (startCol < deletedColExcel && endCol >= deletedColExcel) {
            // Gelöschte Spalte innerhalb des Merge - Ende verringern
            newEndCol--;
        } else if (startCol === deletedColExcel) {
            // Merge beginnt bei gelöschter Spalte - Start bleibt, Ende verringern
            newEndCol--;
        }
        
        if (newEndCol < newStartCol) return;
        
        const newStartLetter = String.fromCharCode(64 + newStartCol);
        const newEndLetter = String.fromCharCode(64 + newEndCol);
        const newRange = newStartLetter + startRow + ':' + newEndLetter + endRow;
        
        newMerges.push({ range: newRange, value: masterValue, fill: masterFill });
        
        // Unmerge
        try { ws.unMergeCells(mergeRange); } catch(e) {}
    });
    
    console.log('Neue Merges vorbereitet:', newMerges.map(m => m.range));
    
    // 2. spliceColumns
    console.log('\n=== 2. spliceColumns(1, 1) ===');
    ws.spliceColumns(1, 1);
    console.log('Spalte A gelöscht');
    
    // 3. Apply Merged Cells und entferne cellStyles
    console.log('\n=== 3. applyMergedCellsAfterColumnDelete ===');
    const removedKeys = [];
    
    newMerges.forEach(mergeInfo => {
        const range = mergeInfo.range;
        const parts = range.split(':');
        const startMatch = parts[0].match(/^([A-Z]+)(\d+)$/);
        const endMatch = parts[1].match(/^([A-Z]+)(\d+)$/);
        
        const startCol = colLetterToNumber(startMatch[1]);
        const startRow = parseInt(startMatch[2]);
        const endCol = colLetterToNumber(endMatch[1]);
        const endRow = parseInt(endMatch[2]);
        
        // Setze Wert/Fill auf Master
        const masterCell = ws.getCell(startRow, startCol);
        if (mergeInfo.value) masterCell.value = mergeInfo.value;
        if (mergeInfo.fill) masterCell.fill = mergeInfo.fill;
        
        // Entferne cellStyles für alle Zellen im Merged-Bereich
        for (let row = startRow; row <= endRow; row++) {
            for (let col = startCol; col <= endCol; col++) {
                const rowIdx = row - 1;
                const colIdx = col - 1;
                const key = `${rowIdx}-${colIdx}`;
                
                if (cellStyles[key]) {
                    console.log(`  Entferne cellStyle für ${key} (Excel ${String.fromCharCode(64+col)}${row})`);
                    delete cellStyles[key];
                    removedKeys.push(key);
                }
            }
        }
        
        // Merge
        ws.mergeCells(range);
        console.log(`Merge ${range} gesetzt`);
    });
    
    console.log('\ncellStyles nach Entfernung:', cellStyles);
    console.log('Entfernte Keys:', removedKeys);
    
    // 4. Prüfe Merged Cell F20:H22
    console.log('\n=== 4. ERGEBNIS für F20:H22 ===');
    for (let row = 20; row <= 22; row++) {
        for (let col = 6; col <= 8; col++) {
            const cell = ws.getCell(row, col);
            const letter = String.fromCharCode(64 + col);
            const fill = cell.fill;
            let fillStr = 'keine';
            if (fill && fill.type === 'pattern' && fill.fgColor) {
                fillStr = fill.fgColor.argb || 'unbekannt';
            }
            console.log(`  ${letter}${row}: Fill=${fillStr}, Value="${cell.value || ''}"`);
        }
    }
    
    // 5. Speichern
    await workbook.xlsx.writeFile(testOutputPath);
    console.log('\n=== Gespeichert als', testOutputPath, '===');
    console.log('\nBitte öffne diese Datei und prüfe ob F20:H22 KEINE Hintergrundfarbe hat!');
}

simulateExport().catch(console.error);
