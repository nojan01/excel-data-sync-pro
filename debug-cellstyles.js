// Debug: Was ist in den cellStyles drin?
const ExcelJS = require('exceljs');
const { extractFillsFromXLSX } = require('./exceljs-reader');

async function debug() {
    const filePath = '/Users/nojan/Desktop/test-styles-exceljs.xlsx';
    
    // extractFillsFromXLSX gibt cellFills zurück, nicht cellStyles!
    const cellFills = extractFillsFromXLSX(filePath);
    
    console.log('=== CELLSTYLES DEBUG ===');
    console.log('Anzahl cellFills:', Object.keys(cellFills).length);
    
    // Zeige alle cellFills für Zeile 4 (Key: rowNumber-1 - colIndex, also Zeile 4 = 3)
    console.log('\n=== Zeile 4 (Key-Präfix: 3-) ===');
    for (const [key, fill] of Object.entries(cellFills)) {
        const [rowIdx, colIdx] = key.split('-').map(Number);
        if (rowIdx === 3) {  // Zeile 4 ist Key 3
            console.log('Key:', key, '-> Spalte', colIdx, '(Excel:', colIdx + 1, '), Fill:', fill);
        }
    }
    
    // Was passiert bei Spalte A (Index 0) löschen?
    console.log('\n=== SIMULATION: Spalte A (Index 0) löschen ===');
    const deletedColumnIndex = 0;  // Spalte A = Index 0
    
    console.log('deletedColumnIndex =', deletedColumnIndex);
    
    for (const [key, fill] of Object.entries(cellFills)) {
        const [rowIdx, colIdx] = key.split('-').map(Number);
        if (rowIdx === 3) {  // Zeile 4
            if (colIdx === deletedColumnIndex) {
                console.log('  Key', key, ': GELÖSCHT (colIdx == deletedColumnIndex)');
            } else if (colIdx > deletedColumnIndex) {
                const newKey = rowIdx + '-' + (colIdx - 1);
                console.log('  Key', key, ': VERSCHOBEN zu', newKey, '(colIdx > deletedColumnIndex)');
            } else {
                console.log('  Key', key, ': UNVERÄNDERT (colIdx < deletedColumnIndex)');
            }
        }
    }
    
    // Zeige alle cellFills
    console.log('\n=== ALLE CELLFILLS ===');
    for (const [key, fill] of Object.entries(cellFills)) {
        console.log('  cellFills["' + key + '"] =', fill);
    }
}

debug().catch(e => console.error(e));
