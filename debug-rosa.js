// Debug: Woher kommt die rosa Fill für F20:H22?
const ExcelJS = require('exceljs');
const { extractFillsFromXLSX } = require('./exceljs-reader');

async function debug() {
    const originalPath = '/Users/nojan/Desktop/test-styles-exceljs.xlsx';
    
    // 1. Was sagt ExcelJS über das Original?
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(originalPath);
    const ws = wb.worksheets[0];
    
    console.log('=== ORIGINAL MERGED BEREICH G20:I22 ===');
    for (let row = 20; row <= 22; row++) {
        for (let col = 7; col <= 9; col++) {
            const cell = ws.getCell(row, col);
            const fill = cell.fill;
            let fillStr = 'keine';
            if (fill && fill.type === 'pattern' && fill.fgColor) {
                fillStr = JSON.stringify(fill.fgColor);
            }
            console.log('  ' + cell.address + ': Fill=' + fillStr);
        }
    }
    
    // 2. Was extrahiert extractFillsFromXLSX?
    console.log('\n=== EXTRACTED FILLS ===');
    const cellFills = extractFillsFromXLSX(originalPath);
    
    // Zeige Fills für Zeilen 19-21 (0-basiert: 18-20)
    console.log('\nFills für Zeilen 19-22:');
    for (const [key, fill] of Object.entries(cellFills)) {
        const [rowIdx, colIdx] = key.split('-').map(Number);
        // rowIdx ist 0-basiert? Oder 1-basiert?
        // Aus dem Reader-Code: dataRowIndex = rowNum - 1
        // Also rowNum 20 -> rowIdx 19
        if (rowIdx >= 18 && rowIdx <= 21) {
            console.log('  cellFills["' + key + '"] = ' + fill);
        }
    }
    
    // 3. Prüfe speziell die Spalten 6,7,8 (G,H,I) nach Index 5,6,7
    console.log('\nSpeziell Spalten G,H,I (Index 6,7,8):');
    for (let row = 19; row <= 22; row++) {
        for (let col = 6; col <= 8; col++) {
            const key = row + '-' + col;
            if (cellFills[key]) {
                console.log('  cellFills["' + key + '"] = ' + cellFills[key]);
            }
        }
    }
    
    // 4. Was ist FFE0E0?
    console.log('\n=== SUCHE FFE0E0 (rosa) ===');
    for (const [key, fill] of Object.entries(cellFills)) {
        if (fill && fill.includes('E0E0')) {
            console.log('  cellFills["' + key + '"] = ' + fill);
        }
    }
}

debug().catch(e => console.error(e));
