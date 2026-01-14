// Analyse der Merged Cells und deren Inhalte/Fills
const ExcelJS = require('exceljs');

async function analyze() {
    const originalPath = '/Users/nojan/Desktop/test-styles-exceljs.xlsx';
    const exportPath = '/Users/nojan/Desktop/Export_test-styles-exceljs.xlsx';
    
    const wb1 = new ExcelJS.Workbook();
    const wb2 = new ExcelJS.Workbook();
    
    await wb1.xlsx.readFile(originalPath);
    await wb2.xlsx.readFile(exportPath);
    
    const ws1 = wb1.worksheets[0];
    const ws2 = wb2.worksheets[0];
    
    console.log('=== MERGED CELLS ANALYSE ===');
    console.log('\nOriginal Merges:', ws1.model.merges);
    console.log('Export Merges:', ws2.model.merges);
    
    // Analysiere Zeilen 20-22 im Detail
    console.log('\n=== ZEILEN 20-22 (Merged Cell Bereiche) ===');
    
    console.log('\n--- ORIGINAL ---');
    for (let row = 20; row <= 22; row++) {
        console.log('Zeile ' + row + ':');
        for (let col = 1; col <= 9; col++) {
            const cell = ws1.getCell(row, col);
            const fill = cell.fill;
            let fillStr = '-';
            if (fill && fill.type === 'pattern' && fill.fgColor) {
                fillStr = fill.fgColor.argb || fill.fgColor.theme || 'x';
            }
            const isMaster = cell.master ? '(slave of ' + cell.master.address + ')' : '';
            console.log('  ' + cell.address + ': "' + cell.value + '" Fill=' + fillStr + ' ' + isMaster);
        }
    }
    
    console.log('\n--- EXPORT ---');
    for (let row = 20; row <= 22; row++) {
        console.log('Zeile ' + row + ':');
        for (let col = 1; col <= 8; col++) {
            const cell = ws2.getCell(row, col);
            const fill = cell.fill;
            let fillStr = '-';
            if (fill && fill.type === 'pattern' && fill.fgColor) {
                fillStr = fill.fgColor.argb || fill.fgColor.theme || 'x';
            }
            const isMaster = cell.master ? '(slave of ' + cell.master.address + ')' : '';
            console.log('  ' + cell.address + ': "' + cell.value + '" Fill=' + fillStr + ' ' + isMaster);
        }
    }
    
    // Erwartete Werte nach Spalte A löschen
    console.log('\n=== ERWARTETE WERTE NACH SPALTE A LÖSCHEN ===');
    console.log('Original Merged Regions:');
    console.log('  G20:I22 -> F20:H22 (war "Merged Cell G-I", Fill=FFE0E0)');
    console.log('  A22:C22 -> sollte A22:B22 werden (war "Horizontal Merge (A-C)", Fill=F0F0F0)');
    console.log('  D22:F22 -> sollte C22:E22 werden (war "Merge 2 (D-F)", Fill=E0E0FF)');
    
    console.log('\n=== PRÜFE ZEILE 22 SPEZIFISCH ===');
    console.log('\n--- ORIGINAL ZEILE 22 ---');
    for (let col = 1; col <= 9; col++) {
        const cell = ws1.getCell(22, col);
        const fill = cell.fill;
        let fillStr = '-';
        if (fill && fill.type === 'pattern' && fill.fgColor) {
            fillStr = fill.fgColor.argb?.substring(2) || fill.fgColor.theme || 'x';
        }
        console.log('  Spalte ' + col + ' (' + cell.address + '): "' + cell.value + '" Fill=' + fillStr);
    }
    
    console.log('\n--- EXPORT ZEILE 22 ---');
    for (let col = 1; col <= 8; col++) {
        const cell = ws2.getCell(22, col);
        const fill = cell.fill;
        let fillStr = '-';
        if (fill && fill.type === 'pattern' && fill.fgColor) {
            fillStr = fill.fgColor.argb?.substring(2) || fill.fgColor.theme || 'x';
        }
        console.log('  Spalte ' + col + ' (' + cell.address + '): "' + cell.value + '" Fill=' + fillStr);
    }
    
    console.log('\n--- ERWARTUNG für Export Zeile 22 ---');
    console.log('  Spalte 1 (A): "Horizontal Merge (A-C)" von Orig B (Merge A-C wurde zu A-B)');
    console.log('  Spalte 2 (B): Teil von A22:B22 Merge');
    console.log('  Spalte 3 (C): "Merge 2 (D-F)" von Orig D (Merge D-F wurde zu C-E)');
    console.log('  Spalte 4 (D): Teil von C22:E22 Merge');
    console.log('  Spalte 5 (E): Teil von C22:E22 Merge');
}

analyze().catch(e => console.error(e));
