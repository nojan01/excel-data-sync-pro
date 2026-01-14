const ExcelJS = require('exceljs');

async function compareStyles() {
    const originalPath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    const exportPath = '/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    const wbOriginal = new ExcelJS.Workbook();
    const wbExport = new ExcelJS.Workbook();
    
    await wbOriginal.xlsx.readFile(originalPath);
    await wbExport.xlsx.readFile(exportPath);
    
    const wsOriginal = wbOriginal.getWorksheet(1);
    const wsExport = wbExport.getWorksheet(1);
    
    console.log('=== VERGLEICH: Original vs Export ===\n');
    console.log(`Original: ${wsOriginal.rowCount} Zeilen, ${wsOriginal.columnCount} Spalten`);
    console.log(`Export: ${wsExport.rowCount} Zeilen, ${wsExport.columnCount} Spalten`);
    
    // Vergleiche Styles für die ersten 20 Zeilen, erste 10 Spalten
    const maxRows = Math.min(30, wsOriginal.rowCount);
    const maxCols = Math.min(10, wsOriginal.columnCount);
    
    let missingStyles = [];
    let differentStyles = [];
    
    for (let row = 1; row <= maxRows; row++) {
        for (let col = 1; col <= maxCols; col++) {
            const cellOrig = wsOriginal.getCell(row, col);
            const cellExport = wsExport.getCell(row, col);
            
            const origHasFill = cellOrig.style?.fill && 
                (cellOrig.style.fill.fgColor || cellOrig.style.fill.bgColor || cellOrig.style.fill.pattern !== 'none');
            const exportHasFill = cellExport.style?.fill && 
                (cellExport.style.fill.fgColor || cellExport.style.fill.bgColor || cellExport.style.fill.pattern !== 'none');
            
            const origHasFont = cellOrig.style?.font && Object.keys(cellOrig.style.font).length > 0;
            const exportHasFont = cellExport.style?.font && Object.keys(cellExport.style.font).length > 0;
            
            const origHasBorder = cellOrig.style?.border && Object.keys(cellOrig.style.border).length > 0;
            const exportHasBorder = cellExport.style?.border && Object.keys(cellExport.style.border).length > 0;
            
            // Check for missing fills
            if (origHasFill && !exportHasFill) {
                missingStyles.push({
                    row, col,
                    type: 'fill',
                    original: JSON.stringify(cellOrig.style.fill),
                    export: 'MISSING'
                });
            }
            
            // Check for different fills
            if (origHasFill && exportHasFill) {
                const origFillStr = JSON.stringify(cellOrig.style.fill);
                const exportFillStr = JSON.stringify(cellExport.style.fill);
                if (origFillStr !== exportFillStr) {
                    differentStyles.push({
                        row, col,
                        type: 'fill',
                        original: origFillStr,
                        export: exportFillStr
                    });
                }
            }
            
            // Check for missing fonts
            if (origHasFont && !exportHasFont) {
                missingStyles.push({
                    row, col,
                    type: 'font',
                    original: JSON.stringify(cellOrig.style.font),
                    export: 'MISSING'
                });
            }
            
            // Check for missing borders
            if (origHasBorder && !exportHasBorder) {
                missingStyles.push({
                    row, col,
                    type: 'border',
                    original: JSON.stringify(cellOrig.style.border),
                    export: 'MISSING'
                });
            }
        }
    }
    
    console.log('\n=== FEHLENDE STYLES ===');
    console.log(`Anzahl: ${missingStyles.length}`);
    
    // Gruppiere nach Zeile
    const byRow = {};
    missingStyles.forEach(s => {
        if (!byRow[s.row]) byRow[s.row] = [];
        byRow[s.row].push(s);
    });
    
    Object.keys(byRow).slice(0, 10).forEach(row => {
        console.log(`\nZeile ${row}:`);
        byRow[row].forEach(s => {
            console.log(`  Spalte ${s.col} (${s.type}): ${s.original.substring(0, 80)}...`);
        });
    });
    
    console.log('\n=== UNTERSCHIEDLICHE STYLES ===');
    console.log(`Anzahl: ${differentStyles.length}`);
    differentStyles.slice(0, 5).forEach(s => {
        console.log(`\nZeile ${s.row}, Spalte ${s.col} (${s.type}):`);
        console.log(`  Original: ${s.original.substring(0, 100)}`);
        console.log(`  Export:   ${s.export.substring(0, 100)}`);
    });
    
    // Zeige konkrete Beispiele
    console.log('\n=== BEISPIELE EINZELNER ZELLEN ===');
    
    // Zeile 2 (erste Datenzeile) vs. andere Zeilen
    for (let row of [2, 3, 5, 10, 15]) {
        const cellOrig = wsOriginal.getCell(row, 1);
        const cellExport = wsExport.getCell(row, 1);
        console.log(`\nZeile ${row}, Spalte A:`);
        console.log(`  Original Value: "${cellOrig.value}"`);
        console.log(`  Export Value: "${cellExport.value}"`);
        console.log(`  Original Fill: ${JSON.stringify(cellOrig.style?.fill || 'none')}`);
        console.log(`  Export Fill: ${JSON.stringify(cellExport.style?.fill || 'none')}`);
    }
}

compareStyles().catch(console.error);
