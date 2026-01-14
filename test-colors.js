const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();

wb.xlsx.readFile('/Users/nojan/Desktop/test-styles Kopie.xlsx').then(() => {
    const ws = wb.getWorksheet('Style Tests');
    
    console.log('=== Zeile 4: Hintergrundfarben - VOLLSTÄNDIGE Analyse ===');
    const row4 = ws.getRow(4);
    for (let col = 1; col <= 8; col++) {
        const cell = row4.getCell(col);
        console.log(`\nCol ${col} ("${cell.value}"):`);
        console.log('  fill:', JSON.stringify(cell.fill));
        if (cell.style) {
            console.log('  style.fill:', JSON.stringify(cell.style.fill));
        }
    }
    
    console.log('\n\n=== Rich Text Zeilen (10 und 14) ===');
    [10, 14].forEach(rowNum => {
        const row = ws.getRow(rowNum);
        console.log(`\nRow ${rowNum}:`);
        row.eachCell({ includeEmpty: false }, (cell, colNum) => {
            console.log(`  Col ${colNum}: type=${typeof cell.value}`);
            if (cell.value && typeof cell.value === 'object') {
                console.log('    richText?', !!cell.value.richText);
                console.log('    value:', JSON.stringify(cell.value));
            } else {
                console.log('    value:', cell.value);
            }
        });
    });
}).catch(e => console.error(e));
