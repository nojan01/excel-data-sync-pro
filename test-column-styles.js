const ExcelJS = require('exceljs');

async function testColumnStyles() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/test-styles.xlsx');
    
    const worksheet = workbook.worksheets[0];
    
    console.log('\n=== SPALTEN-STYLES ===');
    worksheet.columns.forEach((column, idx) => {
        if (column && column.font) {
            console.log(`\nSpalte ${idx + 1} (${column.key || 'unnamed'}):`);
            console.log('  Font:', JSON.stringify(column.font, null, 2));
        }
    });
    
    console.log('\n\n=== ZEILEN-STYLES (erste 20) ===');
    for (let rowNum = 1; rowNum <= 20; rowNum++) {
        const row = worksheet.getRow(rowNum);
        if (row && row.font) {
            console.log(`\nZeile ${rowNum}:`);
            console.log('  Font:', JSON.stringify(row.font, null, 2));
        }
    }
    
    console.log('\n\n=== ZELL-STYLES mit BOLD in Zeile 5 (Detail) ===');
    const row5 = worksheet.getRow(5);
    row5.eachCell({ includeEmpty: true }, (cell, colNum) => {
        if (cell.font?.bold) {
            console.log(`\nZelle ${colNum} (${cell.address}):`);
            console.log('  Wert:', cell.value);
            console.log('  Cell.font:', JSON.stringify(cell.font, null, 2));
            console.log('  Row.font:', row5.font ? JSON.stringify(row5.font, null, 2) : 'keine');
            const col = worksheet.getColumn(colNum);
            console.log('  Column.font:', col.font ? JSON.stringify(col.font, null, 2) : 'keine');
        }
    });
}

testColumnStyles().catch(console.error);
