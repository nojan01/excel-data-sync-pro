const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();

wb.xlsx.readFile('/Users/nojan/Desktop/test-styles Kopie.xlsx').then(() => {
    const ws = wb.getWorksheet('Style Tests');
    
    console.log('=== ALLE Zellen mit Fill (nicht undefined) ===');
    ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
        row.eachCell((cell, colNum) => {
            if (cell.fill && cell.fill.fgColor) {
                console.log(`Row ${rowNum} Col ${colNum}: "${cell.value}" -> Fill: ${cell.fill.fgColor.argb}`);
            }
        });
    });
    
    console.log('\n=== ALLE Zellen mit Rich Text (Objekt-Wert) ===');
    ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
        row.eachCell((cell, colNum) => {
            if (cell.value && typeof cell.value === 'object' && cell.value.richText) {
                console.log(`Row ${rowNum} Col ${colNum}: Rich Text ->`, JSON.stringify(cell.value.richText, null, 2));
            }
        });
    });
    
    console.log('\n=== Check Row 1 cell types ===');
    const row1 = ws.getRow(1);
    for (let col = 1; col <= 8; col++) {
        const cell = row1.getCell(col);
        console.log(`Col ${col}: type=${typeof cell.value}, value=`, cell.value);
    }
}).catch(e => console.error(e));
