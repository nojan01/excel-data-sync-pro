// Test: Simuliere exakt was der exceljs-reader macht
const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();

wb.xlsx.readFile('/Users/nojan/Desktop/test-styles Kopie.xlsx').then(() => {
    const ws = wb.getWorksheet('Style Tests');
    
    console.log('=== Simuliere exceljs-reader.js für Row 14 ===\n');
    
    const row = ws.getRow(14);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        let cellValue = cell.value;
        
        console.log(`Col ${colNumber}:`);
        console.log('  Original value type:', typeof cellValue);
        console.log('  Original value:', cellValue);
        
        // Objekt-Werte behandeln (wie im Reader)
        if (cellValue && typeof cellValue === 'object') {
            if (cellValue.richText) {
                console.log('  -> Rich Text erkannt!');
                const richText = cellValue.richText.map(part => part.text);
                cellValue = richText.join('');
                console.log('  -> Konvertiert zu:', cellValue);
            } else if (cellValue.text !== undefined && cellValue.hyperlink !== undefined) {
                console.log('  -> Hyperlink erkannt!');
                cellValue = cellValue.text;
            } else if (cellValue.text !== undefined) {
                console.log('  -> Text-Objekt erkannt!');
                cellValue = cellValue.text;
            } else if (cellValue === null) {
                cellValue = '';
            }
        }
        
        console.log('  Final value:', cellValue);
        console.log('');
    });
}).catch(e => console.error(e));
