const ExcelJS = require('exceljs');

async function testDetailedCells() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/test-styles.xlsx');
    
    const worksheet = workbook.worksheets[0];
    
    console.log('\n=== ALLE ZELLEN MIT FORMATIERUNG (erste 100 Zeilen) ===\n');
    
    let cellsWithBold = 0;
    let cellsWithItalic = 0;
    let cellsWithColor = 0;
    let cellsWithFill = 0;
    
    for (let rowNum = 1; rowNum <= Math.min(100, worksheet.rowCount); rowNum++) {
        const row = worksheet.getRow(rowNum);
        
        row.eachCell({ includeEmpty: false }, (cell, colNum) => {
            const hasStyle = cell.font || cell.fill;
            
            if (hasStyle) {
                let info = `${cell.address}: "${cell.value}"`;
                let styles = [];
                
                if (cell.font) {
                    if (cell.font.bold) {
                        styles.push('BOLD');
                        cellsWithBold++;
                    }
                    if (cell.font.italic) {
                        styles.push('ITALIC');
                        cellsWithItalic++;
                    }
                    if (cell.font.underline) styles.push('UNDERLINE');
                    if (cell.font.strike) styles.push('STRIKE');
                    if (cell.font.color?.argb) {
                        const color = cell.font.color.argb.substring(2);
                        if (color !== '000000') {
                            styles.push(`COLOR:${color}`);
                            cellsWithColor++;
                        }
                    }
                }
                
                if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor?.argb) {
                    const fillColor = cell.fill.fgColor.argb.substring(2);
                    styles.push(`FILL:${fillColor}`);
                    cellsWithFill++;
                }
                
                if (styles.length > 0) {
                    console.log(`  ${info} => [${styles.join(', ')}]`);
                }
            }
        });
    }
    
    console.log('\n\n=== ZUSAMMENFASSUNG ===');
    console.log(`Zellen mit BOLD: ${cellsWithBold}`);
    console.log(`Zellen mit ITALIC: ${cellsWithItalic}`);
    console.log(`Zellen mit COLOR: ${cellsWithColor}`);
    console.log(`Zellen mit FILL: ${cellsWithFill}`);
}

testDetailedCells().catch(console.error);
