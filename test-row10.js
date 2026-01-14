const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();

async function test() {
    await wb.xlsx.readFile('/Users/nojan/Desktop/test-styles.xlsx');
    const ws = wb.getWorksheet('Style Tests');
    
    console.log('=== ZEILE 10 (Schriftformatierungen) ===');
    const row10 = ws.getRow(10);
    row10.eachCell((cell, col) => {
        console.log(`[10-${col}] "${cell.value}" -> Font:`, JSON.stringify(cell.font, null, 2));
    });
    
    console.log('\n=== ZEILE 9 (Label Schriftformatierungen) ===');
    const row9 = ws.getRow(9);
    row9.eachCell((cell, col) => {
        console.log(`[9-${col}] "${cell.value}" -> Font:`, JSON.stringify(cell.font, null, 2));
    });
}

test();
