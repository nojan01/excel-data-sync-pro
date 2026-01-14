/**
 * Tiefer Test: Was passiert GENAU bei spliceColumns?
 */
const ExcelJS = require('exceljs');
const path = require('path');

async function testDeep() {
    const testFile = path.join(process.env.HOME, 'Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    console.log('=== DEEP TEST: spliceColumns ===\n');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(testFile);
    
    const worksheet = workbook.worksheets[0];
    console.log(`Sheet: ${worksheet.name}`);
    
    // Zeile 1, erste 10 Spalten vor dem Splice
    console.log('\n=== ZEILE 1 VOR spliceColumns ===');
    for (let col = 1; col <= 10; col++) {
        const cell = worksheet.getCell(1, col);
        const fill = cell.fill;
        const fillInfo = fill && fill.type === 'pattern' ? 
            (fill.fgColor?.argb || fill.fgColor?.theme || 'kein fgColor') : 
            (fill?.type || 'kein Fill');
        console.log(`  Col ${col} (${getColLetter(col)}): "${cell.value}" | Fill: ${fillInfo}`);
    }
    
    // Splice
    console.log('\n>>> worksheet.spliceColumns(1, 1) <<<\n');
    worksheet.spliceColumns(1, 1);
    
    // Nach dem Splice
    console.log('=== ZEILE 1 NACH spliceColumns ===');
    for (let col = 1; col <= 9; col++) {
        const cell = worksheet.getCell(1, col);
        const fill = cell.fill;
        const fillInfo = fill && fill.type === 'pattern' ? 
            (fill.fgColor?.argb || fill.fgColor?.theme || 'kein fgColor') : 
            (fill?.type || 'kein Fill');
        console.log(`  Col ${col} (${getColLetter(col)}): "${cell.value}" | Fill: ${fillInfo}`);
    }
    
    // Prüfe Row-Level Styles
    console.log('\n=== ROW 1 PROPERTIES ===');
    const row1 = worksheet.getRow(1);
    console.log('Row numFmt:', row1.numFmt);
    console.log('Row font:', JSON.stringify(row1.font));
    console.log('Row fill:', JSON.stringify(row1.fill));
    
    // Prüfe Column-Level Styles
    console.log('\n=== COLUMN STYLES ===');
    for (let col = 1; col <= 10; col++) {
        const column = worksheet.getColumn(col);
        const fill = column.style?.fill;
        const fillInfo = fill && fill.type === 'pattern' ? 
            (fill.fgColor?.argb || fill.fgColor?.theme || 'kein fgColor') : 
            (fill?.type || 'kein Fill');
        console.log(`  Col ${col} (${getColLetter(col)}): width=${column.width}, Fill: ${fillInfo}`);
    }
}

function getColLetter(num) {
    let result = '';
    while (num > 0) {
        num--;
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result;
}

testDeep().catch(console.error);
