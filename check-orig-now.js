const ExcelJS = require('exceljs');

async function check() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws = wb.worksheets[0];
    
    console.log('=== ORIGINAL ===');
    console.log('columnCount:', ws.columnCount);
    
    console.log('\n=== LETZTE SPALTEN ===');
    for (let col = 58; col <= 62; col++) {
        const header = ws.getRow(1).getCell(col).value;
        const letter = colLetter(col);
        console.log('Spalte ' + col + ' (' + letter + '): ' + (header || '(leer)'));
    }
    
    // Suche CFs mit BI
    const cf = ws.conditionalFormattings || [];
    let biCount = 0;
    for (const entry of cf) {
        if (entry.ref && entry.ref.includes('BI')) biCount++;
    }
    console.log('\nCFs mit BI:', biCount);
}

function colLetter(num) {
    let r = '';
    while (num > 0) {
        const rem = (num - 1) % 26;
        r = String.fromCharCode(65 + rem) + r;
        num = Math.floor((num - 1) / 26);
    }
    return r;
}

check().catch(console.error);
