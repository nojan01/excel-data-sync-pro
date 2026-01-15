const ExcelJS = require('exceljs');

function formatDate(date, numFmt) {
    numFmt = numFmt || '';
    
    const hasTime = numFmt.includes('h') || numFmt.includes('H') || numFmt.includes(':');
    
    if (hasTime) {
        return date.toISOString().replace('T', ' ').substring(0, 19);
    }
    
    const day = date.getDate();
    const month = date.getMonth() + 1;
    const year = date.getFullYear();
    
    const dayStr = numFmt.includes('dd') ? String(day).padStart(2, '0') : String(day);
    const monthStr = numFmt.includes('mm') ? String(month).padStart(2, '0') : String(month);
    
    let yearStr = String(year);
    if (!numFmt.includes('yyyy') && numFmt.includes('yy')) {
        yearStr = yearStr.substring(2);
    }
    
    if (numFmt.includes('.')) {
        return `${dayStr}.${monthStr}.${yearStr}`;
    } else if (numFmt.includes('-')) {
        return `${monthStr}-${dayStr}-${yearStr}`;
    } else {
        return `${monthStr}.${dayStr}.${yearStr}`;
    }
}

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws = wb.getWorksheet(1);
    
    console.log('=== Formatierte Datumswerte ===');
    
    for (let row = 2; row <= 5; row++) {
        const r = ws.getRow(row);
        r.eachCell({ includeEmpty: false }, (cell, col) => {
            const val = cell.value;
            if (val instanceof Date) {
                const formatted = formatDate(val, cell.numFmt);
                console.log(`R${row}C${col}: numFmt="${cell.numFmt}" => ${formatted}`);
            }
        });
    }
}

test().catch(e => console.error(e));
