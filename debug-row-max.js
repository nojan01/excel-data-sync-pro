const ExcelJS = require('exceljs');
const { execSync } = require('child_process');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws = wb.getWorksheet(1);
    
    // Lösche Spalte 1
    ws.spliceColumns(1, 1);
    
    console.log('NACH splice:');
    console.log('  dimensions.range:', ws.dimensions.range);
    
    // Finde Rows mit max > 60
    console.log('');
    console.log('Rows mit max > 60:');
    let count = 0;
    ws._rows.forEach((row, idx) => {
        if (row && row.model && row.model.max > 60) {
            if (count < 10) {
                console.log('  Row', idx + 1, ': min=', row.model.min, 'max=', row.model.max);
            }
            count++;
        }
    });
    console.log('  Gesamt:', count, 'rows');
    
    // Prüfe Row 1 genauer
    console.log('');
    console.log('Row 1 Details:');
    const row1 = ws.getRow(1);
    console.log('  model.min:', row1.model?.min);
    console.log('  model.max:', row1.model?.max);
    console.log('  dimensions:', row1.dimensions);
    
    // Prüfe Zellen in Row 1
    console.log('  Zellen 58-62:');
    for (let c = 58; c <= 62; c++) {
        const cell = row1.getCell(c);
        console.log('    Col', c, ':', cell.value, '| address:', cell.address);
    }
}

test().catch(console.error);
