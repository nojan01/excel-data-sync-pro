const ExcelJS = require('exceljs');

async function test() {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Test');
    
    // Erstelle 5 Spalten: A, B, C, D, E
    ws.getCell('A1').value = 'A';
    ws.getCell('B1').value = 'B';
    ws.getCell('C1').value = 'C';
    ws.getCell('D1').value = 'D';
    ws.getCell('E1').value = 'E';
    
    console.log('VOR spliceColumns:');
    console.log('columnCount:', ws.columnCount);
    for (let i = 1; i <= 6; i++) {
        console.log(`  Spalte ${i}: "${ws.getCell(1, i).value || '(leer)'}"`);
    }
    
    // Lösche Spalte A (1)
    ws.spliceColumns(1, 1);
    
    console.log('\nNACH spliceColumns(1, 1):');
    console.log('columnCount:', ws.columnCount);
    for (let i = 1; i <= 6; i++) {
        console.log(`  Spalte ${i}: "${ws.getCell(1, i).value || '(leer)'}"`);
    }
    
    // Was passiert mit CF?
    console.log('\n=== TEST MIT CF ===');
    const workbook2 = new ExcelJS.Workbook();
    const ws2 = workbook2.addWorksheet('Test2');
    
    for (let i = 1; i <= 5; i++) {
        ws2.getCell(1, i).value = 'Col' + i;
    }
    
    // Füge CF auf Spalte E (5) hinzu
    ws2.addConditionalFormatting({
        ref: 'E1:E10',
        rules: [{ type: 'cellIs', operator: 'equal', formulae: ['"X"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } } }]
    });
    
    console.log('VOR spliceColumns:');
    console.log('CF count:', ws2.conditionalFormattings.length);
    ws2.conditionalFormattings.forEach(cf => console.log('  CF ref:', cf.ref));
    
    // Lösche Spalte A
    ws2.spliceColumns(1, 1);
    
    console.log('\nNACH spliceColumns(1, 1):');
    console.log('columnCount:', ws2.columnCount);
    console.log('CF count:', ws2.conditionalFormattings.length);
    ws2.conditionalFormattings.forEach(cf => console.log('  CF ref:', cf.ref));
    
    // Zeige Spalten
    for (let i = 1; i <= 5; i++) {
        console.log(`  Spalte ${i}: "${ws2.getCell(1, i).value || '(leer)'}"`);
    }
}

test().catch(console.error);
