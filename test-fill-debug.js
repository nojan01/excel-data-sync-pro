const ExcelJS = require('exceljs');

async function testFillApplication() {
    const sourcePath = '/Users/nojan/Desktop/AutoFilter_Test.xlsx';
    const targetPath = '/Users/nojan/Desktop/Test_Debug_Fill.xlsx';
    
    // Test cellStyles wie sie vom Reader kommen würden
    const cellStyles = {
        '0-0': { fill: '#D9EAD3' },
        '0-1': { fill: '#FF0000' },  // B1 sollte rot werden!
        '0-2': { fill: '#D9EAD3' },
        '0-3': { fill: '#D9EAD3' }
    };
    
    console.log('=== TEST FILL APPLICATION ===\n');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(sourcePath);
    
    const worksheet = workbook.worksheets[0];
    
    console.log('VOR dem Anwenden:');
    console.log('  B1 fill:', JSON.stringify(worksheet.getCell('B1').fill));
    
    // Fills anwenden
    for (const [styleKey, style] of Object.entries(cellStyles)) {
        if (!style.fill) continue;
        
        const [rowIdx, colIdx] = styleKey.split('-').map(Number);
        const cell = worksheet.getCell(rowIdx + 1, colIdx + 1);
        
        const existingFill = cell.fill;
        const hasFill = existingFill && 
                        existingFill.type === 'pattern' && 
                        existingFill.pattern === 'solid' &&
                        existingFill.fgColor?.argb;
        
        console.log(`Zelle ${cell.address}: existingFill=${JSON.stringify(existingFill)}, hasFill=${hasFill}`);
        
        if (!hasFill) {
            const hex = style.fill.replace('#', '');
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF' + hex }
            };
            console.log(`  -> Fill angewendet: #${hex}`);
        }
    }
    
    console.log('\nNACH dem Anwenden:');
    console.log('  B1 fill:', JSON.stringify(worksheet.getCell('B1').fill));
    
    // Speichern
    await workbook.xlsx.writeFile(targetPath);
    console.log('\nGespeichert in:', targetPath);
    
    // Erneut laden und prüfen
    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(targetPath);
    console.log('\nNach erneutem Laden:');
    console.log('  B1 fill:', JSON.stringify(workbook2.worksheets[0].getCell('B1').fill));
}

testFillApplication().catch(console.error);
