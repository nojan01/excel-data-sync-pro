const ExcelJS = require('exceljs');
const { extractFillsFromXLSX } = require('./exceljs-reader');

async function testExportFix() {
    const sourcePath = '/Users/nojan/Desktop/AutoFilter_Test.xlsx';
    const targetPath = '/Users/nojan/Desktop/Test_Fixed_Export.xlsx';
    
    console.log('=== TEST EXPORT FIX ===\n');
    
    // 1. Workbook laden
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(sourcePath);
    
    const worksheet = workbook.worksheets[0];
    console.log('Sheet:', worksheet.name);
    
    // 2. Prüfen was ExcelJS liest
    console.log('\nExcelJS liest für B1:');
    console.log('  fill:', JSON.stringify(worksheet.getCell('B1').fill));
    
    // 3. Fills direkt aus XML extrahieren
    console.log('\nExtrahiere Fills aus XLSX...');
    const directFills = extractFillsFromXLSX(sourcePath, worksheet.name);
    console.log('Gefundene Fills:', JSON.stringify(directFills, null, 2));
    
    // 4. cellStyles aufbauen wie es der Writer macht
    const cellStyles = {};
    for (const [key, fillColor] of Object.entries(directFills)) {
        cellStyles[key] = { fill: fillColor };
    }
    
    // 5. Fehlende Fills anwenden
    console.log('\nApplyMissingFills simulieren...');
    for (const [styleKey, style] of Object.entries(cellStyles)) {
        if (!style.fill) continue;
        
        const [rowIdx, colIdx] = styleKey.split('-').map(Number);
        const cell = worksheet.getCell(rowIdx + 1, colIdx + 1);
        
        const existingFill = cell.fill;
        const hasFill = existingFill && 
                        existingFill.type === 'pattern' && 
                        existingFill.pattern === 'solid' &&
                        existingFill.fgColor?.argb;
        
        if (!hasFill) {
            const hex = style.fill.replace('#', '');
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF' + hex }
            };
            console.log(`  -> ${cell.address}: Fill angewendet #${hex}`);
        } else {
            console.log(`  -> ${cell.address}: Bereits Fill vorhanden (${existingFill.fgColor.argb})`);
        }
    }
    
    // 6. Speichern
    await workbook.xlsx.writeFile(targetPath);
    console.log('\nGespeichert:', targetPath);
    
    // 7. Ergebnis prüfen
    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(targetPath);
    
    console.log('\n=== ERGEBNIS ===');
    console.log('A1 fill:', JSON.stringify(workbook2.worksheets[0].getCell('A1').fill));
    console.log('B1 fill:', JSON.stringify(workbook2.worksheets[0].getCell('B1').fill));
    console.log('C1 fill:', JSON.stringify(workbook2.worksheets[0].getCell('C1').fill));
}

testExportFix().catch(console.error);
