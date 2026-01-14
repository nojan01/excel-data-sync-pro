// Test: Alle Fill-Typen in der Testdatei analysieren
const ExcelJS = require('exceljs');

async function analyzeFills() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/test-styles.xlsx');
    
    const worksheet = workbook.getWorksheet('Style Tests');
    
    console.log('='.repeat(60));
    console.log('ANALYSE: Alle Zellen mit Fill/Font');
    console.log('='.repeat(60));
    
    // Zeile 4 hat die Hintergrundfarben (0-basiert: Zeile 3)
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            const hasFont = cell.font && (cell.font.bold || cell.font.italic || cell.font.color);
            const hasFill = cell.fill && Object.keys(cell.fill).length > 0;
            
            if (hasFont || hasFill) {
                console.log(`\n[${rowNumber}-${colNumber-1}] Value: "${cell.value}"`);
                
                if (cell.font) {
                    console.log('  Font:', JSON.stringify(cell.font));
                }
                
                if (cell.fill) {
                    console.log('  Fill:', JSON.stringify(cell.fill));
                }
            }
        });
    });
    
    // Speziell Zeile 4 (Hintergrundfarben) analysieren
    console.log('\n' + '='.repeat(60));
    console.log('ZEILE 4 (Hintergrundfarben) DETAILLIERT:');
    console.log('='.repeat(60));
    
    const row4 = worksheet.getRow(4);
    row4.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        console.log(`\nSpalte ${colNumber}: "${cell.value}"`);
        console.log('  fill:', JSON.stringify(cell.fill, null, 2));
        console.log('  style.fill:', cell.style?.fill ? JSON.stringify(cell.style.fill, null, 2) : 'KEINE');
    });
}

analyzeFills().catch(console.error);
