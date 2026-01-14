const ExcelJS = require('exceljs');

async function analyzeIndexMismatch() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/test-styles.xlsx');
    
    const worksheet = workbook.getWorksheet('Style Tests');
    
    console.log('=== MIT includeEmpty: false ===');
    let dataIndex = 0;
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber === 1) {
            console.log(`HEADER: Excel-Zeile ${rowNumber}`);
            return;
        }
        
        const firstCellValue = row.getCell(1).value;
        const dataRowIndex = rowNumber - 2;
        const styleKeyRow = dataRowIndex + 1;
        
        // Was das Frontend erwartet (originalIndex basiert auf data Array)
        const frontendExpectedKey = `${dataIndex + 1}-0`;
        const actualStyleKey = `${styleKeyRow}-0`;
        
        const mismatch = frontendExpectedKey !== actualStyleKey ? ' ❌ MISMATCH!' : ' ✓';
        
        console.log(`dataArrayIndex=${dataIndex}: Excel-Zeile ${rowNumber} "${String(firstCellValue).slice(0,15)}" -> Frontend erwartet "${frontendExpectedKey}", Style-Key ist "${actualStyleKey}"${mismatch}`);
        
        dataIndex++;
    });
    
    console.log('\n=== LEERE ZEILEN ANALYSE ===');
    for (let i = 1; i <= 15; i++) {
        const row = worksheet.getRow(i);
        const hasData = row.cellCount > 0;
        console.log(`Excel-Zeile ${i}: ${hasData ? 'hat Daten' : 'LEER'}`);
    }
}

analyzeIndexMismatch();
