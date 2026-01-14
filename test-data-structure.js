const ExcelJS = require('exceljs');

async function analyzeDataStructure() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/test-styles.xlsx');
    
    const worksheet = workbook.getWorksheet('Style Tests');
    
    console.log('=== EXCEL ZEILEN-STRUKTUR ===');
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const firstCellValue = row.getCell(1).value;
        const dataRowIndex = rowNumber - 2; // Wie im Code
        const styleKeyRow = dataRowIndex + 1;
        const isHeader = rowNumber === 1;
        
        console.log(`Excel-Zeile ${rowNumber}: "${firstCellValue}" -> dataRowIndex=${dataRowIndex}, styleKeyRow=${styleKeyRow}, isHeader=${isHeader}`);
    });
}

analyzeDataStructure();
