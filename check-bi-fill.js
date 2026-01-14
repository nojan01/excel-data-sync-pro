const ExcelJS = require('exceljs');

async function check() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    const ws = workbook.worksheets[0];
    
    console.log('=== PRÜFE SPALTE 61 (BI) DIREKT ===');
    
    // Prüfe die ersten 20 Zellen in Spalte 61
    for (let row = 1; row <= 20; row++) {
        const cell = ws.getCell(row, 61);
        const value = cell.value;
        const fill = cell.fill;
        const hasFill = fill && fill.type === 'pattern' && fill.pattern !== 'none';
        
        if (value || hasFill) {
            console.log(`BI${row}: value="${value || ''}", fill=${JSON.stringify(fill)}`);
        }
    }
    
    // Prüfe auch Spalte 60 (BH) zum Vergleich
    console.log('\n=== PRÜFE SPALTE 60 (BH) ZUM VERGLEICH ===');
    for (let row = 1; row <= 5; row++) {
        const cell = ws.getCell(row, 60);
        console.log(`BH${row}: value="${cell.value || ''}", fill=${JSON.stringify(cell.fill)}`);
    }
    
    // Zähle Zellen mit Fill in Spalte 61
    console.log('\n=== ZÄHLE FILLS IN SPALTE 61 ===');
    let fillCount = 0;
    for (let row = 1; row <= 2500; row++) {
        const cell = ws.getCell(row, 61);
        const fill = cell.fill;
        if (fill && fill.type === 'pattern' && fill.pattern !== 'none') {
            fillCount++;
        }
    }
    console.log('Zellen mit Fill in Spalte 61 (BI):', fillCount);
    
    // Zähle in Spalte 60
    let fillCount60 = 0;
    for (let row = 1; row <= 2500; row++) {
        const cell = ws.getCell(row, 60);
        const fill = cell.fill;
        if (fill && fill.type === 'pattern' && fill.pattern !== 'none') {
            fillCount60++;
        }
    }
    console.log('Zellen mit Fill in Spalte 60 (BH):', fillCount60);
}

check().catch(console.error);
