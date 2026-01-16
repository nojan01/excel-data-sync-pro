const ExcelJS = require('exceljs');

async function analyzeTable() {
    const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    console.log('Lade Excel-Datei...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.worksheets[0];
    
    console.log('=== Worksheet Tables Property ===');
    console.log('worksheet.tables:', worksheet.tables);
    
    if (worksheet.tables) {
        for (const [name, table] of Object.entries(worksheet.tables)) {
            console.log(`\nTabelle: ${name}`);
            console.log('Alle Properties:');
            for (const [key, value] of Object.entries(table)) {
                if (typeof value === 'object') {
                    console.log(`  ${key}: ${JSON.stringify(value)}`);
                } else {
                    console.log(`  ${key}: ${value}`);
                }
            }
        }
    }
    
    // Prüfe auch die interne _tables Eigenschaft
    console.log('\n=== Internal _tables ===');
    if (worksheet._tables) {
        console.log('worksheet._tables:', worksheet._tables);
    }
    
    // Versuche die Tabellendefinition zu finden
    console.log('\n=== Model ===');
    if (worksheet.model && worksheet.model.tables) {
        console.log('worksheet.model.tables:', JSON.stringify(worksheet.model.tables, null, 2));
    }
}

analyzeTable().catch(console.error);
