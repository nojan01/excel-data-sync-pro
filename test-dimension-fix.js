const ExcelJS = require('exceljs');
const { execSync } = require('child_process');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws = wb.getWorksheet(1);
    
    console.log('VOR:', ws.columnCount);
    
    // Lösche Spalte 1
    ws.spliceColumns(1, 1);
    
    console.log('NACH splice:', ws.columnCount);
    console.log('dimensions.right:', ws.dimensions.model.right);
    
    // Fix: Entferne die letzte (leere) Spalte aus dem internen Array
    if (ws._columns && ws._columns.length > ws.dimensions.model.right) {
        console.log('_columns Länge vor:', ws._columns.length);
        ws._columns.length = ws.dimensions.model.right;
        console.log('_columns Länge nach:', ws._columns.length);
    }
    
    console.log('columnCount nach Fix:', ws.columnCount);
    
    // Speichern
    await wb.xlsx.writeFile('/tmp/test-splice-fixed.xlsx');
    
    // Prüfe XML
    execSync('cd /tmp && rm -rf test_xml2 && mkdir test_xml2 && cd test_xml2 && unzip -q /tmp/test-splice-fixed.xlsx');
    const result = execSync('grep -oE \'dimension ref="[^"]*"\' /tmp/test_xml2/xl/worksheets/sheet1.xml').toString();
    console.log('Dimension nach Fix:', result.trim());
}

test().catch(console.error);
