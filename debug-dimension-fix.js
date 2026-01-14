const ExcelJS = require('exceljs');
const { execSync } = require('child_process');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws = wb.getWorksheet(1);
    
    console.log('VOR splice:');
    console.log('  columnCount:', ws.columnCount);
    console.log('  actualColumnCount:', ws.actualColumnCount);
    console.log('  dimensions.range:', ws.dimensions.range);
    console.log('  _columns.length:', ws._columns?.length);
    
    // Lösche Spalte 1
    ws.spliceColumns(1, 1);
    
    console.log('');
    console.log('NACH splice:');
    console.log('  columnCount:', ws.columnCount);
    console.log('  actualColumnCount:', ws.actualColumnCount);
    console.log('  dimensions.range:', ws.dimensions.range);
    console.log('  _columns.length:', ws._columns?.length);
    
    // FIX: Entferne überschüssige Spalten aus _columns
    const actualColCount = ws.dimensions.model.right;
    console.log('');
    console.log('FIX: Kürze _columns auf', actualColCount);
    
    // Entferne alle Einträge > actualColCount
    if (ws._columns) {
        for (let i = actualColCount; i < ws._columns.length; i++) {
            delete ws._columns[i];
        }
        // Sparse array kürzen
        ws._columns.length = actualColCount;
    }
    
    console.log('');
    console.log('NACH Fix:');
    console.log('  columnCount:', ws.columnCount);
    console.log('  actualColumnCount:', ws.actualColumnCount);
    console.log('  dimensions.range:', ws.dimensions.range);
    console.log('  _columns.length:', ws._columns?.length);
    
    // Speichern
    await wb.xlsx.writeFile('/tmp/test-dim-fixed.xlsx');
    
    // Prüfe XML
    execSync('cd /tmp && rm -rf dim_xml2 && mkdir dim_xml2 && cd dim_xml2 && unzip -q /tmp/test-dim-fixed.xlsx');
    const dimResult = execSync("grep -oE 'dimension ref=\"[^\"]*\"' /tmp/dim_xml2/xl/worksheets/sheet1.xml").toString().trim();
    console.log('');
    console.log('In gespeicherter XML:', dimResult);
    
    // Prüfe auch die Spalten im XML
    const colsResult = execSync("grep -oE '<col min=\"[0-9]+\" max=\"[0-9]+\"' /tmp/dim_xml2/xl/worksheets/sheet1.xml | tail -5").toString();
    console.log('');
    console.log('Letzte Spalten-Definitionen:');
    console.log(colsResult);
}

test().catch(console.error);
