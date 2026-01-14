const ExcelJS = require('exceljs');
const { execSync } = require('child_process');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws = wb.getWorksheet(1);
    
    console.log('VOR splice:');
    console.log('  dimensions.range:', ws.dimensions.range);
    console.log('  dimensions.model:', ws.dimensions.model);
    
    // Lösche Spalte 1
    ws.spliceColumns(1, 1);
    
    console.log('');
    console.log('NACH splice:');
    console.log('  dimensions.range:', ws.dimensions.range);
    console.log('  dimensions.model:', ws.dimensions.model);
    
    // Speichern
    await wb.xlsx.writeFile('/tmp/test-dim.xlsx');
    
    // Prüfe XML
    execSync('cd /tmp && rm -rf dim_xml && mkdir dim_xml && cd dim_xml && unzip -q /tmp/test-dim.xlsx');
    const dimResult = execSync("grep -oE 'dimension ref=\"[^\"]*\"' /tmp/dim_xml/xl/worksheets/sheet1.xml").toString().trim();
    console.log('');
    console.log('In gespeicherter XML:', dimResult);
    
    // Prüfe auch die Spalten im XML
    const colsResult = execSync("grep -oE '<col min=\"[0-9]+\" max=\"[0-9]+\"' /tmp/dim_xml/xl/worksheets/sheet1.xml | tail -5").toString();
    console.log('');
    console.log('Letzte Spalten-Definitionen:');
    console.log(colsResult);
}

test().catch(console.error);
