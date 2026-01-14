const ExcelJS = require('exceljs');
const { execSync } = require('child_process');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws = wb.getWorksheet(1);
    
    console.log('VOR splice:');
    console.log('  dimensions:', ws.dimensions.range);
    
    // Lösche Spalte 1
    ws.spliceColumns(1, 1);
    
    console.log('');
    console.log('NACH splice (VOR Fix):');
    console.log('  dimensions:', ws.dimensions.range);
    console.log('  Row 1 model.max:', ws.getRow(1).model?.max);
    
    // FIX: Aktualisiere alle Row-Models
    console.log('');
    console.log('Fixe Row-Models...');
    const actualColCount = ws.dimensions.model.right; // 60
    
    ws._rows.forEach((row, idx) => {
        if (row && row._cells) {
            // Entferne Zellen > actualColCount
            for (let i = actualColCount; i < row._cells.length; i++) {
                if (row._cells[i]) {
                    delete row._cells[i];
                }
            }
            row._cells.length = actualColCount;
        }
    });
    
    console.log('');
    console.log('NACH Fix:');
    console.log('  Row 1 dimensions:', ws.getRow(1).dimensions);
    console.log('  Row 1 model.max:', ws.getRow(1).model?.max);
    
    // Speichern
    await wb.xlsx.writeFile('/tmp/test-row-fixed.xlsx');
    
    // Prüfe XML
    execSync('cd /tmp && rm -rf row_xml && mkdir row_xml && cd row_xml && unzip -q /tmp/test-row-fixed.xlsx');
    const dimResult = execSync("grep -oE 'dimension ref=\"[^\"]*\"' /tmp/row_xml/xl/worksheets/sheet1.xml").toString().trim();
    console.log('');
    console.log('In gespeicherter XML:', dimResult);
}

test().catch(console.error);
