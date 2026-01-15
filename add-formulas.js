const ExcelJS = require('exceljs');

async function addFormulas() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/test-styles-exceljs.xlsx');
    const ws = wb.getWorksheet(1);
    
    console.log('Sheet:', ws.name);
    console.log('Zeilen:', ws.rowCount);
    console.log('Spalten:', ws.columnCount);
    
    // Füge verschiedene Formeln in eine neue Spalte ein
    const lastCol = ws.columnCount;
    const newCol = lastCol + 1;
    
    // Header für Formel-Spalte
    ws.getCell(1, newCol).value = 'Formeln';
    ws.getCell(1, newCol).font = { bold: true };
    
    // Verschiedene Formeln einfügen
    ws.getCell(2, newCol).value = { formula: 'SUM(A2:C2)', result: 0 };
    ws.getCell(3, newCol).value = { formula: 'AVERAGE(A3:C3)', result: 0 };
    ws.getCell(4, newCol).value = { formula: 'MAX(A4:C4)', result: 0 };
    ws.getCell(5, newCol).value = { formula: 'MIN(A5:C5)', result: 0 };
    ws.getCell(6, newCol).value = { formula: 'COUNT(A6:C6)', result: 0 };
    ws.getCell(7, newCol).value = { formula: 'IF(A7>0,"Ja","Nein")', result: '' };
    ws.getCell(8, newCol).value = { formula: 'CONCATENATE(A8," ",B8)', result: '' };
    ws.getCell(9, newCol).value = { formula: 'LEN(A9)', result: 0 };
    ws.getCell(10, newCol).value = { formula: 'TODAY()', result: new Date() };
    
    // Auch ein paar Formeln in bestehende Zellen (Spalte B)
    ws.getCell(2, 2).value = { formula: 'UPPER(A2)', result: '' };
    ws.getCell(3, 2).value = { formula: 'LOWER(A3)', result: '' };
    
    await wb.xlsx.writeFile('/Users/nojan/Desktop/test-styles-exceljs.xlsx');
    console.log('Formeln hinzugefuegt in Spalte', newCol);
    console.log('Fertig!');
}

addFormulas().catch(e => console.error('Fehler:', e));
