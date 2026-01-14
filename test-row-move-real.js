const ExcelJS = require('exceljs');

async function test() {
    const sourcePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    console.log('=== Lade Original-Datei ===');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(sourcePath);
    const ws = wb.getWorksheet(1);
    
    // Prüfe Zeile 2 und 10 vor dem Swap
    console.log('\n=== VOR dem Umordnen ===');
    console.log('Zeile 2, Zelle A2:', ws.getCell('A2').value);
    console.log('Zeile 2, Zelle A2 Fill:', JSON.stringify(ws.getCell('A2').fill));
    console.log('Zeile 10, Zelle A10:', ws.getCell('A10').value);
    console.log('Zeile 10, Zelle A10 Fill:', JSON.stringify(ws.getCell('A10').fill));
    
    // Simuliere rowMapping: Tausche Zeile 2 und 10 (Datenzeilen 0 und 8)
    // rowMapping[newPos] = originalPos
    // newPos 0 (Zeile 2) soll originalPos 8 (Zeile 10) haben
    // newPos 8 (Zeile 10) soll originalPos 0 (Zeile 2) haben
    const rowMapping = [];
    for (let i = 0; i < 20; i++) {
        if (i === 0) rowMapping[i] = 8;      // Zeile 2 bekommt Inhalt von Zeile 10
        else if (i === 8) rowMapping[i] = 0; // Zeile 10 bekommt Inhalt von Zeile 2
        else rowMapping[i] = i;
    }
    
    console.log('\n=== Führe Zeilen-Umordnung durch ===');
    console.log('rowMapping[0] =', rowMapping[0], '(Zeile 2 <- Original-Zeile', rowMapping[0] + 2, ')');
    console.log('rowMapping[8] =', rowMapping[8], '(Zeile 10 <- Original-Zeile', rowMapping[8] + 2, ')');
    
    // Neues _rows Array erstellen
    const headerRow = ws._rows[0];
    const newRows = [headerRow];
    
    for (let newDataIdx = 0; newDataIdx < rowMapping.length; newDataIdx++) {
        const originalDataIdx = rowMapping[newDataIdx];
        const originalRowsIdx = originalDataIdx + 1;
        const row = ws._rows[originalRowsIdx];
        
        if (row) {
            row._number = newDataIdx + 2;
            newRows.push(row);
        } else {
            newRows.push(undefined);
        }
    }
    
    ws._rows = newRows;
    
    console.log('\n=== NACH dem Umordnen (vor Werte schreiben) ===');
    console.log('Zeile 2, Zelle A2:', ws.getCell('A2').value);
    console.log('Zeile 2, Zelle A2 Fill:', JSON.stringify(ws.getCell('A2').fill));
    console.log('Zeile 10, Zelle A10:', ws.getCell('A10').value);
    console.log('Zeile 10, Zelle A10 Fill:', JSON.stringify(ws.getCell('A10').fill));
    
    // Speichern
    await wb.xlsx.writeFile('/tmp/test-row-move-result.xlsx');
    console.log('\nDatei gespeichert: /tmp/test-row-move-result.xlsx');
    
    // Neu laden und prüfen
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile('/tmp/test-row-move-result.xlsx');
    const ws2 = wb2.getWorksheet(1);
    
    console.log('\n=== Nach Reload ===');
    console.log('Zeile 2, Zelle A2:', ws2.getCell('A2').value);
    console.log('Zeile 2, Zelle A2 Fill:', JSON.stringify(ws2.getCell('A2').fill));
    console.log('Zeile 10, Zelle A10:', ws2.getCell('A10').value);
    console.log('Zeile 10, Zelle A10 Fill:', JSON.stringify(ws2.getCell('A10').fill));
}

test().catch(console.error);
