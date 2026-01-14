const ExcelJS = require('exceljs');

async function test() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Test');
    
    // Zeilen mit unterschiedlichen Styles erstellen
    ws.getCell('A1').value = 'Row1';
    ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } }; // Rot
    
    ws.getCell('A2').value = 'Row2';
    ws.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } }; // Grün
    
    ws.getCell('A3').value = 'Row3';
    ws.getCell('A3').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0000FF' } }; // Blau
    
    console.log('=== Vor Swap ===');
    console.log('Row1:', ws.getCell('A1').value, 'Fill:', ws.getCell('A1').fill?.fgColor?.argb);
    console.log('Row2:', ws.getCell('A2').value, 'Fill:', ws.getCell('A2').fill?.fgColor?.argb);
    console.log('Row3:', ws.getCell('A3').value, 'Fill:', ws.getCell('A3').fill?.fgColor?.argb);
    
    // EINFACHE LÖSUNG: Zeilen komplett austauschen via _rows Array
    // _rows ist 0-basiert, Zeilen sind 1-basiert
    const temp = ws._rows[0]; // Row 1
    ws._rows[0] = ws._rows[2]; // Row 3 -> Row 1 Position
    ws._rows[2] = temp;       // Row 1 -> Row 3 Position
    
    // Row Number in den Zeilen-Objekten aktualisieren
    if (ws._rows[0]) ws._rows[0]._number = 1;
    if (ws._rows[2]) ws._rows[2]._number = 3;
    
    // Zellen-Adressen in beiden Rows aktualisieren
    [0, 2].forEach(rowIdx => {
        const row = ws._rows[rowIdx];
        if (row && row._cells) {
            Object.keys(row._cells).forEach(colKey => {
                const cell = row._cells[colKey];
                if (cell) {
                    cell._row = row;
                }
            });
        }
    });
    
    console.log('\n=== Nach Swap via _rows ===');
    console.log('Row1:', ws.getCell('A1').value, 'Fill:', ws.getCell('A1').fill?.fgColor?.argb);
    console.log('Row2:', ws.getCell('A2').value, 'Fill:', ws.getCell('A2').fill?.fgColor?.argb);
    console.log('Row3:', ws.getCell('A3').value, 'Fill:', ws.getCell('A3').fill?.fgColor?.argb);
    
    // Speichern und prüfen
    await wb.xlsx.writeFile('/tmp/test-row-swap.xlsx');
    console.log('\nDatei gespeichert: /tmp/test-row-swap.xlsx');
    
    // Neu laden und prüfen
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile('/tmp/test-row-swap.xlsx');
    const ws2 = wb2.getWorksheet('Test');
    
    console.log('\n=== Nach Reload ===');
    console.log('Row1:', ws2.getCell('A1').value, 'Fill:', ws2.getCell('A1').fill?.fgColor?.argb);
    console.log('Row2:', ws2.getCell('A2').value, 'Fill:', ws2.getCell('A2').fill?.fgColor?.argb);
    console.log('Row3:', ws2.getCell('A3').value, 'Fill:', ws2.getCell('A3').fill?.fgColor?.argb);
}

test().catch(console.error);
