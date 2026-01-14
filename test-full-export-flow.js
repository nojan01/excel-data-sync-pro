const ExcelJS = require('exceljs');

async function test() {
    console.log('=== Simulating Full Export Flow ===\n');
    
    const sourcePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER Kopie.xlsx';
    const targetPath = '/tmp/test-export-widths.xlsx';
    
    // 1. Load workbook (wie exceljs-writer.js)
    console.log('1. Loading workbook from:', sourcePath);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(sourcePath);
    
    const ws = workbook.getWorksheet('DEFENCE&SPACE Aug-2025');
    const headers = [];
    const data = [];
    
    // Get headers and data (simulate what the app does)
    ws.getRow(1).eachCell((cell, colNumber) => {
        headers[colNumber - 1] = cell.value;
    });
    
    // Get data rows (just first 10 for test)
    for (let i = 2; i <= 11; i++) {
        const rowData = [];
        ws.getRow(i).eachCell((cell, colNumber) => {
            rowData[colNumber - 1] = cell.value;
        });
        data.push(rowData);
    }
    
    console.log('   Headers count:', headers.length);
    console.log('   Data rows:', data.length);
    
    // 2. Save column widths BEFORE any changes
    const columnWidths = {};
    const actualColumnCount = Math.max(headers.length, ws.columnCount || 0);
    
    console.log('\n2. Saving column widths...');
    console.log('   actualColumnCount:', actualColumnCount);
    
    for (let colIdx = 1; colIdx <= actualColumnCount; colIdx++) {
        const col = ws.getColumn(colIdx);
        if (col.width !== undefined && col.width !== null) {
            columnWidths[colIdx] = col.width;
        }
    }
    console.log('   Saved widths:', Object.keys(columnWidths).length);
    
    // Show first 10 saved widths
    console.log('   First 10 saved:', Object.entries(columnWidths).slice(0, 10).map(([k, v]) => `${k}:${v.toFixed(2)}`).join(', '));
    
    // 3. Write headers and data (like processSheet does)
    console.log('\n3. Writing headers and data...');
    headers.forEach((header, colIndex) => {
        const cell = ws.getCell(1, colIndex + 1);
        cell.value = header;
    });
    
    data.forEach((row, rowIndex) => {
        row.forEach((value, colIndex) => {
            const cell = ws.getCell(rowIndex + 2, colIndex + 1);
            cell.value = value === null || value === undefined ? '' : value;
        });
    });
    
    // 4. Check widths AFTER writing
    console.log('\n4. Checking widths AFTER writing...');
    let missingAfterWrite = 0;
    for (let i = 1; i <= 10; i++) {
        const w = ws.getColumn(i).width;
        if (w === undefined) missingAfterWrite++;
    }
    console.log('   Missing widths in first 10 cols:', missingAfterWrite);
    
    // 5. Restore widths (like processSheet does)
    console.log('\n5. Restoring column widths...');
    for (const [colIdx, width] of Object.entries(columnWidths)) {
        const col = ws.getColumn(parseInt(colIdx));
        col.width = width;
    }
    console.log('   Restored:', Object.keys(columnWidths).length, 'widths');
    
    // 6. Check widths AFTER restore
    console.log('\n6. Checking widths AFTER restore...');
    for (let i = 1; i <= 15; i++) {
        const w = ws.getColumn(i).width;
        console.log('   Column', i, ':', w);
    }
    
    // 7. Save file
    console.log('\n7. Saving to:', targetPath);
    await workbook.xlsx.writeFile(targetPath);
    
    // 8. Reload and check
    console.log('\n8. Reloading saved file to verify...');
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(targetPath);
    const ws2 = wb2.getWorksheet('DEFENCE&SPACE Aug-2025');
    
    console.log('\n9. Widths in saved file:');
    let missingInSaved = 0;
    for (let i = 1; i <= 20; i++) {
        const w = ws2.getColumn(i).width;
        const status = w === undefined ? 'MISSING!' : w.toFixed(2);
        console.log('   Column', i, ':', status);
        if (w === undefined) missingInSaved++;
    }
    
    console.log('\n=== RESULT ===');
    console.log('Widths saved from original:', Object.keys(columnWidths).length);
    console.log('Missing widths after reload:', missingInSaved, '(of first 20 columns)');
    
    if (missingInSaved > 0) {
        console.log('\n!!! BUG CONFIRMED: Column widths are lost during export!');
    } else {
        console.log('\n✓ All widths preserved correctly');
    }
}

test().catch(console.error);
