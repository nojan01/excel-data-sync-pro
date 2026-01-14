const ExcelJS = require('exceljs');

async function test() {
    console.log('=== Test: Hidden Column with Width Preservation ===\n');
    
    // Use the GOOD file that has all widths
    const sourcePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER Kopie.xlsx';
    const targetPath = '/tmp/test-hidden-column.xlsx';
    
    console.log('1. Loading workbook from:', sourcePath);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(sourcePath);
    
    const ws = workbook.getWorksheet('DEFENCE&SPACE Aug-2025');
    
    // Show initial widths
    console.log('\n2. Initial column widths (columns 1-15):');
    for (let i = 1; i <= 15; i++) {
        const col = ws.getColumn(i);
        console.log(`   Column ${i}: width=${col.width}, hidden=${col.hidden}`);
    }
    
    // Save widths FIRST
    const columnWidths = {};
    for (let i = 1; i <= 61; i++) {
        const col = ws.getColumn(i);
        if (col.width !== undefined && col.width !== null) {
            columnWidths[i] = col.width;
        }
    }
    console.log('\n3. Saved', Object.keys(columnWidths).length, 'column widths');
    
    // Now hide column G (7)
    console.log('\n4. Setting column G (7) to hidden=true...');
    const col7 = ws.getColumn(7);
    col7.hidden = true;
    
    // Check what happened to widths
    console.log('\n5. Widths AFTER hiding column 7:');
    let missingAfter = 0;
    for (let i = 1; i <= 15; i++) {
        const col = ws.getColumn(i);
        const w = col.width;
        if (w === undefined) missingAfter++;
        console.log(`   Column ${i}: width=${w}, hidden=${col.hidden}`);
    }
    
    // Now RESTORE widths from saved object
    console.log('\n6. Restoring widths from saved object...');
    for (const [colIdx, width] of Object.entries(columnWidths)) {
        ws.getColumn(parseInt(colIdx)).width = width;
    }
    
    // Check after restore
    console.log('\n7. Widths AFTER restoring:');
    for (let i = 1; i <= 15; i++) {
        const col = ws.getColumn(i);
        console.log(`   Column ${i}: width=${col.width}, hidden=${col.hidden}`);
    }
    
    // Save file
    console.log('\n8. Saving to:', targetPath);
    await workbook.xlsx.writeFile(targetPath);
    
    // Reload and check
    console.log('\n9. Reloading saved file...');
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(targetPath);
    const ws2 = wb2.getWorksheet('DEFENCE&SPACE Aug-2025');
    
    console.log('\n10. Widths in SAVED file:');
    let missingInSaved = 0;
    for (let i = 1; i <= 15; i++) {
        const col = ws2.getColumn(i);
        const w = col.width;
        if (w === undefined) missingInSaved++;
        console.log(`   Column ${i}: width=${w}, hidden=${col.hidden}`);
    }
    
    console.log('\n=== RESULT ===');
    if (missingInSaved > 0) {
        console.log('❌ BUG: ' + missingInSaved + ' columns lost their width after save!');
    } else {
        console.log('✓ All widths preserved correctly even with hidden column!');
    }
}

test().catch(console.error);
