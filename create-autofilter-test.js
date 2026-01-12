#!/usr/bin/env node
// ============================================================================
// TEST: Erstelle Excel-Datei MIT AutoFilter für Visual-Test
// ============================================================================

const ExcelJS = require('exceljs');

async function createAutoFilterTestFile() {
    console.log('📝 Erstelle Test-Datei mit AutoFilter...\n');
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Test Sheet');
    
    // Header setzen
    worksheet.columns = [
        { header: 'ID', key: 'id', width: 10 },
        { header: 'Name', key: 'name', width: 30 },
        { header: 'Status', key: 'status', width: 15 },
        { header: 'Wert', key: 'value', width: 15 }
    ];
    
    // Testdaten einfügen
    const testData = [
        { id: 1, name: 'Projekt Alpha', status: 'Aktiv', value: 1000 },
        { id: 2, name: 'Projekt Beta', status: 'Inaktiv', value: 2000 },
        { id: 3, name: 'Projekt Gamma', status: 'Aktiv', value: 1500 },
        { id: 4, name: 'Projekt Delta', status: 'Geplant', value: 3000 },
        { id: 5, name: 'Projekt Epsilon', status: 'Aktiv', value: 2500 },
        { id: 6, name: 'Projekt Zeta', status: 'Inaktiv', value: 500 },
        { id: 7, name: 'Projekt Eta', status: 'Aktiv', value: 1800 },
        { id: 8, name: 'Projekt Theta', status: 'Geplant', value: 4000 },
        { id: 9, name: 'Projekt Iota', status: 'Aktiv', value: 2200 },
        { id: 10, name: 'Projekt Kappa', status: 'Inaktiv', value: 800 }
    ];
    
    testData.forEach(row => {
        worksheet.addRow(row);
    });
    
    // Formatierung für Header
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9EAD3' } // Hellgrün
    };
    
    // Wichtig: AutoFilter setzen!
    worksheet.autoFilter = {
        from: 'A1',
        to: 'D1'
    };
    
    console.log('   ✓ 10 Zeilen mit Daten');
    console.log('   ✓ AutoFilter gesetzt: A1:D1\n');
    
    // Speichern
    const outputPath = '/Users/nojan/Desktop/AutoFilter_Test.xlsx';
    await workbook.xlsx.writeFile(outputPath);
    
    console.log(`✅ Datei erstellt: ${outputPath}\n`);
    console.log('📋 Jetzt können Sie testen:');
    console.log('   1. Öffnen Sie die Datei in Excel');
    console.log('   2. Prüfen Sie, dass Filter-Buttons in Zeile 1 vorhanden sind');
    console.log('   3. Führen Sie test-visual-rowmove.js damit aus');
    console.log('   4. Öffnen Sie die Test-Ausgabe und prüfen Sie die Filter');
}

createAutoFilterTestFile();
