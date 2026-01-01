/**
 * Erstellt eine Testdatei mit Formatierungen
 */
const XlsxPopulate = require('xlsx-populate');

async function createTestFile() {
    const workbook = await XlsxPopulate.fromBlankAsync();
    const sheet = workbook.sheet(0).name('Testdaten');
    
    // Header mit Formatierung
    sheet.cell('A1').value('ID').style({ bold: true, fill: '4472C4', fontColor: 'FFFFFF' });
    sheet.cell('B1').value('Name').style({ bold: true, fill: '4472C4', fontColor: 'FFFFFF' });
    sheet.cell('C1').value('Wert').style({ bold: true, fill: '4472C4', fontColor: 'FFFFFF' });
    sheet.cell('D1').value('Status').style({ bold: true, fill: '4472C4', fontColor: 'FFFFFF' });
    
    // Daten mit verschiedenen Formatierungen
    sheet.cell('A2').value(1);
    sheet.cell('B2').value('Test Eintrag 1');
    sheet.cell('C2').value(100).style({ numberFormat: '#,##0.00' });
    sheet.cell('D2').value('Aktiv').style({ fill: '70AD47', fontColor: 'FFFFFF' });
    
    sheet.cell('A3').value(2);
    sheet.cell('B3').value('Test Eintrag 2');
    sheet.cell('C3').value(250.50).style({ numberFormat: '#,##0.00' });
    sheet.cell('D3').value('Inaktiv').style({ fill: 'ED7D31', fontColor: 'FFFFFF' });
    
    sheet.cell('A4').value(3);
    sheet.cell('B4').value('Test Eintrag 3');
    sheet.cell('C4').value(75.25).style({ numberFormat: '#,##0.00' });
    sheet.cell('D4').value('Aktiv').style({ fill: '70AD47', fontColor: 'FFFFFF' });
    
    // Spaltenbreiten
    sheet.column('A').width(10);
    sheet.column('B').width(25);
    sheet.column('C').width(15);
    sheet.column('D').width(15);
    
    await workbook.toFileAsync('test-input.xlsx');
    console.log('Testdatei erstellt: test-input.xlsx');
}

createTestFile().catch(console.error);
