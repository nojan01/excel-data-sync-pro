const ExcelJS = require('exceljs');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws = wb.getWorksheet(1);
    
    console.log('=== ExcelJS conditionalFormattings Analyse ===');
    console.log('');
    
    // Prüfe ob es ein Getter ist
    const descriptor = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(ws), 'conditionalFormattings');
    console.log('Property descriptor:');
    console.log('  Has getter:', !!descriptor?.get);
    console.log('  Has setter:', !!descriptor?.set);
    console.log('  Writable:', descriptor?.writable);
    console.log('  Configurable:', descriptor?.configurable);
    console.log('');
    
    // Hole CF
    const cf = ws.conditionalFormattings;
    console.log('CF Type:', typeof cf);
    console.log('CF is Array:', Array.isArray(cf));
    console.log('CF Length:', cf?.length);
    console.log('');
    
    // Suche AY-Referenzen
    const ayRefs = cf.filter(c => c.ref && c.ref.includes('AY'));
    console.log('AY-Referenzen im Original:', ayRefs.length);
    if (ayRefs.length > 0) {
        console.log('  Erste:', ayRefs[0].ref);
    }
    
    console.log('');
    console.log('=== Test: Ändere eine CF-Referenz ===');
    
    // Ändere die erste CF
    if (cf.length > 0) {
        const oldRef = cf[0].ref;
        console.log('Vorher cf[0].ref:', oldRef);
        cf[0].ref = 'TEST:TEST';
        console.log('Nachher cf[0].ref:', cf[0].ref);
        
        // Prüfe ob die Änderung persistiert
        const cfNachher = ws.conditionalFormattings;
        console.log('Nach erneutem Aufruf von ws.conditionalFormattings:');
        console.log('  cf[0].ref:', cfNachher[0].ref);
        console.log('  Gleiche Referenz?:', cf === cfNachher);
    }
    
    console.log('');
    console.log('=== Test: Speichern und erneut laden ===');
    
    // Speichern
    const testPath = '/tmp/cf-test.xlsx';
    await wb.xlsx.writeFile(testPath);
    
    // Neu laden
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(testPath);
    const ws2 = wb2.getWorksheet(1);
    const cf2 = ws2.conditionalFormattings;
    
    console.log('Nach Speichern und Laden:');
    console.log('  cf[0].ref:', cf2[0]?.ref);
    console.log('  Hat sich geändert?:', cf2[0]?.ref === 'TEST:TEST');
}

test().catch(console.error);
