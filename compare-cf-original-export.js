const ExcelJS = require('exceljs');

async function test() {
    // Prüfe Original
    console.log('=== Original-Datei ===');
    const wb1 = new ExcelJS.Workbook();
    await wb1.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws1 = wb1.getWorksheet(1);
    const cf1 = ws1.conditionalFormattings;
    
    // Suche nach AZ1:AZ2404 (was nach Löschung zu AY1:AY2404 werden würde)
    console.log('Suche nach AZ-Referenzen (Spalte 52):');
    const azRefs = cf1.filter(c => c.ref && /\bAZ\d/.test(c.ref));
    console.log('Gefunden:', azRefs.length);
    azRefs.slice(0, 5).forEach(c => console.log('  -', c.ref));
    
    console.log('');
    console.log('Suche nach genau "AZ1:AZ2404":');
    const exactAZ = cf1.filter(c => c.ref === 'AZ1:AZ2404');
    console.log('Gefunden:', exactAZ.length);
    
    // Prüfe Export
    console.log('');
    console.log('=== Export-Datei ===');
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile('/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws2 = wb2.getWorksheet(1);
    const cf2 = ws2.conditionalFormattings;
    
    console.log('Suche nach genau "AY1:AY2404":');
    const exactAY = cf2.filter(c => c.ref === 'AY1:AY2404');
    console.log('Gefunden:', exactAY.length);
    
    if (exactAY.length > 0) {
        console.log('');
        console.log('Details dieser CF-Regeln:');
        exactAY.forEach((cf, i) => {
            console.log(`  [${i}] ref: ${cf.ref}`);
            console.log(`      rules: ${cf.rules?.length} Regeln`);
            if (cf.rules?.[0]) {
                console.log(`      erste Regel type: ${cf.rules[0].type}`);
                console.log(`      erste Regel priority: ${cf.rules[0].priority}`);
            }
        });
    }
    
    // Vergleiche: Gibt es im Original eine AZ1:AZ2404?
    console.log('');
    console.log('=== Vergleich ===');
    console.log('Original hat AZ1:AZ2404:', exactAZ.length > 0);
    console.log('Export hat AY1:AY2404:', exactAY.length > 0);
    
    // Wenn nicht, woher kommen die AY1:AY2404?
    console.log('');
    console.log('Original: Alle einzigartigen Refs die nur Spalte 51 (AY) oder 52 (AZ) referenzieren:');
    const onlyAYorAZ = cf1.filter(c => {
        if (!c.ref) return false;
        // Prüfe ob alle Teile nur AY oder AZ sind
        const parts = c.ref.split(/[:\s]/);
        return parts.every(p => /^A[YZ]\d+$/.test(p));
    });
    console.log('Gefunden:', onlyAYorAZ.length);
    onlyAYorAZ.slice(0, 10).forEach(c => console.log('  -', c.ref));
}

test().catch(console.error);
