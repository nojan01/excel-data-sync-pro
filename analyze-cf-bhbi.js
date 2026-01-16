const ExcelJS = require('exceljs');

async function analyzeCF() {
    const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    console.log('Lade Excel-Datei...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.worksheets[0];
    console.log(`Worksheet: ${worksheet.name}`);
    console.log(`Spaltenanzahl: ${worksheet.columnCount}`);
    console.log(`Zeilenanzahl: ${worksheet.rowCount}`);
    
    // Spalten BH (60) und BI (61) - 1-basiert
    const colBH = 60;
    const colBI = 61;
    
    console.log(`\nSpalte ${colBH} (BH): ${worksheet.getColumn(colBH).letter}`);
    console.log(`Spalte ${colBI} (BI): ${worksheet.getColumn(colBI).letter}`);
    
    // Alle bedingten Formatierungen durchsuchen
    const cf = worksheet.conditionalFormattings;
    console.log(`\nAnzahl CF-Regeln: ${cf ? cf.length : 0}`);
    
    if (cf && cf.length > 0) {
        console.log('\n=== CFs die BH oder BI enthalten ===');
        let foundCount = 0;
        
        cf.forEach((cfEntry, idx) => {
            if (!cfEntry.ref) return;
            
            // Prüfe ob BH oder BI in der Referenz vorkommt
            if (cfEntry.ref.includes('BH') || cfEntry.ref.includes('BI')) {
                console.log(`\nCF ${idx}:`);
                console.log(`  Ref: ${cfEntry.ref}`);
                if (cfEntry.rules) {
                    cfEntry.rules.forEach((rule, rIdx) => {
                        console.log(`  Rule ${rIdx}: type=${rule.type}, priority=${rule.priority}`);
                        if (rule.formulae) {
                            console.log(`    Formulae: ${JSON.stringify(rule.formulae)}`);
                        }
                        if (rule.style) {
                            console.log(`    Style: ${JSON.stringify(rule.style)}`);
                        }
                    });
                }
                foundCount++;
            }
        });
        
        if (foundCount === 0) {
            console.log('Keine CFs gefunden die explizit BH oder BI referenzieren.');
        }
        
        // Suche nach CFs die auf die gesamte Zeile angewendet werden könnten
        console.log('\n=== CFs mit großen Bereichen (könnten BH/BI einschließen) ===');
        cf.forEach((cfEntry, idx) => {
            if (!cfEntry.ref) return;
            
            // Parse den Bereich - suche nach Bereichen die bis zur letzten Spalte gehen
            const ranges = cfEntry.ref.split(' ');
            ranges.forEach(range => {
                // Prüfe auf Bereiche wie A1:BI2404 oder ähnlich
                const match = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
                if (match) {
                    const endCol = match[3];
                    // Konvertiere Spaltenbuchstaben zu Nummer
                    let endColNum = 0;
                    for (let i = 0; i < endCol.length; i++) {
                        endColNum = endColNum * 26 + (endCol.charCodeAt(i) - 64);
                    }
                    
                    // Wenn der Bereich bis Spalte 59+ geht
                    if (endColNum >= 59) {
                        console.log(`\nCF ${idx} - Bereich bis ${endCol} (Spalte ${endColNum}):`);
                        console.log(`  Ref: ${cfEntry.ref.substring(0, 200)}${cfEntry.ref.length > 200 ? '...' : ''}`);
                        if (cfEntry.rules && cfEntry.rules.length > 0) {
                            console.log(`  Rules: ${cfEntry.rules.length}`);
                            cfEntry.rules.forEach((rule, rIdx) => {
                                console.log(`    Rule ${rIdx}: type=${rule.type}`);
                                if (rule.formulae) {
                                    console.log(`      Formulae: ${JSON.stringify(rule.formulae)}`);
                                }
                            });
                        }
                    }
                }
            });
        });
        
        // Zeige die letzten 10 CFs
        console.log('\n=== Letzte 10 CF-Einträge ===');
        const lastCFs = cf.slice(-10);
        lastCFs.forEach((cfEntry, idx) => {
            const realIdx = cf.length - 10 + idx;
            console.log(`\nCF ${realIdx}:`);
            console.log(`  Ref: ${cfEntry.ref}`);
        });
    }
    
    // Prüfe auch die Zellen direkt
    console.log('\n=== Zellen BH2 und BI2 ===');
    const cellBH2 = worksheet.getCell(2, colBH);
    const cellBI2 = worksheet.getCell(2, colBI);
    
    console.log(`BH2: value=${cellBH2.value}, style=${JSON.stringify(cellBH2.style)}`);
    console.log(`BI2: value=${cellBI2.value}, style=${JSON.stringify(cellBI2.style)}`);
}

analyzeCF().catch(console.error);
