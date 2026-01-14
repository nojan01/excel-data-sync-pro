const ExcelJS = require('exceljs');

// Analysiere die bedingte Formatierung der großen Datei
async function analyzeCF() {
    // Passe den Pfad an deine große Datei an!
    const exportPath = '/Users/nojan/Desktop/Export_test-styles-exceljs.xlsx';
    
    console.log('=== ANALYSE BEDINGTE FORMATIERUNG ===\n');
    console.log('Datei:', exportPath);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(exportPath);
    const ws = workbook.getWorksheet(1);
    
    console.log('\n=== BEDINGTE FORMATIERUNGEN ===');
    const cf = ws.conditionalFormattings;
    
    if (!cf || cf.length === 0) {
        console.log('Keine bedingten Formatierungen gefunden');
        return;
    }
    
    console.log('Anzahl CF-Einträge:', cf.length);
    
    cf.forEach((cfEntry, idx) => {
        console.log(`\n--- CF #${idx + 1} ---`);
        console.log('Ref:', cfEntry.ref);
        
        if (cfEntry.rules) {
            cfEntry.rules.forEach((rule, ruleIdx) => {
                console.log(`  Rule ${ruleIdx + 1}:`);
                console.log('    Type:', rule.type);
                console.log('    Priority:', rule.priority);
                if (rule.formulae) {
                    console.log('    Formeln:', rule.formulae);
                }
                if (rule.style) {
                    console.log('    Style:', JSON.stringify(rule.style));
                }
            });
        }
    });
    
    // Prüfe ob CF-Bereiche sich mit Merged Cells überlappen
    console.log('\n=== MERGED CELLS ===');
    const merges = ws.model.merges || [];
    console.log('Merged Cells:', merges);
    
    console.log('\n=== ÜBERLAPPUNGEN CF <-> MERGED ===');
    merges.forEach(mergeRange => {
        cf.forEach((cfEntry, cfIdx) => {
            if (rangesOverlap(mergeRange, cfEntry.ref)) {
                console.log(`ÜBERLAPPUNG: Merge ${mergeRange} <-> CF #${cfIdx + 1} (${cfEntry.ref})`);
            }
        });
    });
}

// Hilfsfunktion: Prüft ob zwei Bereiche sich überlappen
function rangesOverlap(range1, range2) {
    const r1 = parseRange(range1);
    const r2 = parseRange(range2);
    if (!r1 || !r2) return false;
    
    // Keine Überlappung wenn einer komplett links/rechts/oben/unten vom anderen ist
    return !(r1.endCol < r2.startCol || r1.startCol > r2.endCol ||
             r1.endRow < r2.startRow || r1.startRow > r2.endRow);
}

function parseRange(range) {
    const parts = range.split(':');
    if (parts.length !== 2) return null;
    
    const start = parseCell(parts[0]);
    const end = parseCell(parts[1]);
    if (!start || !end) return null;
    
    return {
        startCol: start.col,
        startRow: start.row,
        endCol: end.col,
        endRow: end.row
    };
}

function parseCell(cell) {
    const match = cell.match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;
    
    let col = 0;
    for (let i = 0; i < match[1].length; i++) {
        col = col * 26 + (match[1].charCodeAt(i) - 64);
    }
    
    return { col, row: parseInt(match[2]) };
}

analyzeCF().catch(console.error);
