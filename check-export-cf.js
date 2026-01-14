const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function checkExportCF() {
    const files = fs.readdirSync('/Users/nojan/Desktop').filter(f => f.startsWith('Export_') && f.endsWith('.xlsx'));
    
    if (files.length === 0) {
        console.log('Keine Export-Datei gefunden auf Desktop');
        return;
    }
    
    // Sortiere nach Datum (neueste zuerst)
    files.sort().reverse();
    const filePath = path.join('/Users/nojan/Desktop', files[0]);
    console.log('Lade neueste Export:', filePath);
    
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    
    const ws = wb.worksheets[0];
    console.log('\nSheet:', ws.name);
    console.log('Column Count:', ws.columnCount);
    
    // Was steht in den letzten Spalten?
    console.log('\n=== Letzte Spalten ===');
    for (let col = ws.columnCount - 2; col <= ws.columnCount + 1; col++) {
        const header = ws.getCell(1, col).value;
        const letter = colNumToLetter(col);
        console.log(`Spalte ${col} (${letter}): "${header || '(leer)'}"`);
    }
    
    const cf = ws.conditionalFormattings;
    if (!cf || cf.length === 0) {
        console.log('\nKeine CF gefunden!');
        return;
    }
    
    console.log('\n=== CF auf hohen Spalten (nach BG) ===');
    const highColCF = cf.filter(c => {
        if (!c.ref) return false;
        const matches = c.ref.match(/([A-Z]+)\d+/g) || [];
        return matches.some(m => {
            const col = m.match(/([A-Z]+)/)[1];
            return colLetterToNum(col) >= 59; // BG = 59
        });
    });
    
    console.log(`${highColCF.length} CF-Regeln auf Spalten >= BG`);
    highColCF.slice(0, 10).forEach((c, i) => {
        console.log(`  ${i}: ${c.ref}`);
    });
    
    // CF auf letzte Spalte?
    const lastCol = ws.columnCount;
    const lastColLetter = colNumToLetter(lastCol);
    console.log(`\n=== CF auf letzte Spalte ${lastCol} (${lastColLetter}) ===`);
    
    const lastColCF = cf.filter(c => {
        if (!c.ref) return false;
        return c.ref.includes(lastColLetter);
    });
    console.log(`${lastColCF.length} CF-Regeln auf Spalte ${lastColLetter}`);
}

function colNumToLetter(num) {
    let result = '';
    while (num > 0) {
        num--;
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result;
}

function colLetterToNum(letters) {
    let result = 0;
    for (let i = 0; i < letters.length; i++) {
        result = result * 26 + (letters.charCodeAt(i) - 64);
    }
    return result;
}

checkExportCF().catch(console.error);
