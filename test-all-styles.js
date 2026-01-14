/**
 * Test: Was passiert mit ALLEN Styles bei spliceColumns?
 */
const ExcelJS = require('exceljs');
const path = require('path');

async function testAllStyles() {
    const testFile = path.join(process.env.HOME, 'Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    console.log('=== TEST: Alle Styles bei spliceColumns ===\n');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(testFile);
    
    const worksheet = workbook.worksheets[0];
    console.log(`Sheet: ${worksheet.name}`);
    
    // Zeilen mit verschiedenen Styles finden
    console.log('\n=== ZEILE 1-3 VOR spliceColumns ===');
    for (let row = 1; row <= 3; row++) {
        console.log(`\nZeile ${row}:`);
        for (let col = 1; col <= 10; col++) {
            const cell = worksheet.getCell(row, col);
            const styles = [];
            
            // Fill
            if (cell.fill?.type === 'pattern' && cell.fill.fgColor) {
                styles.push(`Fill:${cell.fill.fgColor.argb || cell.fill.fgColor.theme}`);
            }
            // Font
            if (cell.font) {
                if (cell.font.bold) styles.push('Bold');
                if (cell.font.color?.argb) styles.push(`Color:${cell.font.color.argb.substring(0,6)}`);
                if (cell.font.size) styles.push(`Size:${cell.font.size}`);
            }
            // Border
            if (cell.border) {
                const borders = [];
                if (cell.border.top) borders.push('T');
                if (cell.border.bottom) borders.push('B');
                if (cell.border.left) borders.push('L');
                if (cell.border.right) borders.push('R');
                if (borders.length) styles.push(`Border:${borders.join('')}`);
            }
            
            const value = String(cell.value || '').substring(0, 15);
            const styleStr = styles.length ? styles.join(',') : '-';
            console.log(`  ${getColLetter(col)}: "${value}" [${styleStr}]`);
        }
    }
    
    // Speichere Style-Snapshots vor dem Splice
    const beforeStyles = {};
    for (let row = 1; row <= 100; row++) {
        for (let col = 1; col <= 20; col++) {
            const cell = worksheet.getCell(row, col);
            const key = `${row}-${col}`;
            beforeStyles[key] = getStyleSnapshot(cell);
        }
    }
    
    // Splice
    console.log('\n\n>>> worksheet.spliceColumns(1, 1) <<<\n');
    worksheet.spliceColumns(1, 1);
    
    // Nach dem Splice
    console.log('=== ZEILE 1-3 NACH spliceColumns ===');
    for (let row = 1; row <= 3; row++) {
        console.log(`\nZeile ${row}:`);
        for (let col = 1; col <= 9; col++) {
            const cell = worksheet.getCell(row, col);
            const styles = [];
            
            // Fill
            if (cell.fill?.type === 'pattern' && cell.fill.fgColor) {
                styles.push(`Fill:${cell.fill.fgColor.argb || cell.fill.fgColor.theme}`);
            }
            // Font
            if (cell.font) {
                if (cell.font.bold) styles.push('Bold');
                if (cell.font.color?.argb) styles.push(`Color:${cell.font.color.argb.substring(0,6)}`);
                if (cell.font.size) styles.push(`Size:${cell.font.size}`);
            }
            // Border
            if (cell.border) {
                const borders = [];
                if (cell.border.top) borders.push('T');
                if (cell.border.bottom) borders.push('B');
                if (cell.border.left) borders.push('L');
                if (cell.border.right) borders.push('R');
                if (borders.length) styles.push(`Border:${borders.join('')}`);
            }
            
            const value = String(cell.value || '').substring(0, 15);
            const styleStr = styles.length ? styles.join(',') : '-';
            console.log(`  ${getColLetter(col)}: "${value}" [${styleStr}]`);
        }
    }
    
    // Vergleich
    console.log('\n\n=== VERGLEICH: Styles nach Verschiebung ===');
    let correct = 0;
    let wrong = 0;
    let lost = 0;
    
    for (let row = 1; row <= 100; row++) {
        for (let col = 2; col <= 20; col++) { // Starte bei col 2 (war vorher B)
            const beforeKey = `${row}-${col}`;
            const afterKey = `${row}-${col - 1}`; // Jetzt eine Spalte links
            const before = beforeStyles[beforeKey];
            const after = getStyleSnapshot(worksheet.getCell(row, col - 1));
            
            if (!before.hasStyle && !after.hasStyle) {
                continue; // Beide ohne Style - OK
            }
            
            if (before.hasStyle && !after.hasStyle) {
                lost++;
                if (lost <= 10) {
                    console.log(`❌ LOST: Zeile ${row}, ${getColLetter(col)} -> ${getColLetter(col-1)}: Style verloren`);
                    console.log(`   Vorher: ${before.summary}`);
                }
            } else if (before.summary !== after.summary) {
                wrong++;
                if (wrong <= 10) {
                    console.log(`⚠️  CHANGED: Zeile ${row}, ${getColLetter(col)} -> ${getColLetter(col-1)}:`);
                    console.log(`   Vorher: ${before.summary}`);
                    console.log(`   Nachher: ${after.summary}`);
                }
            } else {
                correct++;
            }
        }
    }
    
    console.log(`\n=== ZUSAMMENFASSUNG ===`);
    console.log(`Korrekt: ${correct}`);
    console.log(`Verändert: ${wrong}`);
    console.log(`Verloren: ${lost}`);
}

function getStyleSnapshot(cell) {
    const styles = [];
    
    if (cell.fill?.type === 'pattern' && cell.fill.fgColor) {
        styles.push(`Fill:${cell.fill.fgColor.argb || cell.fill.fgColor.theme}`);
    }
    if (cell.font) {
        if (cell.font.bold) styles.push('Bold');
        if (cell.font.italic) styles.push('Italic');
        if (cell.font.underline) styles.push('Underline');
        if (cell.font.color?.argb) styles.push(`Color:${cell.font.color.argb}`);
        if (cell.font.size) styles.push(`Size:${cell.font.size}`);
        if (cell.font.name) styles.push(`Font:${cell.font.name}`);
    }
    if (cell.border) {
        const borders = [];
        if (cell.border.top) borders.push('T');
        if (cell.border.bottom) borders.push('B');
        if (cell.border.left) borders.push('L');
        if (cell.border.right) borders.push('R');
        if (borders.length) styles.push(`Border:${borders.join('')}`);
    }
    if (cell.alignment) {
        if (cell.alignment.horizontal) styles.push(`H:${cell.alignment.horizontal}`);
        if (cell.alignment.vertical) styles.push(`V:${cell.alignment.vertical}`);
    }
    
    return {
        hasStyle: styles.length > 0,
        summary: styles.join(',') || '-'
    };
}

function getColLetter(num) {
    let result = '';
    while (num > 0) {
        num--;
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result;
}

testAllStyles().catch(console.error);
