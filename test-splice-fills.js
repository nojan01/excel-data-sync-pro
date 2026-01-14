/**
 * Isolierter Test: Zeigt was spliceColumns mit Fills macht
 */
const ExcelJS = require('exceljs');
const path = require('path');

async function testSpliceFills() {
    const testFile = path.join(process.env.HOME, 'Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    console.log('=== TEST: spliceColumns und Fills ===\n');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(testFile);
    
    const worksheet = workbook.worksheets[0];
    console.log(`Sheet: ${worksheet.name}`);
    console.log(`Rows: ${worksheet.rowCount}, Cols: ${worksheet.columnCount}\n`);
    
    // Finde erst mal wo überhaupt Fills sind
    console.log('Suche nach Zellen mit Fills...');
    const allFills = [];
    for (let row = 1; row <= Math.min(100, worksheet.rowCount); row++) {
        for (let col = 1; col <= Math.min(20, worksheet.columnCount); col++) {
            const cell = worksheet.getCell(row, col);
            const fill = cell.fill;
            if (fill && fill.type === 'pattern' && fill.fgColor) {
                allFills.push({ row, col, fill: fill.fgColor.argb || fill.fgColor.theme || JSON.stringify(fill.fgColor) });
            }
        }
    }
    console.log(`Gefundene Fills (erste 100 Zeilen, erste 20 Spalten): ${allFills.length}`);
    allFills.slice(0, 30).forEach(f => console.log(`  Row ${f.row}, Col ${f.col} (${getColLetter(f.col)}): ${f.fill}`));
    
    // Speichere Fills VOR dem splice
    const beforeFills = {};
    for (let row = 1; row <= 100; row++) {
        for (let col = 1; col <= 20; col++) {
            const cell = worksheet.getCell(row, col);
            const fill = cell.fill;
            const key = `${row}-${col}`;
            if (fill && fill.type === 'pattern' && fill.fgColor) {
                beforeFills[key] = {
                    value: String(cell.value || '').substring(0, 20),
                    fill: fill.fgColor.argb || fill.fgColor.theme || JSON.stringify(fill.fgColor),
                    colLetter: getColLetter(col)
                };
            }
        }
    }
    
    console.log('=== VOR spliceColumns (Spalte A löschen) ===');
    console.log('Zellen mit Fills (Zeilen 10-15, Spalten A-J):');
    for (const [key, data] of Object.entries(beforeFills)) {
        console.log(`  ${key} (${data.colLetter}): "${data.value}" -> Fill: ${data.fill}`);
    }
    console.log();
    
    // Lösche Spalte A (1)
    console.log('>>> worksheet.spliceColumns(1, 1) - Spalte A löschen <<<\n');
    worksheet.spliceColumns(1, 1);
    
    // Speichere Fills NACH dem splice
    // WICHTIG: Nach dem splice ist die alte Spalte B jetzt Spalte A (col 1)
    // Also die Zelle die vorher bei col=2 war, ist jetzt bei col=1
    const afterFills = {};
    for (let row = 10; row <= 15; row++) {
        for (let col = 1; col <= 9; col++) { // Eine Spalte weniger
            const cell = worksheet.getCell(row, col);
            const fill = cell.fill;
            const key = `${row}-${col}`;
            if (fill && fill.type === 'pattern' && fill.fgColor) {
                afterFills[key] = {
                    value: cell.value,
                    fill: fill.fgColor.argb || fill.fgColor.theme || JSON.stringify(fill.fgColor),
                    colLetter: getColLetter(col)
                };
            }
        }
    }
    
    console.log('=== NACH spliceColumns ===');
    console.log('Zellen mit Fills (Zeilen 10-15, Spalten A-I):');
    for (const [key, data] of Object.entries(afterFills)) {
        console.log(`  ${key} (${data.colLetter}): "${data.value}" -> Fill: ${data.fill}`);
    }
    console.log();
    
    // Vergleich: Was sollte wo sein?
    console.log('=== ANALYSE ===');
    console.log('Erwartung: Zelle vorher bei Spalte B (col=2) sollte jetzt bei Spalte A (col=1) sein\n');
    
    // Für jede Zelle die vorher einen Fill hatte (außer in Spalte A):
    // Prüfe ob der Fill jetzt eine Spalte weiter links ist
    for (const [beforeKey, beforeData] of Object.entries(beforeFills)) {
        const [row, col] = beforeKey.split('-').map(Number);
        if (col === 1) {
            console.log(`  ${beforeKey} war in Spalte A - wurde gelöscht`);
            continue;
        }
        
        const newCol = col - 1;
        const afterKey = `${row}-${newCol}`;
        const afterData = afterFills[afterKey];
        
        if (afterData) {
            if (afterData.fill === beforeData.fill) {
                console.log(`  ✅ ${beforeKey} -> ${afterKey}: Fill korrekt verschoben (${beforeData.fill})`);
            } else {
                console.log(`  ❌ ${beforeKey} -> ${afterKey}: Fill FALSCH! Vorher: ${beforeData.fill}, Nachher: ${afterData.fill}`);
            }
        } else {
            console.log(`  ❌ ${beforeKey} -> ${afterKey}: Fill VERLOREN! Vorher: ${beforeData.fill}, Nachher: kein Fill`);
        }
    }
    
    // Prüfe auch ob es neue Fills gibt die vorher nicht da waren
    console.log('\nUnerwartete Fills (vorher nicht da):');
    for (const [afterKey, afterData] of Object.entries(afterFills)) {
        const [row, col] = afterKey.split('-').map(Number);
        const oldCol = col + 1;
        const beforeKey = `${row}-${oldCol}`;
        
        if (!beforeFills[beforeKey]) {
            console.log(`  ⚠️  ${afterKey}: Neuer Fill ${afterData.fill} (war vorher nicht bei ${beforeKey})`);
        }
    }
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

testSpliceFills().catch(console.error);
