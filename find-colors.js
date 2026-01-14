const ExcelJS = require('exceljs');

async function findColoredCellDifferences() {
  const origPath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
  const exportPath = '/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
  
  const wbOrig = new ExcelJS.Workbook();
  const wbExport = new ExcelJS.Workbook();
  
  console.log('Lade Original...');
  await wbOrig.xlsx.readFile(origPath);
  console.log('Lade Export...');
  await wbExport.xlsx.readFile(exportPath);
  
  const sheetOrig = wbOrig.worksheets[0];
  const sheetExport = wbExport.worksheets[0];
  
  // Finde zuerst Zeilen mit echten Farbfills in Spalte L oder M
  console.log('\n=== Suche nach Zellen mit echten Farben (nicht pattern-empty) ===\n');
  
  const coloredCellsOrig = [];
  const coloredCellsExport = [];
  
  // Scan Original Spalten L (12), M (13), N (14)
  for (let row = 1; row <= 500; row++) {
    for (let col = 12; col <= 14; col++) {
      const cell = sheetOrig.getCell(row, col);
      const fillInfo = getDetailedFill(cell);
      if (fillInfo.hasColor) {
        coloredCellsOrig.push({ row, col, colLetter: getColLetter(col), ...fillInfo });
      }
    }
  }
  
  // Scan Export Spalten K (11), L (12), M (13)
  for (let row = 1; row <= 500; row++) {
    for (let col = 11; col <= 13; col++) {
      const cell = sheetExport.getCell(row, col);
      const fillInfo = getDetailedFill(cell);
      if (fillInfo.hasColor) {
        coloredCellsExport.push({ row, col, colLetter: getColLetter(col), ...fillInfo });
      }
    }
  }
  
  console.log(`Original: ${coloredCellsOrig.length} Zellen mit Farbe in L/M/N (erste 500 Zeilen)`);
  console.log(`Export: ${coloredCellsExport.length} Zellen mit Farbe in K/L/M (erste 500 Zeilen)`);
  
  if (coloredCellsOrig.length > 0) {
    console.log('\nOriginal Farbzellen (erste 10):');
    coloredCellsOrig.slice(0, 10).forEach(c => {
      console.log(`  ${c.colLetter}${c.row}: ${c.colorStr}`);
    });
  }
  
  if (coloredCellsExport.length > 0) {
    console.log('\nExport Farbzellen (erste 10):');
    coloredCellsExport.slice(0, 10).forEach(c => {
      console.log(`  ${c.colLetter}${c.row}: ${c.colorStr}`);
    });
  }
  
  // Prüfe alle Spalten für eine bestimmte Zeile mit Farben
  console.log('\n=== Suche erste Zeile mit Farbe überhaupt ===\n');
  
  for (let row = 1; row <= 200; row++) {
    for (let col = 1; col <= 61; col++) {
      const cellOrig = sheetOrig.getCell(row, col);
      const fillInfo = getDetailedFill(cellOrig);
      if (fillInfo.hasColor) {
        console.log(`Erste Farbzelle gefunden: Zeile ${row}, Spalte ${getColLetter(col)}`);
        console.log(`  Farbe: ${fillInfo.colorStr}`);
        
        // Zeige die ganze Zeile
        console.log(`\nZeile ${row} komplett (Original):`);
        for (let c = 1; c <= Math.min(20, sheetOrig.columnCount); c++) {
          const cell = sheetOrig.getCell(row, c);
          const f = getDetailedFill(cell);
          if (f.hasColor) {
            console.log(`  ${getColLetter(c)}: ${f.colorStr}`);
          }
        }
        
        console.log(`\nZeile ${row} komplett (Export):`);
        for (let c = 1; c <= Math.min(20, sheetExport.columnCount); c++) {
          const cell = sheetExport.getCell(row, c);
          const f = getDetailedFill(cell);
          if (f.hasColor) {
            console.log(`  ${getColLetter(c)}: ${f.colorStr}`);
          }
        }
        
        return; // Nur erste zeigen
      }
    }
  }
  
  console.log('Keine Farbzellen in ersten 200 Zeilen gefunden');
  
  // Prüfe weiter hinten
  console.log('\n=== Suche ab Zeile 2000 ===');
  for (let row = 2000; row <= 2100; row++) {
    for (let col = 1; col <= 61; col++) {
      const cellOrig = sheetOrig.getCell(row, col);
      const fillInfo = getDetailedFill(cellOrig);
      if (fillInfo.hasColor) {
        console.log(`Farbzelle: Zeile ${row}, Spalte ${getColLetter(col)}: ${fillInfo.colorStr}`);
      }
    }
  }
}

function getDetailedFill(cell) {
  if (!cell.fill) return { hasColor: false, colorStr: 'none' };
  if (cell.fill.type !== 'pattern') return { hasColor: false, colorStr: cell.fill.type };
  
  const fg = cell.fill.fgColor;
  if (!fg) return { hasColor: false, colorStr: 'pattern-empty' };
  
  if (fg.argb) {
    // Ignore FFFFFFFF (white) and 00000000 (transparent)
    if (fg.argb === 'FFFFFFFF' || fg.argb === '00000000') {
      return { hasColor: false, colorStr: 'white/transparent' };
    }
    return { hasColor: true, colorStr: `argb:${fg.argb}` };
  }
  if (fg.theme !== undefined) {
    return { hasColor: true, colorStr: `theme:${fg.theme}${fg.tint ? `,tint:${fg.tint.toFixed(2)}` : ''}` };
  }
  if (fg.indexed !== undefined) {
    if (fg.indexed === 64) { // Default/none
      return { hasColor: false, colorStr: 'indexed:64' };
    }
    return { hasColor: true, colorStr: `indexed:${fg.indexed}` };
  }
  
  return { hasColor: false, colorStr: 'pattern-?' };
}

function getColLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

findColoredCellDifferences().catch(console.error);
