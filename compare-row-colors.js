const ExcelJS = require('exceljs');

async function compareRowColors() {
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
  
  // Zeige alle Spalten mit Farbe für Zeile 1
  console.log('\n=== Zeile 1: Alle Spalten mit Farben ===\n');
  
  console.log('ORIGINAL:');
  for (let col = 1; col <= 70; col++) {
    const cell = sheetOrig.getCell(1, col);
    const fill = getDetailedFill(cell);
    if (fill.hasColor) {
      const val = getCellValue(cell);
      console.log(`  ${getColLetter(col)} (${col}): ${fill.colorStr} - "${val}"`);
    }
  }
  
  console.log('\nEXPORT:');
  for (let col = 1; col <= 70; col++) {
    const cell = sheetExport.getCell(1, col);
    const fill = getDetailedFill(cell);
    if (fill.hasColor) {
      const val = getCellValue(cell);
      console.log(`  ${getColLetter(col)} (${col}): ${fill.colorStr} - "${val}"`);
    }
  }
  
  // Suche nach Zeilen mit vielen Farben (wahrscheinlich bedingte Formatierung)
  console.log('\n=== Suche Zeilen mit vielen Farbzellen (Zeilen 2-100) ===\n');
  
  for (let row = 2; row <= 100; row++) {
    const origColors = [];
    const exportColors = [];
    
    for (let col = 1; col <= 61; col++) {
      const origCell = sheetOrig.getCell(row, col);
      const exportCell = sheetExport.getCell(row, col);
      
      const origFill = getDetailedFill(origCell);
      const exportFill = getDetailedFill(exportCell);
      
      if (origFill.hasColor) {
        origColors.push({ col, letter: getColLetter(col), color: origFill.colorStr });
      }
      if (exportFill.hasColor) {
        exportColors.push({ col, letter: getColLetter(col), color: exportFill.colorStr });
      }
    }
    
    if (origColors.length > 0 || exportColors.length > 0) {
      console.log(`Zeile ${row}:`);
      console.log(`  Original: ${origColors.map(c => c.letter).join(', ') || 'keine'}`);
      console.log(`  Export:   ${exportColors.map(c => c.letter).join(', ') || 'keine'}`);
      
      // Check if shifted correctly
      const expectedExport = origColors
        .filter(c => c.col > 1) // A wird gelöscht
        .map(c => getColLetter(c.col - 1)); // Alle anderen um 1 nach links
      
      const actualExport = exportColors.map(c => c.letter);
      const isCorrect = JSON.stringify(expectedExport.sort()) === JSON.stringify(actualExport.sort());
      
      console.log(`  Erwartet: ${expectedExport.join(', ') || 'keine'}`);
      console.log(`  Korrekt: ${isCorrect ? '✓' : '✗'}`);
      console.log('');
    }
  }
  
  // Jetzt schauen was der Benutzer mit "L und M" meint - vielleicht im Export sichtbar
  console.log('\n=== Vergleiche Export-Zeilen 2-20 speziell für L/M ===\n');
  
  for (let row = 2; row <= 20; row++) {
    // Export L (12) sollte = Original M (13)
    const exportL = sheetExport.getCell(row, 12);
    const origM = sheetOrig.getCell(row, 13);
    
    // Export M (13) sollte = Original N (14)
    const exportM = sheetExport.getCell(row, 13);
    const origN = sheetOrig.getCell(row, 14);
    
    const exportLFill = getDetailedFill(exportL);
    const origMFill = getDetailedFill(origM);
    const exportMFill = getDetailedFill(exportM);
    const origNFill = getDetailedFill(origN);
    
    // Auch prüfen ob Export L fälschlicherweise = Original L ist (nicht verschoben)
    const origL = sheetOrig.getCell(row, 12);
    const origLFill = getDetailedFill(origL);
    
    if (exportLFill.colorStr !== origMFill.colorStr || exportMFill.colorStr !== origNFill.colorStr) {
      console.log(`Zeile ${row}:`);
      console.log(`  Orig L:  "${getCellValue(origL)}" ${origLFill.colorStr}`);
      console.log(`  Orig M:  "${getCellValue(origM)}" ${origMFill.colorStr}`);
      console.log(`  Orig N:  "${getCellValue(origN)}" ${origNFill.colorStr}`);
      console.log(`  Export L: "${getCellValue(exportL)}" ${exportLFill.colorStr}`);
      console.log(`  Export M: "${getCellValue(exportM)}" ${exportMFill.colorStr}`);
      console.log('');
    }
  }
}

function getDetailedFill(cell) {
  if (!cell.fill) return { hasColor: false, colorStr: 'none' };
  if (cell.fill.type !== 'pattern') return { hasColor: false, colorStr: cell.fill.type };
  
  const fg = cell.fill.fgColor;
  if (!fg) return { hasColor: false, colorStr: 'pattern-empty' };
  
  if (fg.argb) {
    if (fg.argb === 'FFFFFFFF' || fg.argb === '00000000') {
      return { hasColor: false, colorStr: 'white' };
    }
    return { hasColor: true, colorStr: `argb:${fg.argb}` };
  }
  if (fg.theme !== undefined) {
    return { hasColor: true, colorStr: `theme:${fg.theme}${fg.tint ? `,tint:${fg.tint.toFixed(2)}` : ''}` };
  }
  if (fg.indexed !== undefined) {
    if (fg.indexed === 64) return { hasColor: false, colorStr: 'indexed:64' };
    return { hasColor: true, colorStr: `indexed:${fg.indexed}` };
  }
  
  return { hasColor: false, colorStr: 'pattern-?' };
}

function getCellValue(cell) {
  const val = cell.value;
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    if (val.richText) return val.richText.map(r => r.text).join('');
    if (val.text) return val.text;
    return '[obj]';
  }
  return String(val).substring(0, 30);
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

compareRowColors().catch(console.error);
