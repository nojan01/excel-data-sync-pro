const ExcelJS = require('exceljs');

async function compareLMFocused() {
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
  
  // Nach Spaltenlöschung:
  // Export L (12) = Original M (13)
  // Export M (13) = Original N (14)
  
  console.log('\n=== Vergleiche: Export Spalte L (12) mit Original Spalte M (13) ===');
  
  // Prüfe erste 20 Zeilen wo es Unterschiede gibt
  let diffCount = 0;
  const samplesToShow = 20;
  const differences = [];
  
  // Spalte L in Export (sollte = Spalte M in Original)
  for (let row = 1; row <= 100; row++) {
    const origCell = sheetOrig.getCell(row, 13); // M = 13
    const exportCell = sheetExport.getCell(row, 12); // L = 12
    
    const origFill = getFillColor(origCell);
    const exportFill = getFillColor(exportCell);
    
    if (origFill !== exportFill) {
      differences.push({
        row,
        column: 'L (Orig M)',
        origFill,
        exportFill,
        origValue: getCellValue(origCell),
        exportValue: getCellValue(exportCell)
      });
      diffCount++;
      if (diffCount >= samplesToShow) break;
    }
  }
  
  if (differences.length === 0) {
    // Prüfe weitere Zeilen
    for (let row = 100; row <= 500; row++) {
      const origCell = sheetOrig.getCell(row, 13);
      const exportCell = sheetExport.getCell(row, 12);
      
      const origFill = getFillColor(origCell);
      const exportFill = getFillColor(exportCell);
      
      if (origFill !== exportFill) {
        differences.push({
          row,
          column: 'L (Orig M)',
          origFill,
          exportFill,
          origValue: getCellValue(origCell),
          exportValue: getCellValue(exportCell)
        });
        diffCount++;
        if (diffCount >= samplesToShow) break;
      }
    }
  }
  
  if (differences.length === 0) {
    console.log('Keine Unterschiede in Spalte L (erste 500 Zeilen)');
    
    // Prüfe ob die Export-Spalte L vielleicht die Original-Spalte L ist (nicht verschoben)
    console.log('\n=== Prüfe ob Export L = Original L (nicht verschoben) ===');
    let matchesOrigL = 0;
    let matchesOrigM = 0;
    
    for (let row = 1; row <= 100; row++) {
      const exportCell = sheetExport.getCell(row, 12);
      const origCellL = sheetOrig.getCell(row, 12);
      const origCellM = sheetOrig.getCell(row, 13);
      
      const exportFill = getFillColor(exportCell);
      const origLFill = getFillColor(origCellL);
      const origMFill = getFillColor(origCellM);
      
      if (exportFill === origLFill) matchesOrigL++;
      if (exportFill === origMFill) matchesOrigM++;
    }
    
    console.log(`Export L passt zu Original L: ${matchesOrigL}/100`);
    console.log(`Export L passt zu Original M: ${matchesOrigM}/100`);
    
  } else {
    console.log(`${differences.length} Unterschiede gefunden:\n`);
    differences.forEach(d => {
      console.log(`Zeile ${d.row}:`);
      console.log(`  Original M: "${d.origValue}" -> Fill: ${d.origFill}`);
      console.log(`  Export L:   "${d.exportValue}" -> Fill: ${d.exportFill}`);
      console.log('');
    });
  }
  
  // Prüfe auch die Werte ob sie stimmen
  console.log('\n=== Werte-Check: Stimmen die Werte überein? ===');
  let valueMatches = 0;
  let valueMismatches = 0;
  const valueDiffs = [];
  
  for (let row = 1; row <= 100; row++) {
    const origCell = sheetOrig.getCell(row, 13); // M
    const exportCell = sheetExport.getCell(row, 12); // L
    
    const origVal = getCellValue(origCell);
    const exportVal = getCellValue(exportCell);
    
    if (origVal === exportVal) {
      valueMatches++;
    } else {
      valueMismatches++;
      if (valueDiffs.length < 10) {
        valueDiffs.push({ row, origVal, exportVal });
      }
    }
  }
  
  console.log(`Werte Match: ${valueMatches}/100, Mismatch: ${valueMismatches}/100`);
  
  if (valueDiffs.length > 0) {
    console.log('\nErste Wert-Unterschiede:');
    valueDiffs.forEach(d => {
      console.log(`  Zeile ${d.row}: Orig M="${d.origVal}" vs Export L="${d.exportVal}"`);
    });
  }
  
  // Zeige Sample von Fills in beiden Spalten
  console.log('\n=== Sample der ersten 15 Zeilen ===');
  console.log('Zeile | Orig L Fill | Orig M Fill | Export L Fill');
  console.log('------|-------------|-------------|---------------');
  
  for (let row = 1; row <= 15; row++) {
    const origL = sheetOrig.getCell(row, 12);
    const origM = sheetOrig.getCell(row, 13);
    const exportL = sheetExport.getCell(row, 12);
    
    const origLFill = getFillColor(origL);
    const origMFill = getFillColor(origM);
    const exportLFill = getFillColor(exportL);
    
    const match = exportLFill === origMFill ? '✓' : '✗';
    console.log(`${row.toString().padStart(5)} | ${origLFill.padEnd(11)} | ${origMFill.padEnd(11)} | ${exportLFill.padEnd(13)} ${match}`);
  }
}

function getFillColor(cell) {
  if (!cell.fill) return 'none';
  if (cell.fill.type === 'pattern') {
    const fg = cell.fill.fgColor;
    if (!fg) return 'pattern-empty';
    if (fg.argb) return fg.argb;
    if (fg.theme !== undefined) return `theme:${fg.theme}`;
    if (fg.indexed !== undefined) return `idx:${fg.indexed}`;
    return 'pattern-?';
  }
  return cell.fill.type || 'unknown';
}

function getCellValue(cell) {
  const val = cell.value;
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    if (val.richText) return val.richText.map(r => r.text).join('');
    if (val.text) return val.text;
    if (val.result !== undefined) return String(val.result);
    return '[obj]';
  }
  return String(val).substring(0, 30);
}

compareLMFocused().catch(console.error);
