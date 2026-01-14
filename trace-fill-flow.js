const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Simuliere was passiert wenn Spalte A gelöscht wird
async function traceFillFlow() {
  const originalPath = '/Users/nojan/Desktop/test-styles-exceljs.xlsx';
  
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(originalPath);
  const ws = workbook.getWorksheet(1);
  
  console.log('=== ORIGINAL MERGED CELLS ===');
  console.log(ws.model.merges);
  
  console.log('\n=== ORIGINAL FILLS ZEILE 20-22, SPALTE G-I ===');
  for (let row = 20; row <= 22; row++) {
    for (let col = 7; col <= 9; col++) { // G=7, H=8, I=9
      const cell = ws.getCell(row, col);
      const colLetter = String.fromCharCode(64 + col);
      const fill = cell.fill;
      let fillStr = 'keine';
      if (fill && fill.type === 'pattern' && fill.fgColor) {
        fillStr = fill.fgColor.argb || fill.fgColor.theme || 'unbekannt';
      }
      console.log(`  ${colLetter}${row}: Fill=${fillStr}, Value="${cell.value || ''}"`);
    }
  }
  
  console.log('\n=== NACH spliceColumns(1, 1) ===');
  // Speichere Merged-Info vor dem Löschen
  const mergedRanges = [...ws.model.merges];
  console.log('Merged Ranges vor splice:', mergedRanges);
  
  // Unmerge alle
  for (const range of mergedRanges) {
    try { ws.unMergeCells(range); } catch(e) {}
  }
  
  // Speichere Master-Zellen Info für G20:I22
  const g20 = ws.getCell('G20');
  console.log('\nMaster G20 vor splice:');
  console.log('  Value:', g20.value);
  console.log('  Fill:', g20.fill?.fgColor?.argb || 'keine');
  
  // Splice
  ws.spliceColumns(1, 1);
  
  console.log('\n=== NACH SPLICE - ZEILE 20-22, SPALTE F-H (ehemals G-I) ===');
  for (let row = 20; row <= 22; row++) {
    for (let col = 6; col <= 8; col++) { // F=6, G=7, H=8
      const cell = ws.getCell(row, col);
      const colLetter = String.fromCharCode(64 + col);
      const fill = cell.fill;
      let fillStr = 'keine';
      if (fill && fill.type === 'pattern' && fill.fgColor) {
        fillStr = fill.fgColor.argb || fill.fgColor.theme || 'unbekannt';
      }
      console.log(`  ${colLetter}${row}: Fill=${fillStr}, Value="${cell.value || ''}"`);
    }
  }
  
  console.log('\n=== DAS PROBLEM ===');
  console.log('Im Original haben G22, H22, I22 rosa Fill (FFE0E0)');
  console.log('Nach splice werden diese zu F22, G22, H22');
  console.log('Wenn wir F20:H22 mergen, übernimmt die Master-Zelle F20 den Fill der ersten Zelle');
  console.log('Aber applyMissingFills setzt danach auf Basis der cellStyles die Fills,');
  console.log('und die cellStyles haben für F22, G22, H22 den rosa Fill (vom Frontend angepasst)');
  
  console.log('\n=== LÖSUNG ===');
  console.log('Bei Merged Cells müssen wir die cellStyles-Einträge für ALLE Zellen im Merge-Bereich entfernen,');
  console.log('damit applyMissingFills sie nicht überschreibt.');
}

traceFillFlow();
