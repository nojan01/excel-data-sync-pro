const ExcelJS = require('exceljs');

async function checkCFForLM() {
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
  
  // Finde alle CF-Regeln die Spalte L oder M betreffen
  console.log('\n=== Original: CF-Regeln die Spalte L oder M betreffen ===\n');
  
  const origCFs = sheetOrig.conditionalFormattings || [];
  origCFs.forEach((cf, idx) => {
    const ref = cf.ref;
    if (refIncludesColumn(ref, 'L') || refIncludesColumn(ref, 'M')) {
      console.log(`CF #${idx}: ref="${ref}"`);
      if (cf.rules) {
        cf.rules.forEach((rule, rIdx) => {
          console.log(`  Rule ${rIdx}: type=${rule.type}, priority=${rule.priority}`);
          if (rule.formulae) console.log(`    formulae: ${JSON.stringify(rule.formulae)}`);
          if (rule.style && rule.style.fill) {
            console.log(`    fill: ${JSON.stringify(rule.style.fill)}`);
          }
        });
      }
    }
  });
  
  console.log('\n=== Export: CF-Regeln die Spalte K, L oder M betreffen ===\n');
  // Nach Löschung von A: Orig L→Export K, Orig M→Export L
  
  const exportCFs = sheetExport.conditionalFormattings || [];
  exportCFs.forEach((cf, idx) => {
    const ref = cf.ref;
    if (refIncludesColumn(ref, 'K') || refIncludesColumn(ref, 'L') || refIncludesColumn(ref, 'M')) {
      console.log(`CF #${idx}: ref="${ref}"`);
      if (cf.rules) {
        cf.rules.forEach((rule, rIdx) => {
          console.log(`  Rule ${rIdx}: type=${rule.type}, priority=${rule.priority}`);
          if (rule.formulae) console.log(`    formulae: ${JSON.stringify(rule.formulae)}`);
          if (rule.style && rule.style.fill) {
            console.log(`    fill: ${JSON.stringify(rule.style.fill)}`);
          }
        });
      }
    }
  });
  
  // Prüfe ob alle CF refs richtig angepasst wurden
  console.log('\n=== Vergleiche CF-Verschiebung ===\n');
  
  // Suche CF mit Spalte L im Original
  const origLCFs = [];
  origCFs.forEach((cf, idx) => {
    const ref = cf.ref;
    // Parse the start column from the ref
    const match = ref.match(/^([A-Z]+)/);
    if (match) {
      const startCol = match[1];
      origLCFs.push({ idx, ref, startCol });
    }
  });
  
  // Zeige Mapping
  console.log('Original CF Spalten-Starts (erste 20):');
  origLCFs.slice(0, 20).forEach(cf => {
    const expectedExportCol = shiftColumnLeft(cf.startCol);
    const exportCF = exportCFs[cf.idx];
    const exportMatch = exportCF?.ref?.match(/^([A-Z]+)/);
    const actualExportCol = exportMatch ? exportMatch[1] : 'N/A';
    
    const correct = expectedExportCol === actualExportCol;
    console.log(`  CF #${cf.idx}: ${cf.startCol} → erwartet ${expectedExportCol}, tatsächlich ${actualExportCol} ${correct ? '✓' : '✗'}`);
  });
}

function refIncludesColumn(ref, colLetter) {
  // Parse A1:Z100 or just A1
  const match = ref.match(/^([A-Z]+)(\d+):?([A-Z]+)?/);
  if (!match) return false;
  
  const startCol = columnToNumber(match[1]);
  const endCol = match[3] ? columnToNumber(match[3]) : startCol;
  const targetCol = columnToNumber(colLetter);
  
  return targetCol >= startCol && targetCol <= endCol;
}

function columnToNumber(col) {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
}

function shiftColumnLeft(col) {
  const num = columnToNumber(col);
  if (num <= 1) return ''; // A wird gelöscht
  return numberToColumn(num - 1);
}

function numberToColumn(num) {
  let col = '';
  while (num > 0) {
    const mod = (num - 1) % 26;
    col = String.fromCharCode(65 + mod) + col;
    num = Math.floor((num - 1) / 26);
  }
  return col;
}

checkCFForLM().catch(console.error);
