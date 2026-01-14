const ExcelJS = require('exceljs');

async function compare() {
  const original = new ExcelJS.Workbook();
  const exported = new ExcelJS.Workbook();
  
  await original.xlsx.readFile('/Users/nojan/Desktop/AutoFilter_Test.xlsx');
  await exported.xlsx.readFile('/Users/nojan/Desktop/Export_AutoFilter_Test.xlsx');
  
  const origSheet = original.worksheets[0];
  const expSheet = exported.worksheets[0];
  
  console.log('=== VERGLEICH: Original vs Export ===\n');
  console.log('Sheet-Namen:', origSheet.name, 'vs', expSheet.name);
  
  // Header-Zeile vergleichen (Zeile 1)
  console.log('\nHEADER (Zeile 1):');
  for (let col = 1; col <= 4; col++) {
    const origCell = origSheet.getCell(1, col);
    const expCell = expSheet.getCell(1, col);
    
    console.log(`Spalte ${col} (${origCell.value}):`);
    console.log('  Orig Fill:', JSON.stringify(origCell.fill));
    console.log('  Exp Fill: ', JSON.stringify(expCell.fill));
    
    const match = JSON.stringify(origCell.fill) === JSON.stringify(expCell.fill);
    console.log('  Match:', match ? '✓' : '✗ UNTERSCHIED');
  }
  
  // AutoFilter vergleichen
  console.log('\n=== AUTOFILTER ===');
  console.log('Original:', origSheet.autoFilter);
  console.log('Export:  ', expSheet.autoFilter);
}

compare().catch(console.error);
