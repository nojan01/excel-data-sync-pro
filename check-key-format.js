const ExcelJS = require('exceljs');
const path = require('path');

// Debug: Prüfe das cellStyles Key-Format
async function checkKeyFormat() {
  const originalPath = '/Users/nojan/Desktop/test-styles-exceljs.xlsx';
  
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(originalPath);
  const ws = workbook.getWorksheet(1);
  
  console.log('=== PRÜFE cellStyles KEY FORMAT ===');
  console.log('\nMerged Cell G20:I22 nach Spalte A löschen wird zu F20:H22');
  console.log('Excel-Zeilen: 20-22, Excel-Spalten: F=6, G=7, H=8');
  
  // Simuliere cellStyles wie sie vom Frontend kommen würden
  // Das Frontend passt die Indizes an: G(7)->F(6), H(8)->G(7), I(9)->H(8)
  // Ursprünglich: G22, H22, I22 mit rosa Fill
  // Nach Anpassung: F22, G22, H22 mit rosa Fill
  
  console.log('\n=== MÖGLICHE KEY-FORMATE für F22 ===');
  
  // Zeile 22, Spalte F (6)
  const excelRow = 22;
  const excelCol = 6;
  
  // Format 1: 1-basiert wie Excel
  console.log(`Format 1 (1-basiert): "${excelRow}-${excelCol}" = "22-6"`);
  
  // Format 2: 0-basiert für Datenzeilen (Zeile 2 = Index 0)
  console.log(`Format 2 (Datenzeile): "${excelRow - 2}-${excelCol - 1}" = "20-5"`);
  
  // Format 3: Komplett 0-basiert
  console.log(`Format 3 (0-basiert): "${excelRow - 1}-${excelCol - 1}" = "21-5"`);
  
  console.log('\n=== applyMissingFills VERWENDET FOLGENDES FORMAT ===');
  console.log('Schauen wir in exceljs-writer.js...');
}

checkKeyFormat();
