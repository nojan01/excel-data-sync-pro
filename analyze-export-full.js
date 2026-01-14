const AdmZip = require('adm-zip');
const ExcelJS = require('exceljs');

async function analyzeExport() {
    const originalPath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    const exportPath = '/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    console.log('=== STYLE ANALYSE ===\n');
    
    // XML-Analyse
    const zipOrig = new AdmZip(originalPath);
    const zipExp = new AdmZip(exportPath);
    
    const stylesOrig = zipOrig.readAsText('xl/styles.xml');
    const stylesExp = zipExp.readAsText('xl/styles.xml');
    
    // Fills zählen
    const fillsOrigCount = (stylesOrig.match(/<fill>/g) || []).length;
    const fillsExpCount = (stylesExp.match(/<fill>/g) || []).length;
    console.log('Fills - Original:', fillsOrigCount, ', Export:', fillsExpCount);
    
    // CellXf zählen
    const xfOrigCount = (stylesOrig.match(/<xf /g) || []).length;
    const xfExpCount = (stylesExp.match(/<xf /g) || []).length;
    console.log('CellXf - Original:', xfOrigCount, ', Export:', xfExpCount);
    
    // Theme-Colors prüfen
    const themeOrigCount = (stylesOrig.match(/theme="/g) || []).length;
    const themeExpCount = (stylesExp.match(/theme="/g) || []).length;
    console.log('Theme refs - Original:', themeOrigCount, ', Export:', themeExpCount);
    
    console.log('\n=== CONDITIONAL FORMATTING ===\n');
    
    const sheetOrig = zipOrig.readAsText('xl/worksheets/sheet1.xml');
    const sheetExp = zipExp.readAsText('xl/worksheets/sheet1.xml');
    
    // CF zählen
    const cfOrigCount = (sheetOrig.match(/<conditionalFormatting/g) || []).length;
    const cfExpCount = (sheetExp.match(/<conditionalFormatting/g) || []).length;
    console.log('Conditional Formatting Blöcke - Original:', cfOrigCount, ', Export:', cfExpCount);
    
    // CF Rules zählen
    const cfRulesOrigCount = (sheetOrig.match(/<cfRule/g) || []).length;
    const cfRulesExpCount = (sheetExp.match(/<cfRule/g) || []).length;
    console.log('CF Rules - Original:', cfRulesOrigCount, ', Export:', cfRulesExpCount);
    
    console.log('\n=== EXCELJS STYLE-ANALYSE ===\n');
    
    // ExcelJS
    const wbOrig = new ExcelJS.Workbook();
    const wbExp = new ExcelJS.Workbook();
    await wbOrig.xlsx.readFile(originalPath);
    await wbExp.xlsx.readFile(exportPath);
    
    const wsOrig = wbOrig.getWorksheet(1);
    const wsExp = wbExp.getWorksheet(1);
    
    // Stichprobe Zeile 2-10
    console.log('Stichprobe Zeilen 2-10, Spalte A-E:\n');
    
    for (let row = 2; row <= 10; row++) {
        const origStyles = [];
        const expStyles = [];
        
        for (let col = 1; col <= 5; col++) {
            const cellOrig = wsOrig.getCell(row, col);
            const cellExp = wsExp.getCell(row, col);
            
            const origHasFill = cellOrig.style?.fill?.fgColor || cellOrig.style?.fill?.bgColor;
            const expHasFill = cellExp.style?.fill?.fgColor || cellExp.style?.fill?.bgColor;
            
            origStyles.push(origHasFill ? 'F' : '-');
            expStyles.push(expHasFill ? 'F' : '-');
        }
        
        console.log(`Zeile ${row}: Orig [${origStyles.join('')}] Export [${expStyles.join('')}]`);
    }
    
    // CF Details
    console.log('\n=== CF DETAILS ===\n');
    
    console.log('Original CF:');
    const cfOrigMatches = sheetOrig.match(/<conditionalFormatting[^>]*sqref="([^"]+)"/g) || [];
    cfOrigMatches.slice(0, 5).forEach(m => {
        const sqref = m.match(/sqref="([^"]+)"/)?.[1];
        console.log('  sqref:', sqref?.substring(0, 50));
    });
    
    console.log('\nExport CF:');
    const cfExpMatches = sheetExp.match(/<conditionalFormatting[^>]*sqref="([^"]+)"/g) || [];
    cfExpMatches.slice(0, 5).forEach(m => {
        const sqref = m.match(/sqref="([^"]+)"/)?.[1];
        console.log('  sqref:', sqref?.substring(0, 50));
    });
}

analyzeExport().catch(console.error);
