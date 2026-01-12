#!/usr/bin/env node
const ExcelJS = require('exceljs');

async function countFormattedCells() {
    const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    const sheetName = 'DEFENCE&SPACE Aug-2025';
    
    console.log('Zähle formatierte Zellen in Excel...\n');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(sheetName);
    
    let totalCells = 0;
    let cellsWithFont = 0;
    let cellsWithFill = 0;
    let cellsWithBold = 0;
    let cellsWithItalic = 0;
    let cellsWithColor = 0;
    
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber > 100) return; // Nur erste 100 Zeilen
        
        row.eachCell({ includeEmpty: false }, (cell) => {
            totalCells++;
            
            if (cell.font) {
                cellsWithFont++;
                if (cell.font.bold) cellsWithBold++;
                if (cell.font.italic) cellsWithItalic++;
                if (cell.font.color) cellsWithColor++;
            }
            
            if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor) {
                cellsWithFill++;
            }
        });
    });
    
    console.log('Ergebnisse (erste 100 Zeilen):');
    console.log(`  Gesamt Zellen: ${totalCells}`);
    console.log(`  Mit Font-Eigenschaften: ${cellsWithFont}`);
    console.log(`  Mit Bold: ${cellsWithBold}`);
    console.log(`  Mit Italic: ${cellsWithItalic}`);
    console.log(`  Mit Farbe: ${cellsWithColor}`);
    console.log(`  Mit Fill: ${cellsWithFill}`);
    console.log();
    
    // Zeige Beispiele
    console.log('Beispiel-Zellen mit Bold/Italic in Zeilen 2-20:');
    let count = 0;
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber < 2 || rowNumber > 20 || count >= 10) return;
        
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            if (count >= 10) return;
            
            if (cell.font && (cell.font.bold || cell.font.italic)) {
                const colorInfo = cell.font.color?.argb ? ` Farbe: ${cell.font.color.argb}` : '';
                console.log(`Zeile ${rowNumber}, Spalte ${colNumber}: Bold=${cell.font.bold || false}, Italic=${cell.font.italic || false}${colorInfo} - Wert: "${cell.value}"`);
                count++;
            }
        });
    });
}

countFormattedCells();
