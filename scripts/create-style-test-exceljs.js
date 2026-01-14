/**
 * Erstellt eine Excel-Testdatei mit ExcelJS
 * mit verschiedenen Styles zum Testen der Style-Anzeige im DatenExplorer
 */

const ExcelJS = require('exceljs');
const path = require('path');

async function createStyleTestFile() {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'MVMS Tool';
    workbook.created = new Date();
    
    const sheet = workbook.addWorksheet('Style Tests');

    // === HEADER (Zeile 1) - Merged Cell ===
    sheet.mergeCells('A1:H1');
    const headerCell = sheet.getCell('A1');
    headerCell.value = 'Style Test Datei (ExcelJS)';
    headerCell.font = { bold: true, size: 16, color: { argb: 'FF1F4E79' } };
    headerCell.alignment = { horizontal: 'center' };

    // === Zeile 2: Hyperlink Test ===
    sheet.getCell('A2').value = 'Links:';
    sheet.getCell('A2').font = { bold: true };
    
    const linkCell = sheet.getCell('D2');
    linkCell.value = { text: 'Google Link', hyperlink: 'https://www.google.com' };
    linkCell.font = { color: { argb: 'FF0000FF' }, underline: true };

    // === Abschnitt 1: Hintergrundfarben (Zeile 3-4) ===
    sheet.getCell('A3').value = 'Hintergrundfarben:';
    sheet.getCell('A3').font = { bold: true };

    const bgColors = [
        { col: 'A', name: 'Rot', color: 'FFFF0000' },
        { col: 'B', name: 'Grün', color: 'FF00FF00' },
        { col: 'C', name: 'Blau', color: 'FF0000FF', fontColor: 'FFFFFFFF' },
        { col: 'D', name: 'Gelb', color: 'FFFFFF00' },
        { col: 'E', name: 'Orange', color: 'FFFFA500' },
        { col: 'F', name: 'Lila', color: 'FF800080', fontColor: 'FFFFFFFF' },
        { col: 'G', name: 'Cyan', color: 'FF00FFFF' },
        { col: 'H', name: 'Pink', color: 'FFFF69B4' }
    ];

    bgColors.forEach(({ col, name, color, fontColor }) => {
        const cell = sheet.getCell(`${col}4`);
        cell.value = name;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
        if (fontColor) {
            cell.font = { color: { argb: fontColor } };
        }
    });

    // === Abschnitt 2: Schriftfarben (Zeile 6-7) ===
    sheet.getCell('A6').value = 'Schriftfarben:';
    sheet.getCell('A6').font = { bold: true };

    const fontColors = [
        { col: 'A', name: 'Rot', color: 'FFFF0000' },
        { col: 'B', name: 'Grün', color: 'FF008000' },
        { col: 'C', name: 'Blau', color: 'FF0000FF' },
        { col: 'D', name: 'Orange', color: 'FFFFA500' },
        { col: 'E', name: 'Lila', color: 'FF800080' },
        { col: 'F', name: 'Türkis', color: 'FF008B8B' }
    ];

    fontColors.forEach(({ col, name, color }) => {
        const cell = sheet.getCell(`${col}7`);
        cell.value = name;
        cell.font = { color: { argb: color } };
    });

    // === Abschnitt 3: Schriftformatierungen (Zeile 9-10) ===
    sheet.getCell('A9').value = 'Schriftformatierungen:';
    sheet.getCell('A9').font = { bold: true };

    sheet.getCell('A10').value = 'Fett';
    sheet.getCell('A10').font = { bold: true };

    sheet.getCell('B10').value = 'Kursiv';
    sheet.getCell('B10').font = { italic: true };

    sheet.getCell('C10').value = 'Unterstrichen';
    sheet.getCell('C10').font = { underline: true };

    sheet.getCell('D10').value = 'Durchgestrichen';
    sheet.getCell('D10').font = { strike: true };

    sheet.getCell('E10').value = 'Fett + Kursiv';
    sheet.getCell('E10').font = { bold: true, italic: true };

    sheet.getCell('F10').value = 'Alle Stile';
    sheet.getCell('F10').font = { bold: true, italic: true, underline: true };

    // === Abschnitt 4: Schriftgrößen (Zeile 12-13) ===
    sheet.getCell('A12').value = 'Schriftgrößen:';
    sheet.getCell('A12').font = { bold: true };

    const sizes = [8, 10, 12, 14, 16, 20];
    sizes.forEach((size, index) => {
        const col = String.fromCharCode(65 + index); // A, B, C, ...
        const cell = sheet.getCell(`${col}13`);
        cell.value = `${size}pt`;
        cell.font = { size: size };
    });

    // === Abschnitt 5: Rich Text (Zeile 15) ===
    sheet.getCell('A15').value = 'Rich Text:';
    sheet.getCell('A15').font = { bold: true };

    // Rich Text mit unterschiedlichen Formatierungen pro Buchstabe
    sheet.getCell('A16').value = {
        richText: [
            { text: 'R', font: { bold: true, color: { argb: 'FFFF0000' }, size: 14 } },
            { text: 'E', font: { italic: true, color: { argb: 'FF00FF00' }, size: 16 } },
            { text: 'I', font: { underline: true, color: { argb: 'FF0000FF' }, size: 18 } },
            { text: 'C', font: { bold: true, italic: true, color: { argb: 'FFFFA500' }, size: 20 } },
            { text: 'H', font: { strike: true, color: { argb: 'FF800080' }, size: 22 } }
        ]
    };

    sheet.getCell('B16').value = {
        richText: [
            { text: 'Multi', font: { bold: true, color: { argb: 'FF1F4E79' } } },
            { text: '-', font: { size: 11 } },
            { text: 'Format', font: { italic: true, color: { argb: 'FF9C0006' } } },
            { text: ' Text', font: { underline: true, size: 14 } }
        ]
    };

    // === Abschnitt 6: Kombinationen (Zeile 18-19) ===
    sheet.getCell('A18').value = 'Kombinationen:';
    sheet.getCell('A18').font = { bold: true };

    // Warnung
    sheet.getCell('A19').value = 'Warnung';
    sheet.getCell('A19').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
    sheet.getCell('A19').font = { bold: true, color: { argb: 'FF9C5700' } };

    // Fehler
    sheet.getCell('B19').value = 'Fehler';
    sheet.getCell('B19').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
    sheet.getCell('B19').font = { bold: true, color: { argb: 'FF9C0006' } };

    // Erfolg
    sheet.getCell('C19').value = 'Erfolg';
    sheet.getCell('C19').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
    sheet.getCell('C19').font = { bold: true, color: { argb: 'FF006100' } };

    // Info
    sheet.getCell('D19').value = 'Info';
    sheet.getCell('D19').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
    sheet.getCell('D19').font = { bold: true, color: { argb: 'FF1F4E79' } };

    // Neutral
    sheet.getCell('E19').value = 'Neutral';
    sheet.getCell('E19').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDEDED' } };
    sheet.getCell('E19').font = { color: { argb: 'FF333333' } };

    // === Abschnitt 7: Merged Cells (Zeile 21-22) ===
    sheet.getCell('A21').value = 'Merged Cells:';
    sheet.getCell('A21').font = { bold: true };

    // Horizontal merge
    sheet.mergeCells('A22:C22');
    sheet.getCell('A22').value = 'Horizontal Merge (A-C)';
    sheet.getCell('A22').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } };
    sheet.getCell('A22').alignment = { horizontal: 'center' };

    sheet.mergeCells('D22:F22');
    sheet.getCell('D22').value = 'Merge 2 (D-F)';
    sheet.getCell('D22').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0FF' } };
    sheet.getCell('D22').alignment = { horizontal: 'center' };

    sheet.mergeCells('G22:I22');
    sheet.getCell('G22').value = 'Merge 3 (G-I)';
    sheet.getCell('G22').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE0E0' } };
    sheet.getCell('G22').alignment = { horizontal: 'center' };

    // === Abschnitt 8: Tabellen-Style (Zeile 24-28) ===
    sheet.getCell('A24').value = 'Tabellen-Demo:';
    sheet.getCell('A24').font = { bold: true };

    // Header-Zeile
    const tableHeaders = ['ID', 'Name', 'Status', 'Datum', 'Wert'];
    tableHeaders.forEach((header, index) => {
        const col = String.fromCharCode(65 + index);
        const cell = sheet.getCell(`${col}25`);
        cell.value = header;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
        cell.alignment = { horizontal: 'center' };
    });

    // Datenzeilen
    const tableData = [
        ['1', 'Projekt Alpha', 'Aktiv', '2026-01-01', '1.234,56'],
        ['2', 'Projekt Beta', 'Pausiert', '2026-01-05', '987,65'],
        ['3', 'Projekt Gamma', 'Abgeschlossen', '2026-01-10', '5.432,10']
    ];

    tableData.forEach((rowData, rowIndex) => {
        rowData.forEach((value, colIndex) => {
            const col = String.fromCharCode(65 + colIndex);
            const cell = sheet.getCell(`${col}${26 + rowIndex}`);
            cell.value = value;
            
            // Alternating row colors
            if (rowIndex % 2 === 0) {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
            }
            
            // Status-Farben
            if (colIndex === 2) {
                if (value === 'Aktiv') {
                    cell.font = { color: { argb: 'FF006100' } };
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
                } else if (value === 'Pausiert') {
                    cell.font = { color: { argb: 'FF9C5700' } };
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
                } else if (value === 'Abgeschlossen') {
                    cell.font = { color: { argb: 'FF1F4E79' } };
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
                }
            }
        });
    });

    // Spaltenbreiten anpassen
    sheet.columns = [
        { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 },
        { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }
    ];

    // Datei speichern
    const outputPath = path.join(__dirname, '..', 'test-styles-exceljs.xlsx');
    await workbook.xlsx.writeFile(outputPath);
    console.log(`✅ Testdatei erstellt: ${outputPath}`);
    
    // Auch auf Desktop speichern
    const desktopPath = path.join(require('os').homedir(), 'Desktop', 'test-styles-exceljs.xlsx');
    await workbook.xlsx.writeFile(desktopPath);
    console.log(`✅ Kopie auf Desktop: ${desktopPath}`);
}

createStyleTestFile().catch(console.error);
