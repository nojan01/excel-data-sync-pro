/**
 * Erstellt eine Excel-Testdatei mit verschiedenen Styles
 * zum Testen der Style-Anzeige im DatenExplorer
 */

const XlsxPopulate = require('xlsx-populate');
const path = require('path');

async function createStyleTestFile() {
    const workbook = await XlsxPopulate.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name('Style Tests');

    // === HEADER ===
    sheet.cell('A1').value('Style Test Datei').style({
        bold: true,
        fontSize: 16,
        fontColor: '1F4E79'
    });
    sheet.range('A1:H1').merged(true);

    // === Abschnitt 1: Hintergrundfarben ===
    sheet.cell('A3').value('Hintergrundfarben:').style({ bold: true });
    
    sheet.cell('A4').value('Rot').style({ fill: 'FF0000' });
    sheet.cell('B4').value('Grün').style({ fill: '00FF00' });
    sheet.cell('C4').value('Blau').style({ fill: '0000FF', fontColor: 'FFFFFF' });
    sheet.cell('D4').value('Gelb').style({ fill: 'FFFF00' });
    sheet.cell('E4').value('Orange').style({ fill: 'FFA500' });
    sheet.cell('F4').value('Lila').style({ fill: '800080', fontColor: 'FFFFFF' });
    sheet.cell('G4').value('Cyan').style({ fill: '00FFFF' });
    sheet.cell('H4').value('Pink').style({ fill: 'FF69B4' });

    // === Abschnitt 2: Schriftfarben ===
    sheet.cell('A6').value('Schriftfarben:').style({ bold: true });
    
    sheet.cell('A7').value('Rot').style({ fontColor: 'FF0000' });
    sheet.cell('B7').value('Grün').style({ fontColor: '008000' });
    sheet.cell('C7').value('Blau').style({ fontColor: '0000FF' });
    sheet.cell('D7').value('Orange').style({ fontColor: 'FFA500' });
    sheet.cell('E7').value('Lila').style({ fontColor: '800080' });
    sheet.cell('F7').value('Türkis').style({ fontColor: '008B8B' });

    // === Abschnitt 3: Schriftformatierungen ===
    sheet.cell('A9').value('Schriftformatierungen:').style({ bold: true });
    
    sheet.cell('A10').value('Fett').style({ bold: true });
    sheet.cell('B10').value('Kursiv').style({ italic: true });
    sheet.cell('C10').value('Unterstrichen').style({ underline: true });
    sheet.cell('D10').value('Durchgestrichen').style({ strikethrough: true });
    sheet.cell('E10').value('Fett + Kursiv').style({ bold: true, italic: true });
    sheet.cell('F10').value('Alle Stile').style({ 
        bold: true, 
        italic: true, 
        underline: true 
    });

    // === Abschnitt 4: Schriftgrößen ===
    sheet.cell('A12').value('Schriftgrößen:').style({ bold: true });
    
    sheet.cell('A13').value('8pt').style({ fontSize: 8 });
    sheet.cell('B13').value('10pt').style({ fontSize: 10 });
    sheet.cell('C13').value('12pt').style({ fontSize: 12 });
    sheet.cell('D13').value('14pt').style({ fontSize: 14 });
    sheet.cell('E13').value('16pt').style({ fontSize: 16 });
    sheet.cell('F13').value('20pt').style({ fontSize: 20 });

    // === Abschnitt 5: Kombinationen ===
    sheet.cell('A15').value('Kombinationen:').style({ bold: true });
    
    sheet.cell('A16').value('Warnung').style({ 
        fill: 'FFEB9C', 
        fontColor: '9C5700',
        bold: true 
    });
    sheet.cell('B16').value('Fehler').style({ 
        fill: 'FFC7CE', 
        fontColor: '9C0006',
        bold: true 
    });
    sheet.cell('C16').value('Erfolg').style({ 
        fill: 'C6EFCE', 
        fontColor: '006100',
        bold: true 
    });
    sheet.cell('D16').value('Info').style({ 
        fill: 'BDD7EE', 
        fontColor: '1F4E79',
        bold: true 
    });
    sheet.cell('E16').value('Neutral').style({ 
        fill: 'EDEDED', 
        fontColor: '3F3F3F' 
    });

    // === Abschnitt 6: Ausrichtung ===
    sheet.cell('A18').value('Ausrichtung:').style({ bold: true });
    
    sheet.range('A19:C19').merged(true);
    sheet.cell('A19').value('Links ausgerichtet').style({ 
        horizontalAlignment: 'left',
        fill: 'F0F0F0'
    });
    
    sheet.range('D19:F19').merged(true);
    sheet.cell('D19').value('Zentriert').style({ 
        horizontalAlignment: 'center',
        fill: 'F0F0F0'
    });
    
    sheet.range('G19:I19').merged(true);
    sheet.cell('G19').value('Rechts ausgerichtet').style({ 
        horizontalAlignment: 'right',
        fill: 'F0F0F0'
    });

    // === Abschnitt 7: Tabelle mit gemischten Styles ===
    sheet.cell('A21').value('Beispiel-Tabelle:').style({ bold: true });
    
    // Header-Zeile
    const headers = ['ID', 'Name', 'Status', 'Wert', 'Datum'];
    headers.forEach((header, idx) => {
        sheet.cell(22, idx + 1).value(header).style({
            bold: true,
            fill: '4472C4',
            fontColor: 'FFFFFF',
            horizontalAlignment: 'center'
        });
    });

    // Datenzeilen
    const data = [
        [1, 'Projekt Alpha', 'Aktiv', 1500, '01.01.2026'],
        [2, 'Projekt Beta', 'Pausiert', 2300, '15.01.2026'],
        [3, 'Projekt Gamma', 'Abgeschlossen', 4200, '20.01.2026'],
        [4, 'Projekt Delta', 'Fehler', 800, '25.01.2026'],
    ];

    const statusColors = {
        'Aktiv': { fill: 'C6EFCE', fontColor: '006100' },
        'Pausiert': { fill: 'FFEB9C', fontColor: '9C5700' },
        'Abgeschlossen': { fill: 'BDD7EE', fontColor: '1F4E79' },
        'Fehler': { fill: 'FFC7CE', fontColor: '9C0006' }
    };

    data.forEach((row, rowIdx) => {
        row.forEach((value, colIdx) => {
            const cell = sheet.cell(23 + rowIdx, colIdx + 1);
            cell.value(value);
            
            // Zebra-Streifen
            if (rowIdx % 2 === 1) {
                cell.style({ fill: 'F2F2F2' });
            }
            
            // Status-Spalte einfärben
            if (colIdx === 2) {
                const style = statusColors[value];
                if (style) {
                    cell.style(style);
                }
            }
        });
    });

    // Spaltenbreiten anpassen
    sheet.column('A').width(15);
    sheet.column('B').width(18);
    sheet.column('C').width(15);
    sheet.column('D').width(12);
    sheet.column('E').width(15);
    sheet.column('F').width(15);
    sheet.column('G').width(15);
    sheet.column('H').width(15);
    sheet.column('I').width(15);

    // Datei speichern
    const outputPath = path.join(__dirname, '..', 'test-styles.xlsx');
    await workbook.toFileAsync(outputPath);
    
    console.log(`✅ Style-Testdatei erstellt: ${outputPath}`);
    console.log('\nEnthaltene Style-Tests:');
    console.log('  - Hintergrundfarben (8 Farben)');
    console.log('  - Schriftfarben (6 Farben)');
    console.log('  - Schriftformatierungen (Fett, Kursiv, Unterstrichen, Durchgestrichen)');
    console.log('  - Schriftgrößen (8pt - 20pt)');
    console.log('  - Kombinationen (Warnung, Fehler, Erfolg, Info)');
    console.log('  - Ausrichtung (Links, Zentriert, Rechts)');
    console.log('  - Beispiel-Tabelle mit Zebra-Streifen und Status-Farben');
}

createStyleTestFile().catch(err => {
    console.error('Fehler beim Erstellen der Testdatei:', err);
    process.exit(1);
});
