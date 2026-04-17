// Testdaten-Generator für Serial-Check Feature
// Erzeugt 1 Soll- und 3 Ist-Listen
//
// Verwendung:
//   node scripts/generate-test-data.js                       -> Standard: <repo>/test-data
//   node scripts/generate-test-data.js ./mein-ordner         -> relativ
//   node scripts/generate-test-data.js /absoluter/pfad       -> absolut
//   node scripts/generate-test-data.js --out=~/Desktop/test  -> mit Tilde
//   node scripts/generate-test-data.js -o D:\Testdaten       -> kurz
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');

function parseOutDir() {
    const args = process.argv.slice(2);
    let target = null;
    for (let i = 0; i < args.length; i++) {
        const a = args[i];
        if (a === '-o' || a === '--out') {
            target = args[i + 1];
            i++;
        } else if (a.startsWith('--out=')) {
            target = a.slice('--out='.length);
        } else if (!a.startsWith('-') && !target) {
            target = a;
        }
    }
    if (!target) {
        return path.join(__dirname, '..', 'test-data');
    }
    // ~ expandieren
    if (target.startsWith('~')) {
        target = path.join(os.homedir(), target.slice(1));
    }
    return path.resolve(target);
}

const OUT_DIR = parseOutDir();
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });
console.log('→ Ausgabeverzeichnis:', OUT_DIR);

// ----- Soll-Liste: 30 Server (SN0001 ... SN0030) -----
const sollServers = [];
for (let i = 1; i <= 30; i++) {
    sollServers.push({
        sn: 'SN' + String(i).padStart(4, '0'),
        model: i % 3 === 0 ? 'Dell R750' : i % 3 === 1 ? 'HPE DL380' : 'Lenovo SR650',
        standort: i <= 10 ? 'Abt. IT' : i <= 20 ? 'Abt. Buchhaltung' : 'Abt. Produktion',
        inbetriebnahme: `202${3 + (i % 3)}-0${1 + (i % 9)}-15`
    });
}

// ----- Ist-Listen: absichtlich mit Lücken + unterschiedlichen Spaltennamen + Varianten -----
// IT hat: 1-8   (fehlen: SN0009, SN0010)
// Buchhaltung hat: 11-18  (fehlen: SN0019, SN0020)
// Produktion hat: 21-28, 30  (fehlt: SN0029)
// → Gesamt fehlen in ALLEN Ist-Listen: SN0009, SN0010, SN0019, SN0020, SN0029

// Variationen einbauen, die das Normalisieren testen:
//  - führende Nullen ("0SN0003" etc. wäre anders — besser nur Zahl-SNs):
//    wir nutzen zusätzlich Varianten "sn0003" (lowercase), " SN0004 " (whitespace),
//    "0000SN0005" (führende Nullen am Gesamtstring)
const ist1 = [ // Abt. IT — Spaltenname "Serial No."
    ['SN0001', 'Dell R750',  'Server-Raum 1', 'aktiv'],
    ['sn0002', 'HPE DL380',  'Server-Raum 1', 'aktiv'],             // lowercase
    [' SN0003 ', 'Lenovo SR650', 'Server-Raum 2', 'aktiv'],         // whitespace
    ['0000SN0004', 'Dell R750', 'Server-Raum 2', 'aktiv'],          // führende Nullen
    ['SN0005', 'HPE DL380',  'Server-Raum 1', 'aktiv'],
    ['SN0006', 'Lenovo SR650', 'Server-Raum 1', 'aktiv'],
    ['SN0007', 'Dell R750',  'Server-Raum 2', 'aktiv'],
    ['SN0008', 'HPE DL380',  'Server-Raum 2', 'aktiv'],
    // SN0009, SN0010 fehlen absichtlich
];

const ist2 = [ // Abt. Buchhaltung — Spaltenname "S/N"
    ['SN0011', 'Server-BH-01', 'Erdgeschoss'],
    ['SN0012', 'Server-BH-02', 'Erdgeschoss'],
    ['SN0013', 'Server-BH-03', '1. OG'],
    ['SN0014', 'Server-BH-04', '1. OG'],
    ['SN0015', 'Server-BH-05', '2. OG'],
    ['SN0016', 'Server-BH-06', '2. OG'],
    ['SN0017', 'Server-BH-07', '2. OG'],
    ['SN0018', 'Server-BH-08', '2. OG'],
    // SN0019, SN0020 fehlen absichtlich
];

const ist3 = [ // Abt. Produktion — Spaltenname "Seriennummer"
    ['SN0021', 'Halle A', 'Produktion 1'],
    ['SN0022', 'Halle A', 'Produktion 1'],
    ['SN0023', 'Halle A', 'Produktion 2'],
    ['SN0024', 'Halle B', 'Produktion 2'],
    ['SN0025', 'Halle B', 'Produktion 3'],
    ['SN0026', 'Halle B', 'Produktion 3'],
    ['SN0027', 'Halle C', 'Produktion 4'],
    ['SN0028', 'Halle C', 'Produktion 4'],
    // SN0029 fehlt absichtlich
    ['SN0030', 'Halle C', 'Produktion 4'],
];

async function writeSoll() {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Test-Generator';
    const ws = wb.addWorksheet('Soll-Server');
    const header = ws.addRow(['Seriennummer', 'Modell', 'Standort (Soll)', 'Inbetriebnahme']);
    header.font = { bold: true };
    header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };
    for (const s of sollServers) {
        ws.addRow([s.sn, s.model, s.standort, s.inbetriebnahme]);
    }
    [20, 18, 22, 18].forEach((w, i) => ws.getColumn(i + 1).width = w);
    const p = path.join(OUT_DIR, 'Soll-Liste_Server.xlsx');
    await wb.xlsx.writeFile(p);
    console.log('✓ geschrieben:', p);
}

async function writeIst(filename, sheetName, snHeader, extraHeaders, rows) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Test-Generator';
    const ws = wb.addWorksheet(sheetName);
    const header = ws.addRow([snHeader, ...extraHeaders]);
    header.font = { bold: true };
    header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } };
    for (const r of rows) ws.addRow(r);
    ws.columns.forEach(c => { c.width = 20; });
    const p = path.join(OUT_DIR, filename);
    await wb.xlsx.writeFile(p);
    console.log('✓ geschrieben:', p);
}

(async () => {
    try {
        await writeSoll();
        await writeIst('Ist-Liste_Abt-IT.xlsx',          'IT-Inventar',  'Serial No.',    ['Modell', 'Standort', 'Status'],     ist1);
        await writeIst('Ist-Liste_Abt-Buchhaltung.xlsx', 'Inventur',     'S/N',           ['Bezeichnung', 'Etage'],             ist2);
        await writeIst('Ist-Liste_Abt-Produktion.xlsx',  'Server',       'Seriennummer',  ['Gebäude', 'Bereich'],               ist3);
        
        console.log('\n--- Erwartetes Ergebnis beim Serial-Check ---');
        console.log('Fehlend in allen Ist-Listen: SN0009, SN0010, SN0019, SN0020, SN0029  (5 Einträge)');
        console.log('\nTests für Normalisierung (müssen als GEFUNDEN gelten):');
        console.log('  sn0002  -> SN0002 (lowercase)');
        console.log('  " SN0003 " -> SN0003 (whitespace)');
        console.log('  0000SN0004 -> SN0004 (führende Nullen)');
    } catch (e) {
        console.error('Fehler:', e);
        process.exit(1);
    }
})();
