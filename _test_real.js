const fs = require('fs');
const path = require('path');

async function main() {
    const testFile = '/Users/nojan/Desktop/Enclosure_02-06-2025.xlsx';
    
    if (!fs.existsSync(testFile)) {
        console.log('FEHLER: Datei nicht gefunden');
        process.exit(1);
    }
    
    const fileSize = fs.statSync(testFile).size;
    console.log('Datei:', path.basename(testFile), '- Groesse:', (fileSize / 1024).toFixed(1), 'KB');
    
    // Step 1: AdmZip Sheet-Namen
    console.log('\n--- Step 1: AdmZip ---');
    const AdmZip = require('adm-zip');
    let sheets = [];
    try {
        const buf = fs.readFileSync(testFile);
        const zip = new AdmZip(buf);
        const wbEntry = zip.getEntry('xl/workbook.xml');
        if (!wbEntry) {
            console.log('FAIL: workbook.xml fehlt');
            process.exit(1);
        }
        const xml = wbEntry.getData().toString('utf8');
        const pat = /<sheet[^>]*\bname="([^"]*)"[^>]*>/g;
        let m;
        while ((m = pat.exec(xml)) !== null) sheets.push(m[1]);
        console.log('OK:', sheets.length, 'Sheets:', sheets.slice(0, 5).join(', '), sheets.length > 5 ? '...' : '');
    } catch (e) {
        console.log('FAIL:', e.message);
        process.exit(1);
    }
    
    // Step 2: readSheetWithExcelJS mit erstem Sheet
    if (sheets.length > 0) {
        console.log('\n--- Step 2: readSheetWithExcelJS ("' + sheets[0] + '") ---');
        try {
            const { readSheetWithExcelJS } = require('./exceljs-reader');
            const r = await readSheetWithExcelJS(testFile, sheets[0]);
            if (r.success) {
                console.log('OK:', r.headers.length, 'cols,', r.data.length, 'rows');
                console.log('Headers (erste 5):', r.headers.slice(0, 5));
                if (r.data.length > 1) console.log('Erste Datenzeile:', JSON.stringify(r.data[1]?.slice(0, 5)));
            } else {
                console.log('FAIL:', r.error);
            }
        } catch (e) {
            console.log('EXCEPTION:', e.message);
        }
    }
    
    console.log('\nDONE');
}

main();
