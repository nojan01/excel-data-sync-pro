'use strict';

/* eslint-disable no-restricted-globals */

importScripts('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');

let workbookCache = null;

self.postMessage({ type: 'ready' });

self.addEventListener('message', (event) => {
    const { type, payload } = event.data || {};
    switch (type) {
        case 'parse-file':
            parseFile(payload);
            break;
        case 'load-sheet':
            loadSheet(payload?.sheetName);
            break;
        default:
            break;
    }
});

function parseFile(payload) {
    if (!payload || !payload.buffer) {
        emitError('Keine gültige Datei erhalten.');
        return;
    }
    try {
        const buffer = normalizeBuffer(payload.buffer);
        const options = {
            type: 'array',
            cellDates: true,
            cellNF: false,
            cellText: false,
            dateNF: 'yyyy-MM-dd'
        };
        workbookCache = XLSX.read(buffer, options);
        const sheetNames = Array.isArray(workbookCache.SheetNames) ? workbookCache.SheetNames : [];
        self.postMessage({ type: 'workbook-meta', payload: { sheetNames } });
        if (sheetNames.length > 0) {
            const preferredSheet = payload.preferredSheet;
            const target = preferredSheet && sheetNames.includes(preferredSheet) ? preferredSheet : sheetNames[0];
            loadSheet(target);
        } else {
            self.postMessage({ type: 'sheet', payload: { sheetName: '', columns: [], rows: [] } });
        }
    } catch (error) {
        emitError(error?.message || String(error));
    }
}

function loadSheet(sheetName) {
    if (!workbookCache) {
        emitError('Es wurde noch keine Arbeitsmappe geladen.');
        return;
    }
    if (!sheetName) {
        emitError('Es wurde kein Arbeitsblatt angegeben.');
        return;
    }
    const worksheet = workbookCache.Sheets?.[sheetName];
    if (!worksheet) {
        emitError(`Arbeitsblatt "${sheetName}" wurde nicht gefunden.`);
        return;
    }
    try {
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: true });
        const columns = inferColumns(rows);
        self.postMessage({ type: 'sheet', payload: { sheetName, columns, rows } });
    } catch (error) {
        emitError(error?.message || String(error));
    }
}

function inferColumns(rows) {
    const columnSet = new Set();
    rows.forEach((row) => {
        Object.keys(row).forEach((column) => columnSet.add(column));
    });
    return Array.from(columnSet);
}

function emitError(message) {
    self.postMessage({ type: 'error', payload: { message } });
}

function normalizeBuffer(input) {
    if (input instanceof Uint8Array) {
        return input;
    }
    if (input instanceof ArrayBuffer) {
        return new Uint8Array(input);
    }
    if (input && typeof input.byteLength === 'number') {
        try {
            return new Uint8Array(input);
        } catch (error) {
            emitError('Der Dateipuffer konnte nicht verarbeitet werden.');
        }
    }
    throw new Error('Ungültiger Dateipuffer.');
}
