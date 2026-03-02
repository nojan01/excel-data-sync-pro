/**
 * Python Environment Utilities
 * 
 * Gemeinsames Modul für Python-Pfad-Ermittlung.
 * Wird von python_bridge.js und excel_live_bridge.js verwendet.
 */

const path = require('path');
const fs = require('fs');

// Cache für ermittelte Pfade (werden nur einmal ermittelt)
let _cachedPythonPath = null;
let _cachedBasePath = null;
let _pythonEnvPath = null;

/**
 * Sichere Log-Funktion (verhindert EIO-Fehler wenn keine Konsole vorhanden)
 */
function safeLog(...args) {
    try {
        if (process.stdout && process.stdout.writable) {
            console.log(...args);
        }
    } catch (e) {
        // Ignoriere Konsolenfehler
    }
}

/**
 * Liefert den Basis-Pfad für Python-Skripte (entpackt im Produktionsmodus)
 * Ergebnis wird gecacht.
 */
function getPythonBasePath() {
    if (_cachedBasePath) return _cachedBasePath;

    const isPackaged = process.mainModule 
        ? process.mainModule.filename.includes('app.asar')
        : (require.main && require.main.filename.includes('app.asar'));
    
    const hasAsar = process.resourcesPath && fs.existsSync(path.join(process.resourcesPath, 'app.asar'));
    
    if (isPackaged || hasAsar) {
        _cachedBasePath = path.join(process.resourcesPath, 'app.asar.unpacked', 'python');
        safeLog(`[Python] Gepackter Modus - Python-Pfad: ${_cachedBasePath}`);
    } else {
        _cachedBasePath = path.join(__dirname);
        safeLog(`[Python] Dev-Modus - Python-Pfad: ${_cachedBasePath}`);
    }
    
    return _cachedBasePath;
}

/**
 * Python-Pfad ermitteln (embedded oder system).
 * Ergebnis wird gecacht — fs.existsSync wird nur beim ersten Aufruf ausgeführt.
 */
function getPythonPath() {
    if (_cachedPythonPath) return _cachedPythonPath;

    const basePath = getPythonBasePath();
    const isPackaged = basePath.includes('app.asar.unpacked');
    
    if (isPackaged) {
        const resourcesPath = process.resourcesPath;
        
        const embeddedPython = path.join(resourcesPath, 'app.asar.unpacked', 'python-embed', 'win-x64', 'python.exe');
        if (fs.existsSync(embeddedPython)) {
            safeLog(`[Python] Eingebettetes Python gefunden: ${embeddedPython}`);
            _cachedPythonPath = embeddedPython;
            return _cachedPythonPath;
        }
        
        safeLog('[Python] WARNUNG: Eingebettetes Python nicht gefunden, versuche System-Python');
    }
    
    // Dev-Modus: Prüfe ob venv existiert UND das Python-Executable vorhanden ist
    if (!isPackaged) {
        const venvPath = path.join(basePath, '..', '.venv');
        if (fs.existsSync(venvPath)) {
            const venvPython = path.join(venvPath, 'Scripts', 'python.exe');
            if (fs.existsSync(venvPython)) {
                safeLog(`[Python] venv-Python gefunden: ${venvPython}`);
                _cachedPythonPath = venvPython;
                return _cachedPythonPath;
            } else {
                safeLog(`[Python] venv existiert aber Python nicht gefunden: ${venvPython} (überspringe venv)`);
            }
        }
    }
    
    // Fallback: System-Python suchen (Windows)
    const winPythonPaths = [
        'C:\\Python312\\python.exe',
        'C:\\Python311\\python.exe',
        'C:\\Python310\\python.exe',
        'C:\\Python39\\python.exe',
        (process.env.LOCALAPPDATA || '') + '\\Programs\\Python\\Python312\\python.exe',
        (process.env.LOCALAPPDATA || '') + '\\Programs\\Python\\Python311\\python.exe',
        (process.env.LOCALAPPDATA || '') + '\\Programs\\Python\\Python310\\python.exe'
    ];
    
    for (const pyPath of winPythonPaths) {
        if (fs.existsSync(pyPath)) {
            safeLog(`[Python] System-Python gefunden: ${pyPath}`);
            _cachedPythonPath = pyPath;
            return _cachedPythonPath;
        }
    }
    
    // Letzter Fallback
    const pythonCmd = 'python';
    safeLog(`[Python] Fallback auf PATH-Python: ${pythonCmd}`);
    _cachedPythonPath = pythonCmd;
    return _cachedPythonPath;
}

/**
 * Gibt die Umgebungsvariablen für Python-Prozesse zurück
 */
function getPythonEnv() {
    const env = { ...process.env };
    if (_pythonEnvPath) {
        env.PYTHONPATH = _pythonEnvPath + (env.PYTHONPATH ? path.delimiter + env.PYTHONPATH : '');
    }
    return env;
}

/**
 * Setzt den Cache zurück (für Tests)
 */
function resetCache() {
    _cachedPythonPath = null;
    _cachedBasePath = null;
    _pythonEnvPath = null;
}

module.exports = {
    getPythonPath,
    getPythonBasePath,
    getPythonEnv,
    resetCache,
    safeLog
};
