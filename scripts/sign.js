/**
 * Windows Code Signing Script für electron-builder
 * 
 * SETUP-ANLEITUNG:
 * 
 * 1. Code Signing Certificate besorgen:
 *    - EV Certificate (empfohlen): DigiCert, Sectigo, GlobalSign (~400-600€/Jahr)
 *    - Standard Certificate: Sectigo, Comodo (~100-300€/Jahr)
 * 
 * 2. Environment Variables setzen:
 *    
 *    Option A - PFX-Datei (Standard Certificate):
 *    - CSC_LINK: Pfad zur .pfx Datei (oder Base64-encoded)
 *    - CSC_KEY_PASSWORD: Passwort für die .pfx Datei
 *    
 *    Option B - Hardware Token (EV Certificate):
 *    - SIGNTOOL_PATH: Pfad zu signtool.exe (optional)
 *    - CSC_NAME: Name des Zertifikats im Windows Store
 *    
 * 3. Build ausführen:
 *    npm run build:installer
 * 
 * BEISPIEL .env Datei (nicht committen!):
 * CSC_LINK=./certificates/my-cert.pfx
 * CSC_KEY_PASSWORD=mein-geheimes-passwort
 */

const { execSync } = require('child_process');
const path = require('path');

exports.default = async function sign(configuration) {
  // Wenn keine Signierung konfiguriert, überspringen
  if (!process.env.CSC_LINK && !process.env.CSC_NAME) {
    console.log('⚠️  Code Signing übersprungen - keine Zertifikat-Konfiguration gefunden');
    console.log('   Setze CSC_LINK und CSC_KEY_PASSWORD für PFX-Signierung');
    console.log('   Setze CSC_NAME für Hardware Token Signierung');
    return;
  }

  const filePath = configuration.path;
  const fileName = path.basename(filePath);
  
  console.log(`🔐 Signiere: ${fileName}`);

  try {
    // Option A: PFX-Datei Signierung
    if (process.env.CSC_LINK) {
      const pfxPath = process.env.CSC_LINK;
      const password = process.env.CSC_KEY_PASSWORD || '';
      
      // signtool.exe Pfad ermitteln
      const signtoolPath = process.env.SIGNTOOL_PATH || 
        'C:\\Program Files (x86)\\Windows Kits\\10\\bin\\10.0.22621.0\\x64\\signtool.exe';
      
      const command = `"${signtoolPath}" sign /f "${pfxPath}" /p "${password}" /tr http://timestamp.digicert.com /td sha256 /fd sha256 "${filePath}"`;
      
      execSync(command, { stdio: 'inherit' });
    }
    // Option B: Hardware Token (EV Certificate)
    else if (process.env.CSC_NAME) {
      const certName = process.env.CSC_NAME;
      
      const signtoolPath = process.env.SIGNTOOL_PATH || 
        'C:\\Program Files (x86)\\Windows Kits\\10\\bin\\10.0.22621.0\\x64\\signtool.exe';
      
      const command = `"${signtoolPath}" sign /n "${certName}" /tr http://timestamp.digicert.com /td sha256 /fd sha256 "${filePath}"`;
      
      execSync(command, { stdio: 'inherit' });
    }
    
    console.log(`✅ Erfolgreich signiert: ${fileName}`);
  } catch (error) {
    console.error(`❌ Signierung fehlgeschlagen für ${fileName}:`, error.message);
    // Fehler nicht werfen, damit der Build weiterläuft (unsigniert)
    // throw error; // Aktivieren um Build bei Signierungsfehler abzubrechen
  }
};
