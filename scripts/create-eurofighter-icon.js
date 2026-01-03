const fs = require('fs');
const path = require('path');

// Eurofighter SVG - stilisiertes Jet-Icon
const eurofighterSVG = `<?xml version="1.0" encoding="UTF-8"?>
<svg width="256" height="256" viewBox="0 0 256 256" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="skyGradient" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" style="stop-color:#1a5fb4"/>
      <stop offset="100%" style="stop-color:#3584e4"/>
    </linearGradient>
    <linearGradient id="jetGradient" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" style="stop-color:#f5f5f5"/>
      <stop offset="50%" style="stop-color:#c0c0c0"/>
      <stop offset="100%" style="stop-color:#888888"/>
    </linearGradient>
  </defs>
  
  <!-- Hintergrund - Himmel -->
  <rect width="256" height="256" rx="40" fill="url(#skyGradient)"/>
  
  <!-- Wolken -->
  <ellipse cx="50" cy="200" rx="30" ry="15" fill="rgba(255,255,255,0.3)"/>
  <ellipse cx="200" cy="220" rx="40" ry="18" fill="rgba(255,255,255,0.25)"/>
  <ellipse cx="180" cy="50" rx="25" ry="12" fill="rgba(255,255,255,0.2)"/>
  
  <!-- Eurofighter - von der Seite -->
  <g transform="translate(128, 128) rotate(-15)">
    <!-- Rumpf -->
    <path d="M-80,0 L-60,-8 L60,-8 L90,0 L60,8 L-60,8 Z" fill="url(#jetGradient)" stroke="#555" stroke-width="1"/>
    
    <!-- Cockpit -->
    <path d="M30,-7 Q50,-7 60,-4 L60,4 Q50,7 30,7 Q35,0 30,-7" fill="#2ec27e" opacity="0.8"/>
    
    <!-- Hauptflügel (Delta) -->
    <path d="M-20,-8 L-50,-45 L-10,-45 L10,-8 Z" fill="url(#jetGradient)" stroke="#555" stroke-width="1"/>
    <path d="M-20,8 L-50,45 L-10,45 L10,8 Z" fill="url(#jetGradient)" stroke="#555" stroke-width="1"/>
    
    <!-- Canard-Flügel (vorne) -->
    <path d="M40,-6 L55,-25 L65,-25 L55,-6 Z" fill="url(#jetGradient)" stroke="#555" stroke-width="1"/>
    <path d="M40,6 L55,25 L65,25 L55,6 Z" fill="url(#jetGradient)" stroke="#555" stroke-width="1"/>
    
    <!-- Seitenleitwerk -->
    <path d="M-70,-8 L-80,-30 L-60,-30 L-55,-8 Z" fill="url(#jetGradient)" stroke="#555" stroke-width="1"/>
    
    <!-- Triebwerksauslass -->
    <ellipse cx="-75" cy="0" rx="8" ry="6" fill="#ff7800"/>
    <ellipse cx="-78" cy="0" rx="5" ry="4" fill="#ffcc00"/>
    
    <!-- Details -->
    <line x1="-40" y1="-6" x2="20" y2="-6" stroke="#777" stroke-width="0.5"/>
    <line x1="-40" y1="6" x2="20" y2="6" stroke="#777" stroke-width="0.5"/>
  </g>
  
  <!-- MVMS Text -->
  <text x="128" y="240" font-family="Arial, sans-serif" font-size="24" font-weight="bold" 
        fill="white" text-anchor="middle" opacity="0.9">MVMS</text>
</svg>`;

// Speichere SVG
const assetsDir = path.join(__dirname, '..', 'assets');
const svgPath = path.join(assetsDir, 'icon.svg');

fs.writeFileSync(svgPath, eurofighterSVG);
console.log('✅ SVG erstellt:', svgPath);

console.log(`
📋 Nächste Schritte für Icon-Konvertierung:

1. Online-Konverter (einfachste Methode):
   - Öffne https://cloudconvert.com/svg-to-ico
   - Lade assets/icon.svg hoch
   - Konvertiere zu .ico (256x256) → speichere als assets/icon.ico
   
   - Für macOS: https://cloudconvert.com/svg-to-icns
   - Konvertiere zu .icns → speichere als assets/icon.icns

2. Oder mit ImageMagick (falls installiert):
   brew install imagemagick
   
   # PNG erstellen
   convert -background none -resize 256x256 assets/icon.svg assets/icon.png
   
   # ICO für Windows
   convert assets/icon.png -define icon:auto-resize=256,128,64,48,32,16 assets/icon.ico
   
   # ICNS für macOS
   mkdir -p assets/icon.iconset
   for size in 16 32 64 128 256 512; do
     convert -background none -resize ${size}x${size} assets/icon.svg assets/icon.iconset/icon_${size}x${size}.png
   done
   iconutil -c icns assets/icon.iconset -o assets/icon.icns

Das SVG wurde erstellt und kann in jedem Browser angesehen werden!
`);
