const fs = require('fs');
const path = require('path');

// Erstelle ein 256x256 ICO mit MVMS-Logo
function createIcon() {
    const size = 256;
    const pixels = [];
    
    for (let y = 0; y < size; y++) {
        for (let x = 0; x < size; x++) {
            const cx = size / 2;
            const cy = size / 2;
            const cornerRadius = size * 0.12;
            
            // Prüfe ob Pixel im abgerundeten Rechteck liegt
            let inRect = true;
            if (x < cornerRadius && y < cornerRadius) {
                inRect = Math.sqrt((x - cornerRadius) ** 2 + (y - cornerRadius) ** 2) <= cornerRadius;
            } else if (x >= size - cornerRadius && y < cornerRadius) {
                inRect = Math.sqrt((x - (size - cornerRadius)) ** 2 + (y - cornerRadius) ** 2) <= cornerRadius;
            } else if (x < cornerRadius && y >= size - cornerRadius) {
                inRect = Math.sqrt((x - cornerRadius) ** 2 + (y - (size - cornerRadius)) ** 2) <= cornerRadius;
            } else if (x >= size - cornerRadius && y >= size - cornerRadius) {
                inRect = Math.sqrt((x - (size - cornerRadius)) ** 2 + (y - (size - cornerRadius)) ** 2) <= cornerRadius;
            }
            
            if (!inRect) {
                pixels.push(0, 0, 0, 0);
                continue;
            }
            
            // Blauer Hintergrund (#1565C0)
            let r = 0x15, g = 0x65, b = 0xC0, a = 255;
            
            // Normalisierte Koordinaten
            const nx = (x - cx) / (size / 2);
            const ny = (y - cy) / (size / 2);
            
            // Flugzeug-Symbol
            // Rumpf
            if (Math.abs(nx) < 0.06 && ny > -0.65 && ny < 0.55) {
                r = g = b = 255;
            }
            // Hauptflügel
            if (ny > -0.05 && ny < 0.2) {
                const wingWidth = 0.65 - Math.abs(ny) * 0.5;
                if (Math.abs(nx) < wingWidth && Math.abs(nx) > 0.05) {
                    r = g = b = 255;
                }
            }
            // Heckflügel
            if (ny > -0.55 && ny < -0.42 && Math.abs(nx) < 0.3 && Math.abs(nx) > 0.04) {
                r = g = b = 255;
            }
            // Cockpit/Spitze
            if (ny > 0.45 && ny < 0.65) {
                const tipWidth = 0.08 * (0.65 - ny) / 0.2;
                if (Math.abs(nx) < tipWidth) {
                    r = g = b = 255;
                }
            }
            
            // BGRA Format
            pixels.push(b, g, r, a);
        }
    }
    
    // Für 256x256 verwenden wir PNG im ICO (moderner Standard)
    // Erstelle einfaches PNG
    const pngData = createPNG(size, pixels);
    
    // ICO Header
    const iconDir = Buffer.alloc(6);
    iconDir.writeUInt16LE(0, 0);     // Reserved
    iconDir.writeUInt16LE(1, 2);     // Type (1 = ICO)
    iconDir.writeUInt16LE(1, 4);     // Number of images
    
    // ICO Directory Entry für PNG
    const iconEntry = Buffer.alloc(16);
    iconEntry.writeUInt8(0, 0);                    // Width (0 = 256)
    iconEntry.writeUInt8(0, 1);                    // Height (0 = 256)
    iconEntry.writeUInt8(0, 2);                    // Color palette
    iconEntry.writeUInt8(0, 3);                    // Reserved
    iconEntry.writeUInt16LE(1, 4);                 // Color planes
    iconEntry.writeUInt16LE(32, 6);                // Bits per pixel
    iconEntry.writeUInt32LE(pngData.length, 8);   // Size of PNG data
    iconEntry.writeUInt32LE(22, 12);               // Offset to PNG data
    
    const ico = Buffer.concat([iconDir, iconEntry, pngData]);
    
    const outputPath = path.join(__dirname, '..', 'assets', 'icon.ico');
    fs.writeFileSync(outputPath, ico);
    console.log('Icon erstellt:', outputPath, `(${ico.length} bytes)`);
}

function createPNG(size, pixels) {
    const zlib = require('zlib');
    
    // PNG Signature
    const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
    
    // IHDR Chunk
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(size, 0);   // Width
    ihdr.writeUInt32BE(size, 4);   // Height
    ihdr.writeUInt8(8, 8);          // Bit depth
    ihdr.writeUInt8(6, 9);          // Color type (RGBA)
    ihdr.writeUInt8(0, 10);         // Compression
    ihdr.writeUInt8(0, 11);         // Filter
    ihdr.writeUInt8(0, 12);         // Interlace
    
    const ihdrChunk = createChunk('IHDR', ihdr);
    
    // IDAT Chunk (image data)
    // Convert BGRA to RGBA and add filter bytes
    const rawData = [];
    for (let y = 0; y < size; y++) {
        rawData.push(0); // Filter byte (none)
        for (let x = 0; x < size; x++) {
            const idx = (y * size + x) * 4;
            // BGRA -> RGBA
            rawData.push(pixels[idx + 2]); // R
            rawData.push(pixels[idx + 1]); // G
            rawData.push(pixels[idx + 0]); // B
            rawData.push(pixels[idx + 3]); // A
        }
    }
    
    const compressed = zlib.deflateSync(Buffer.from(rawData), { level: 9 });
    const idatChunk = createChunk('IDAT', compressed);
    
    // IEND Chunk
    const iendChunk = createChunk('IEND', Buffer.alloc(0));
    
    return Buffer.concat([signature, ihdrChunk, idatChunk, iendChunk]);
}

function createChunk(type, data) {
    const length = Buffer.alloc(4);
    length.writeUInt32BE(data.length, 0);
    
    const typeBuffer = Buffer.from(type, 'ascii');
    const crcData = Buffer.concat([typeBuffer, data]);
    const crc = crc32(crcData);
    
    const crcBuffer = Buffer.alloc(4);
    crcBuffer.writeUInt32BE(crc, 0);
    
    return Buffer.concat([length, typeBuffer, data, crcBuffer]);
}

function crc32(data) {
    let crc = 0xffffffff;
    const table = [];
    
    for (let i = 0; i < 256; i++) {
        let c = i;
        for (let j = 0; j < 8; j++) {
            c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
        }
        table[i] = c;
    }
    
    for (let i = 0; i < data.length; i++) {
        crc = table[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
    }
    
    return (crc ^ 0xffffffff) >>> 0;
}

createIcon();
