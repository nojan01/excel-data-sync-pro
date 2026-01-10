# Computer-spezifische Konfiguration

## Überblick

Ab Version 1.0.13 unterstützt die `config.json` Computer-spezifische Abschnitte. So können mehrere Benutzer dieselbe zentrale Konfigurationsdatei verwenden, aber unterschiedliche Netzwerkpfade haben.

## Struktur

```json
{
  "default": {
    "file1Path": "\\\\server\\share\\Quelldatei.xlsx",
    "file2Path": "\\\\server\\share\\Zieldatei.xlsx",
    "sheet1Name": "Daten",
    "sheet2Name": "Daten",
    "startColumn": 3,
    "checkColumn": 1,
    "flagColumn": 1,
    "commentColumn": 2
  },
  "PC-MUELLER": {
    "file1Path": "Z:\\Projekt\\Quelldatei.xlsx",
    "file2Path": "Z:\\Projekt\\Zieldatei.xlsx"
  },
  "PC-SCHMIDT": {
    "file1Path": "X:\\Daten\\Quelldatei.xlsx",
    "file2Path": "X:\\Daten\\Zieldatei.xlsx"
  },
  "LAPTOP-MEIER": {
    "file1Path": "M:\\Shared\\Quelldatei.xlsx"
  }
}
```

## Funktionsweise beim Laden

1. **Computername ermitteln**: Die App ermittelt automatisch den Windows-Computernamen (in Großbuchstaben)
2. **Merge-Logik**: 
   - Zuerst werden alle Werte aus `default` geladen
   - Dann werden die Werte des passenden Computer-Abschnitts darüber gelegt
   - Fehlende Werte bleiben aus `default`

## Funktionsweise beim Speichern (Netzwerk-sicher!)

Wenn User 2 seine Config speichert und eine config.json mit verschachtelter Struktur existiert:

1. Die **bestehende config.json wird gelesen**
2. **Nur der eigene Computer-Abschnitt wird aktualisiert**
3. Alle anderen Abschnitte (default, andere Computer) **bleiben unverändert**

### Beispiel

**Vorher** (config.json auf dem Netzlaufwerk):
```json
{
  "default": { "sheet1Name": "Daten" },
  "PC-MUELLER": { "file1Path": "Z:\\Datei.xlsx" }
}
```

**User auf PC-SCHMIDT speichert seine Config**

**Nachher**:
```json
{
  "default": { "sheet1Name": "Daten" },
  "PC-MUELLER": { "file1Path": "Z:\\Datei.xlsx" },
  "PC-SCHMIDT": { "file1Path": "X:\\Datei.xlsx", "file2Path": "X:\\Ziel.xlsx" }
}
```

→ Müllers Einstellungen wurden **nicht überschrieben**!

## Erste Einrichtung (Admin)

1. Config-Datei mit `default`-Abschnitt auf dem Netzlaufwerk anlegen:
   ```json
   {
     "default": {
       "sheet1Name": "Daten",
       "sheet2Name": "Daten",
       "startColumn": 3
     }
   }
   ```

2. Jeder User startet die App und speichert seine Config → sein Abschnitt wird automatisch hinzugefügt

## Beispiel Laden

Wenn PC `PC-SCHMIDT` die Config lädt:
- `file1Path`: `X:\Daten\Quelldatei.xlsx` (vom PC-SCHMIDT-Abschnitt)
- `file2Path`: `X:\Daten\Zieldatei.xlsx` (vom PC-SCHMIDT-Abschnitt)
- `sheet1Name`: `Daten` (vom default-Abschnitt)
- `startColumn`: `3` (vom default-Abschnitt)
- usw.

## Computername ermitteln

Den Computernamen finden Sie:
- **Windows**: Systemsteuerung → System oder `hostname` in CMD
- **In der App**: Wird in der Statuszeile angezeigt beim Laden/Speichern der Config

## Abwärtskompatibilität

Das alte flache Format funktioniert weiterhin:

```json
{
  "file1Path": "C:\\Dateien\\Quelldatei.xlsx",
  "file2Path": "C:\\Dateien\\Zieldatei.xlsx"
}
```

Die App erkennt automatisch ob es sich um das neue oder alte Format handelt.
Bei flachem Format wird die Datei komplett überschrieben (altes Verhalten).

## Tipps

1. **UNC-Pfade im default**: Verwenden Sie im `default`-Abschnitt UNC-Pfade (`\\server\share\...`), dann brauchen nur PCs mit abweichenden Laufwerksbuchstaben einen eigenen Abschnitt.

2. **Nur abweichende Werte**: Tragen Sie im Computer-Abschnitt nur die Werte ein, die vom Standard abweichen.

3. **Groß-/Kleinschreibung**: Computernamen werden automatisch in Großbuchstaben umgewandelt. Schreiben Sie die Abschnittsnamen am besten auch in Großbuchstaben.
