# MVMS Datenexplorer

Interaktiver Browser-Explorer für MVMS-Excel-Dateien mit schneller Visualisierung großer Tabellen und Such-/Filterfunktionen.

## Nutzung

1. `index.html` im Browser öffnen (Doppelklick reicht, kein Webserver nötig).
2. Excel- oder CSV-Datei auswählen.
3. Arbeitsblatt, Suchfeld und Filter nach Bedarf verwenden.
4. Spalten lassen sich per Klick auf die Kopfzeile ausblenden und über den Bereich „Ausgeblendete Spalten“ wieder einblenden (mindestens eine Spalte muss sichtbar bleiben).
5. Ergebnisse optional über die Export-Schaltfläche speichern.

> Hinweis: Die Verarbeitung erfolgt vollständig im Browser – kein Webserver oder Hintergrundprozess notwendig.

## Export

- **Als XLSX speichern:** Exportiert die aktuell gefilterten Daten als Excel-Datei mit Zeitstempel.

## Debugging

Setze `localStorage.setItem('mvms-debug', '1')` im Browser, um zusätzliche Logausgaben zu aktivieren. Mit `localStorage.removeItem('mvms-debug')` wieder deaktivieren.
