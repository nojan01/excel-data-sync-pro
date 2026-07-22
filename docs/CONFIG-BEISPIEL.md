# Benutzerprofile in einer gemeinsamen `config.json`

## Überblick

Eine zentrale `config.json` kann von mehreren Mitarbeitern auf einem Netzlaufwerk oder Terminalserver verwendet werden. Persönliche Einstellungen werden dabei nach dem angemeldeten Betriebssystem-Benutzer getrennt – **nicht** nach dem Hostnamen. Das ist wichtig, wenn mehrere Sitzungen dieselbe Servermaschine verwenden.

Beim Speichern schützt die App die gesamte Lese-Merge-Schreiboperation mit einer exklusiven Lock-Datei (`config.json.lock`). Die neue Version wird zunächst in eine temporäre Datei im gleichen Ordner geschrieben und anschließend atomar veröffentlicht. Dadurch können zwei parallele Speichervorgänge keine Benutzerabschnitte mehr verlieren.

Alle gleichzeitig verwendeten Installationen sollten auf diese Version aktualisiert werden. Ältere Programmversionen kennen die Lock-Datei nicht und können deshalb weiterhin parallel direkt in dieselbe Datei schreiben.

## Struktur

```json
{
  "schemaVersion": 2,
  "default": {
    "file1Path": "\\\\server\\share\\Quelldatei.xlsx",
    "file2Path": "\\\\server\\share\\Zieldatei.xlsx",
    "sheet1Name": "Daten",
    "sheet2Name": "Daten",
    "startColumn": 3
  },
  "users": {
    "CONTOSO_ALICE": {
      "file1Path": "Z:\\Projekt\\Quelldatei.xlsx"
    },
    "CONTOSO_BOB": {
      "file1Path": "X:\\Daten\\Quelldatei.xlsx",
      "file2Path": "X:\\Daten\\Zieldatei.xlsx"
    }
  }
}
```

Die Kennung wird aus Benutzername und – sofern verfügbar – Windows-Domäne gebildet, in Großbuchstaben. Sie wird nach einem Speichern in der Statusmeldung angezeigt.

## Laden und Speichern

1. Beim Laden übernimmt die App zuerst alle Werte aus `default`.
2. Danach überschreibt sie diese mit dem Abschnitt in `users` für den angemeldeten Benutzer.
3. Beim Speichern bleibt `default`, alle anderen Benutzerabschnitte und auch unbekannte Felder erhalten.
4. Ist die Datei noch flach, wird sie beim ersten persönlichen Speichern in das obige Format überführt. Die bisherige flache Config wird dabei zu `default`.

Eine neue, erstmals gespeicherte Config bleibt zunächst flach. So ist sie weiterhin direkt als gemeinsamer Team-Standard nutzbar. Erst wenn ein Benutzer persönliche Einstellungen zu einer vorhandenen flachen Datei speichert, erfolgt die Umstellung auf Benutzerprofile.

## Abwärtskompatibilität

Alte Configs mit Rechnerabschnitten bleiben lesbar. Ihr bisheriger Rechnerabschnitt dient übergangsweise als Fallback, bis der Benutzer seine Config einmal speichert. Danach wird ein Benutzerprofil unter `users` angelegt; die alten Abschnitte werden nicht gelöscht.

```json
{
  "default": { "sheet1Name": "Daten" },
  "PC-ALT": { "file1Path": "Z:\\Datei.xlsx" }
}
```

## Hinweise für Administratoren

- Für gemeinsame Daten möglichst UNC-Pfade (`\\server\freigabe\...`) im `default`-Abschnitt verwenden.
- Persönliche Laufwerksbuchstaben und abweichende Dateipfade gehören in den jeweiligen Abschnitt unter `users`.
- Die Datei `config.json.lock` wird nur während eines Speichervorgangs angelegt. Bleibt sie nach einem abgestürzten Prozess länger als fünf Minuten bestehen, wird sie beim nächsten Speichern automatisch ersetzt.
- Wird die Config gerade gespeichert, wartet die App bis zu zwölf Sekunden und zeigt danach eine verständliche Fehlermeldung statt Daten zu überschreiben.
