# Sprintplan: xlsMerger Produktisierung (4 Wochen)

## Zielbild

Nach 4 Wochen existiert eine produktionsnahe CLI-App `xlsmerger`, bei der Nutzer einen von drei Modi auswählen:

1. `finance_close`
2. `supplier_normalizer`
3. `hr_consolidator`

Die App liefert konsolidierte Output-Dateien und einen strukturierten Report (JSON/CSV) mit Warnungen, Fehlern und Kennzahlen.

---

## Architekturprinzip

Eine gemeinsame Engine mit profilgesteuerten Regeln:

- `core/` → Einlesen, Validieren, Mergen, Schreiben
- `profiles/` → Modusspezifische Regeln und Konfiguration
- `cli/` → Bedienung über Kommandos
- `reports/` → Ergebnis- und Fehlerberichte
- `tests/` → Unit- und Integrations-Tests

---

## Woche 1 — Fundament modernisieren

### Tag 1: Setup & Struktur

- Ordnerstruktur einführen (`core/`, `profiles/`, `cli/`, `tests/`, `examples/`)
- `pyproject.toml` + Tooling (z. B. ruff, pytest)
- README um neue Struktur ergänzen

**Akzeptanzkriterien**

- Projekt installierbar und startbar
- Linting und Test-Runner sind ausführbar

### Tag 2: Python-3-Migration

- Python-2-Altlasten entfernen (`xrange`, `iterkeys`)
- Unnötige Imports entfernen
- Type Hints für Kernfunktionen ergänzen

**Akzeptanzkriterien**

- Keine Python-2-Syntax mehr vorhanden
- Tests starten ohne Syntax-/Importfehler

### Tag 3: Reader stabilisieren

- `.xls` und `.xlsx` über einheitliche Reader-Schnittstelle unterstützen
- Interne Datenrepräsentation vereinheitlichen
- Fehlerfälle sauber behandeln (defekte/leere/gesperrte Dateien)

**Akzeptanzkriterien**

- Mindestens 6 Reader-Tests (Erfolg + Fehlerfälle)
- Fehlerhafte Dateien erzeugen eindeutige Fehlermeldungen

### Tag 4: Merge-Engine erneuern

- Harte Limits auf Zeilen/Spalten entfernen
- Header-Strategien implementieren (`first_file`, `every_file`, `none`)
- Deterministische Reihenfolge definieren (Dateiname + Blattname)

**Akzeptanzkriterien**

- Große Dateien werden vollständig verarbeitet
- Header-Verhalten ist reproduzierbar konfigurierbar

### Tag 5: Reporting-Basis

- Report-Datenmodell einführen (Dateien, Zeilen, Warnungen, Fehler)
- Export nach JSON und CSV
- Exit-Codes definieren (`0` Erfolg, `1` Validierung, `2` Systemfehler)

**Akzeptanzkriterien**

- Jeder Lauf erzeugt einen maschinenlesbaren Report
- Exit-Codes sind dokumentiert und stabil

---

## Woche 2 — Drei Profile implementieren

### Tag 6–7: Profil-Framework

- Profilschnittstelle einführen (`validate()`, `transform()`, `postprocess()`)
- Modusauswahl über `--mode`
- Profilkonfiguration via YAML

**Akzeptanzkriterien**

- Modi sind ohne Core-Änderung austauschbar
- Konfiguration wird sauber validiert

### Tag 8: Profil `finance_close`

- Pflichtspalten definieren
- Datentypvalidierung (Datum, Zahl)
- Summen-/Saldo-Plausibilitätscheck

**Akzeptanzkriterien**

- Fehlende Pflichtspalten werden präzise gemeldet
- Ergebnisdatei entspricht Zielschema

### Tag 9: Profil `supplier_normalizer`

- Spalten-Mapping mit Synonymen
- Dublettenerkennung (`invoice_id`, `order_id`)
- Lieferanten-spezifische Normalisierung

**Akzeptanzkriterien**

- Verschiedene Lieferantenlayouts landen im gleichen Schema
- Dubletten werden im Report markiert

### Tag 10: Profil `hr_consolidator`

- HR-Standardschema einführen
- PII-Maskierung konfigurierbar machen
- Validierung für IDs und Datumsfelder

**Akzeptanzkriterien**

- Maskierte Felder sind im Output zuverlässig anonymisiert
- Fehlerhafte Datensätze werden nachvollziehbar gemeldet

---

## Woche 3 — CLI & Nutzerführung

### Tag 11: CLI-Kommandos

- `xlsmerger run`
- `xlsmerger validate`
- `xlsmerger profiles list`

**Akzeptanzkriterien**

- Vollständige `--help` Ausgaben
- Moduswechsel ohne Codeänderung möglich

### Tag 12: Dry-Run & Vorschau

- `--dry-run` implementieren
- Vorschau der ersten N Zeilen
- Laufstatistiken anzeigen

**Akzeptanzkriterien**

- Dry-Run schreibt keine Ausgabedatei
- Report wird trotzdem erzeugt

### Tag 13: Fehler-UX

- Präzise Fehlermeldungen je Datei/Sheet
- Troubleshooting-Hinweise im Report
- Logging-Level (`info`, `warn`, `debug`)

**Akzeptanzkriterien**

- Jede Fehlermeldung enthält Ursache + Handlungsempfehlung

### Tag 14: Beispielkonfigurationen

- `examples/finance.yml`
- `examples/supplier.yml`
- `examples/hr.yml`

**Akzeptanzkriterien**

- Alle Beispielkonfigurationen sind direkt ausführbar

### Tag 15: Dokumentation v1

- Quickstart
- Profil-Auswahlhilfe („Wann welcher Modus?“)
- FAQ („Warum wurden Zeilen verworfen?“)

**Akzeptanzkriterien**

- Erste erfolgreiche Nutzung in <15 Minuten möglich

---

## Woche 4 — Qualität, Build, Pilot

### Tag 16–17: Tests

- Unit-Tests für Core
- Integrations-Tests je Profil
- Golden-File-Tests für deterministische Outputs

**Akzeptanzkriterien**

- Ziel >80% Abdeckung im Core
- Kritische Flows je Profil abgesichert

### Tag 18: CI/CD

- GitHub Actions (lint, test, build)
- Build-Artefakte pro Pipeline
- Versionierung (SemVer)

**Akzeptanzkriterien**

- PRs benötigen grüne Checks

### Tag 19: Packaging

- Wheel/sdist bereitstellen
- Optional: PyInstaller für Desktop-Distribution
- Release-Notes-Template

**Akzeptanzkriterien**

- Reproduzierbare Release-Artefakte erzeugbar

### Tag 20: Pilot & Entscheidung

- Mindestens 3 Pilotnutzer (je Profil mind. 1)
- Feedback erfassen, Bugs priorisieren
- Go/No-Go-Entscheidung

**Akzeptanzkriterien**

- Kennzahlen erhoben: Zeitersparnis, Wiederholnutzung, Fehlerrate

---

## Priorisierung

### Must (v1)

- Python 3 + `.xls`/`.xlsx`
- Stabile Merge-Engine
- Drei Profile auswählbar
- Reports, Exit-Codes, Tests, CI

### Should

- Dry-Run, Vorschau, Beispielkonfigurationen
- Optionales Binary-Packaging

### Won’t (v1)

- Vollständige Web-App
- Rollen/Rechtemodell
- Cloud-Synchronisation

---

## Go/No-Go Kriterien (nach 4 Wochen)

### Go

- Mindestens 5 aktive Nutzer
- Mindestens 2 Wiederholungen pro Woche
- Messbare Zeitersparnis pro Lauf

### No-Go / Archivieren

- Keine wiederkehrende Nutzung
- Kein klarer Mehrwert gegenüber manueller Arbeit
- Unverhältnismäßiger Supportaufwand

---

## Nächste 48 Stunden

1. Projektstruktur und Tooling erstellen
2. Python-3-Migration + Reader-Refactor starten
3. End-to-End-Lauf mit `finance_close` herstellen

