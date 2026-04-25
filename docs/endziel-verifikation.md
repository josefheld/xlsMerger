# Endziel & Verifikationsdokument für xlsMerger

## Zweck des Dokuments

Dieses Dokument beschreibt das **Endziel** des Projekts und dient als **Abnahme- und Verifikationsgrundlage**. 
Wenn alle Muss-Kriterien erfüllt und nachweisbar sind, gilt das Projekt als korrekt umgesetzt.

---

## 1) Endziel (Produktvision)

`xlsMerger` ist am Ende ein robustes, produktionsnahes CLI-Tool, das Excel-Dateien konsolidiert und validiert.
Nutzer wählen einen von drei fachlichen Modi:

1. `finance_close`
2. `supplier_normalizer`
3. `hr_consolidator`

Das Tool liefert:

- eine konsolidierte Ausgabedatei,
- einen maschinenlesbaren Report (`json`/`csv`) mit Fehlern, Warnungen und Kennzahlen,
- verlässliche Exit-Codes für Automation (CI, Cron, Pipelines).

---

## 2) Scope (In Scope / Out of Scope)

### In Scope

- Python-3-kompatibles Tooling und Codebasis
- Einlesen von `.xls` und `.xlsx`
- Profilgesteuerte Validierung und Transformation
- Deterministische Merge-Logik
- CLI-Bedienung (`run`, `validate`, `profiles list`)
- Automatisierte Tests und CI

### Out of Scope (für v1)

- Vollständige Web-App
- Rollen-/Rechtesystem
- Cloud-Synchronisation und Mandantenverwaltung

---

## 3) Zielarchitektur (Soll-Zustand)

```text
.
├── cli/                # CLI-Commands und Argument-Parsing
├── core/               # Reader, Merge-Engine, Validatoren, Reporter
├── profiles/           # finance/supplier/hr Regeln
├── tests/              # unit + integration + golden tests
├── examples/           # Beispielkonfigurationen und Beispieldaten
├── docs/               # Sprintplan, Endziel, Betriebshinweise
├── pyproject.toml      # Build + Tooling (ruff, pytest)
└── ...
```

Architekturprinzipien:

- **Single Engine, Multiple Profiles**: eine zentrale Engine, austauschbare Profile.
- **Determinismus**: gleiche Inputs + Konfiguration = gleicher Output.
- **Beobachtbarkeit**: jeder Lauf erzeugt strukturierte Reports.
- **Automatisierbarkeit**: stabile Exit-Codes für Batch-Betrieb.

---

## 4) Funktionale Anforderungen (MUSS)

## F1 — Dateiverarbeitung

- System verarbeitet `.xls` und `.xlsx`.
- Mehrere Dateien und mehrere Sheets werden unterstützt.
- Leere/defekte Dateien werden sauber als Fehler gemeldet.

**Nachweis**

- Integrationstest mit gemischtem Dateisatz (`.xls` + `.xlsx`) besteht.
- Fehlerfälle erscheinen im Report.

## F2 — Merge-Logik

- Kein hartes Zeilen-/Spaltenlimit.
- Headerstrategie konfigurierbar (`first_file`, `every_file`, `none`).
- Reihenfolge ist deterministisch (Dateiname + Sheetname).

**Nachweis**

- Golden-Test vergleicht erwartete mit tatsächlicher Output-Datei.

## F3 — Profile

- `finance_close`: Pflichtspalten + Typprüfung + Plausibilitätscheck
- `supplier_normalizer`: Mapping + Dublettencheck
- `hr_consolidator`: Schema + PII-Maskierung

**Nachweis**

- Je Profil mindestens 1 Erfolgs- und 1 Fehler-Integrationstest.

## F4 — CLI

- `xlsmerger run --mode <mode> ...`
- `xlsmerger validate --mode <mode> ...`
- `xlsmerger profiles list`

**Nachweis**

- CLI-Help vorhanden und Befehle mit Exit-Code 0/1/2 nutzbar.

## F5 — Reporting

- Reportausgabe in `json` und `csv`
- Pflichtfelder: Laufzeit, Datei-Anzahl, Zeilen in/out, Warnungen, Fehler

**Nachweis**

- Reportschema in Tests validiert.

---

## 5) Nicht-funktionale Anforderungen (MUSS)

## NF1 — Qualität

- Linting (ruff) ohne Fehler.
- Test-Suite (pytest) grün.
- Kritische Kernlogik mit Unit-Tests abgedeckt.

## NF2 — Wartbarkeit

- Klare Modulgrenzen (`core`, `profiles`, `cli`).
- Keine versteckten Side-Effects in Kernfunktionen.

## NF3 — Reproduzierbarkeit

- Build-/Testprozess über `pyproject.toml` ausführbar.
- CI führt lint + tests bei Pull Requests aus.

---

## 6) Verifikationscheckliste (Abnahme)

Statuslegende:

- `[ ]` offen
- `[x]` erfüllt
- `[n/a]` nicht relevant

### A. Setup & Struktur

- [ ] Verzeichnisstruktur entspricht Zielarchitektur.
- [ ] `pyproject.toml` vorhanden und konsistent.
- [ ] README enthält Start- und Prüfhinweise.

### B. Tooling

- [ ] `python -m ruff check .` ist grün.
- [ ] `python -m pytest` ist grün.
- [ ] Test-Discovery findet Unit- und Integrationstests.

### C. Funktionale Tests

- [ ] `.xls` und `.xlsx` werden eingelesen.
- [ ] Merge ohne harte Limits bestätigt.
- [ ] Headerstrategien funktionieren wie dokumentiert.
- [ ] Alle drei Profile sind ausführbar.
- [ ] Fehlerfälle erscheinen im Report.

### D. CLI

- [ ] `xlsmerger profiles list` zeigt alle Modi.
- [ ] `xlsmerger validate` liefert korrekte Validierung.
- [ ] `xlsmerger run` erzeugt Output + Report.
- [ ] Exit-Codes 0/1/2 sind stabil dokumentiert.

### E. Reporting

- [ ] JSON-Report entspricht Pflichtschema.
- [ ] CSV-Report enthält alle Kernmetriken.
- [ ] Reports sind pro Lauf eindeutig referenzierbar.

### F. CI/CD

- [ ] CI läuft bei jedem PR.
- [ ] CI enthält mindestens lint + tests.
- [ ] Release-Artefakte (optional in v1) reproduzierbar.

### G. Dokumentation

- [ ] README Quickstart vollständig.
- [ ] Beispiele für alle 3 Modi vorhanden.
- [ ] Troubleshooting dokumentiert.

---

## 7) Testprotokoll (auszufüllen bei Abnahme)

| Datum | Prüfer | Commit/Tag | Ergebnis | Notizen |
|---|---|---|---|---|
| YYYY-MM-DD | Name | SHA/Tag | PASS/FAIL | |

Zusätzlich auszufüllen:

- Ausgeführte Commands:
  - `python -m ruff check .`
  - `python -m pytest`
  - relevante CLI-Commands
- Anhänge:
  - Beispiel-Input-Dateien
  - erzeugte Reports
  - ggf. Fehlerscreenshots

---

## 8) Exit-/Entscheidungskriterien (Go/No-Go)

## Go, wenn

- alle Muss-Anforderungen erfüllt sind,
- Abnahme-Checkliste in den Muss-Punkten vollständig auf `[x]` steht,
- Pilotnutzer messbaren Mehrwert bestätigen (Zeitersparnis, geringere Fehlerquote).

## No-Go/Archivierung, wenn

- zentrale Muss-Kriterien trotz zwei Iterationen nicht stabil erreicht werden,
- kein belastbarer Nutzermehrwert nachweisbar ist,
- Betriebsaufwand den Nutzen klar übersteigt.

---

## 9) Definition of Done (Projekt v1)

Projekt v1 gilt als „done“, wenn:

1. Alle Punkte in Abschnitt 4 und 5 erfüllt sind.
2. Abschnitt 6 (Abnahme) in allen Muss-Punkten `[x]` ist.
3. Testprotokoll (Abschnitt 7) vollständig ausgefüllt ist.
4. Go-Kriterien (Abschnitt 8) erfüllt sind.

