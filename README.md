# Projekt-Reporting – ICT-Projektübersicht (Streamlit-App)

## 1. Übersicht

Die App ist ein Streamlit-Dashboard ("ICT-Projektübersicht"), das Projektdaten aus einer Excel-Datei als Kanban-Board darstellt (nach Projektphase gruppiert), Detailansichten pro Projekt bietet und PDF-Exporte (Gesamtübersicht sowie Einzelprojekt) ermöglicht.

## 2. Dateien im Repository

| Datei | Zweck |
|---|---|
| **`Projekt-Reporting.py`** | Hauptcode der App. Hier werden Änderungen vorgenommen, wenn die Streamlit-Seite bearbeitet werden soll. Streamlit nutzt diese Datei als Einstiegspunkt. |
| **`Monatliches Projekt-Reporting.xlsx`** | Datenquelle. Alle inhaltlichen Änderungen müssen hier vorgenommen werden. Sobald die Datei aktualisiert und ins GitHub-Repository gepusht wurde, übernimmt Streamlit die Inhaltsänderungen automatisch (die App lädt die Datei direkt über eine `raw.githubusercontent.com`-URL). |
| **`push.bat`** | Hilfsskript, um lokale Änderungen (Code und/oder Inhalt) per `git add` / `git commit` / `git push` ins GitHub-Repo zu pushen. Dazu muss der lokale Ordner als lokales Git-Repository eingerichtet sein. Fragt beim Ausführen interaktiv nach einer Commit-Message. |
| **`EMPA.jpeg`** | Empa-Logo, das auf der Streamlit-Seite angezeigt wird. |
| **`requirements.txt`** | Python-Abhängigkeiten, die Streamlit vor dem Start der App installieren muss: `streamlit`, `pandas`, `openpyxl`, `reportlab`. |
| **`README.md`** | Diese Dokumentation |

**Wichtig:** Der Stand auf Joels GitHub (`iotoel`) ist nicht aktuell. Jovo hat lokal noch diverse Änderungen am Code vorgenommen, die noch nicht in dieses Repository gepusht wurden. Vor grösseren Anpassungen unbedingt mit Jovo abklären, welche Version aktuell gilt, bzw. seinen Stand zuerst einspielen.

## 3. Ablage-Ort

- Sämtliche Dateien liegen auf dem privaten GitHub-Repository von Jovo (`jovosp`).
- Die Dateien liegen nicht auf dem Empa-GitLab, da Streamlit nur von GitHub aus deployen kann.
- Die Streamlit-Seite wird derzeit über Jovos privaten Streamlit-Account betrieben. Die Seite selbst ist öffentlich zugänglich.
- Die App muss ggf. zuerst "aufgeweckt" werden (Streamlit Community Cloud legt inaktive Apps schlafen).

## 4. Technischer Aufbau der App (`Projekt-Reporting.py`)

### 4.1 Datenquelle
Die App lädt die Excel-Datei standardmässig über folgende URL:
```
https://raw.githubusercontent.com/iotoel/test/main/Monatliches%20Projekt-Reporting.xlsx
```
Ein lokaler Pfad (`Monatliches Projekt-Reporting.xlsx` im Skript-Verzeichnis) ist als Fallback im Code vorgesehen.

### 4.2 Struktur der Excel-Datei
Die Excel-Datei enthält u. a. folgende Spalten (Formular-Export mit Emoji-Präfixen):

| # | Spalte (Excel) | Interner Name im Code |
|---|---|---|
| 1 | Id | `id` |
| 2 | Startzeit | `start` |
| 3 | Fertigstellungszeit | `end` |
| 4 | E-Mail | `email` |
| 5 | Name | `name` |
| 6 | 📂 Projektname | `project` |
| 7 | 👨‍💼 Projektleiter | `pm` |
| 8 | 📊 Aktuelle Phase | `phase` |
| 9 | 📊 Voraussichtliches Enddatum der aktuellen Phase | `phase_end_date` |
| 10 | ⏳ Änderung des Enddatums / Begründung | `date_changed_reason` |
| 11 | ✅ Aktueller Status (Bewertung durch Projektleiter) | `status` |
| 12 | 💰 IST-Budget extern | `budget_ext` |
| 13 | 🏠 IST-Budget intern | `budget_int` |
| 14 | 🏆 Erfolge gegenüber Vormonat | `successes` |
| 15 | ⚠️ Herausforderungen | `challenges` |
| 16 | 🚨 Eskalation an Transformation Agent | `escalation` |
| 17 | 📅 Ausblick nächster Monat / Termine | `outlook` |
| 18 | 📝 Sonstiges | `misc` |

Pro Projekt wird beim Laden jeweils nur der neueste Eintrag (nach `start`-Datum) berücksichtigt (`groupby("project").last()`), sodass mehrfache Meldungen desselben Projekts über die Zeit korrekt zusammengeführt werden.

### 4.3 Projektphasen (Kanban-Spalten)
1. Entry (Erstellung des Projektantrages)
2. Review Projektantrag (Review durch Transformation Agent)
3. Projekt-Start (Initialisierung)
4. Fertigstellung Konzept (Spezifikation)
5. Fertigstellung Realisierung (Realisierung)
6. Projekt Abschluss (Transition)

### 4.4 Statuslogik
Der Status eines Projekts wird aus dem Freitext-/Auswahlfeld klassifiziert:
- gut 🟢 – enthält "gut", "grün" oder "green"
- risiko 🟡 – enthält "risiko", "gelb" oder "yellow"
- kritisch 🔴 – enthält "kritisch", "rot" oder "red"
- unbekannt ❓ – sonst

Das Enddatum der aktuellen Phase wird zusätzlich farblich markiert:
- ok (grün): Enddatum ≥ 14 Tage in der Zukunft
- warn (gelb): Enddatum 0–13 Tage in der Zukunft
- late (rot): Enddatum bereits überschritten

### 4.5 Funktionen der App
- Kanban-Board: Projekte als Karten, gruppiert nach Phase, mit Projekt-ID (`PA-xxx`), Projektname, Projektleiter, Phase-Ende, Status-Badge, Budget (extern/intern), Eskalationshinweis, Erfolgen, Herausforderungen und Ausblick.
- Filter (Sidebar): nach Projektleiter, nach Status, sowie Volltextsuche im Projektnamen.
- Detailansicht: Klick auf "🔍 Details" öffnet eine ausführliche Ansicht eines einzelnen Projekts mit allen Feldern.
- PDF-Export:
  - Gesamtübersicht (gefiltert) als PDF über den Sidebar-Button "📄 Übersicht als PDF".
  - Einzelprojekt als PDF aus der Detailansicht heraus ("📄 Als PDF exportieren").
  - PDF-Erstellung erfolgt mit `reportlab` (Tabellen/Absätze, A4-Format).

## 5. Funktionsweise von Streamlit (Deployment & Betrieb)

Die App ist unter folgender URL erreichbar: https://itprojektreportingempa.streamlit.app/

Streamlit verwandelt das Python-Skript `Projekt-Reporting.py` in eine Webanwendung, ohne dass HTML/CSS/JS geschrieben werden muss. Der Code läuft server-seitig; jede Interaktion (Klick, Filter, Button) löst einen Rerun des Skripts aus, dessen Ausgabe dann im Browser dargestellt wird.

Zusammenspiel App ↔ GitHub-Repo:

1. Hosting: Die App läuft auf Streamlit Community Cloud (kostenloses Hosting-Angebot von Streamlit, daher die `*.streamlit.app`-Domain) unter Jovos privatem Streamlit-Account.
2. Kopplung an GitHub: Streamlit Community Cloud ist direkt mit dem Repository `iotoel/test` verknüpft und deployed dessen Inhalt, nicht lokale Dateien auf einem Rechner.
3. Einstiegspunkt: In der Konfiguration der Streamlit-Cloud-App ist `Projekt-Reporting.py` als auszuführendes Hauptskript hinterlegt.
4. Abhängigkeiten: Beim Start (bzw. Neustart) installiert Streamlit die in `requirements.txt` gelisteten Pakete (`streamlit`, `pandas`, `openpyxl`, `reportlab`) in einer eigenen Umgebung.
5. Daten-Reload ohne Redeploy: Die Excel-Datei wird nicht aus dem lokalen Dateisystem, sondern zur Laufzeit per HTTP direkt von GitHub geladen (`raw.githubusercontent.com/iotoel/test/main/...xlsx`). Wird die Excel-Datei im Repo aktualisiert und gepusht, zieht sich die laufende App bei jedem Reload automatisch die neue Version, ohne dass ein erneutes Deployment nötig ist.
6. Automatischer Redeploy bei Code-Änderungen: Wird `Projekt-Reporting.py` selbst geändert und auf `main` gepusht, erkennt Streamlit Community Cloud dies über einen GitHub-Webhook und startet die App automatisch neu.
7. Schlafmodus / "Aufwecken": Apps ohne Traffic werden nach einiger Zeit von Streamlit Community Cloud in einen Schlafmodus versetzt, um Ressourcen zu sparen. Beim nächsten Aufruf der URL zeigt Streamlit kurz einen "Wake up"-Hinweis, richtet die Umgebung neu ein und startet die App. Das kann ein bis zwei Minuten dauern.
8. `push.bat` als Auslöser: Da der gesamte Betrieb über GitHub läuft, ist `push.bat` faktisch der Deploy-Mechanismus dieses Projekts. Jeder erfolgreiche `git push` auf `main` löst bei Code-Änderungen einen Redeploy aus bzw. wird bei reinen Excel-Änderungen beim nächsten App-Aufruf automatisch übernommen.

## 6. Änderungen vornehmen

| Ziel | Vorgehen |
|---|---|
| Inhalte ändern (Projektdaten, Status, Texte) | `Monatliches Projekt-Reporting.xlsx` bearbeiten und mit `push.bat` (oder manuell via Git) ins Repository pushen. Streamlit lädt die aktualisierte Datei automatisch von GitHub. |
| Layout/Logik/Funktionen ändern | `Projekt-Reporting.py` bearbeiten und pushen. |
| Abhängigkeiten anpassen | `requirements.txt` pflegen (aktuell: `streamlit`, `pandas`, `openpyxl`, `reportlab`). |
| Änderungen veröffentlichen | `push.bat` ausführen (lokales Git-Repo vorausgesetzt): fragt nach Commit-Message, führt `git add`, `git commit`, `git push -u origin main` aus. |

## 7. Offene Punkte / Bekannte Abweichungen

- Der Stand auf Joels GitHub (`iotoel/test`) ist nicht aktuell. Jovo hat noch Änderungen am Code gemacht, die noch nicht gepusht wurden – vor Arbeiten am Code Rücksprache mit Jovo halten.
- Der im Code referenzierte Excel-Pfad zeigt auf das Repo `iotoel/test`; ein auskommentierter älterer Pfad verweist noch auf `iotoel/projektuebersicht`.
