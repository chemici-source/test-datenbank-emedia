<!--
  PROJEKT-STECKBRIEF
  Dieses Repo wurde aus „bibmed-vorlage“ erstellt. Bitte jetzt:
  1. Alle Platzhalter unten ausfüllen (Projektname, Personen, Daten, Ziel).
  2. BOARD-LINK ersetzen: Board „BibMED Übersicht“ öffnen, Ansicht „Nach Repo/Projekt“,
     in der Filterzeile repo:BENUTZERNAME/<name-dieses-repos> eingeben und die Adresse
     aus der Adresszeile kopieren.
  3. HANDBUCH-LINK ersetzen durch die Adresse von bibmed-handbuch.
  Kommentare wie dieser sind in der Ansicht unsichtbar und dürfen danach gelöscht werden.
-->

# Datenbank E-Medien

| | |
|---|---|
| **Status** | 🟢 läuft |
| **Verantwortlich** | @lorena-staiger |
| **Mitarbeit** | @david-stendardo, Yassica Umaparan (ab Nov26), Mo Wa Baile, Antonella Schintu |
| **Auftraggeber/in** | @lorena-staiger, Michelle Schaffer |
| **Start** | 16.07.2026 |
| **Geplantes Ende** | 31.03.2027 |
| **Aufgaben** | [Issues dieses Repos](../../issues) · [im Board](https://github.unibe.ch/users/lorena-staiger/projects/2/views/8?filterQuery=repo%3Alorena-staiger%2Fdatenbank-emedia) |
| **Meilensteine** | [Milestones](../../milestones) |
| **Arbeitsweise** | [BibMED Handbuch](https://github.unibe.ch/lorena-staiger/bibmed-handbuch) |

## Ziel

Eine interne Datenbank für die Verwaltung und das Nachschlagen aller MNW-E-Medien, inklusive Nutzungsstatistiken. 

## Ergebnis / Lieferobjekte

<!-- Was liegt am Ende konkret vor? -->
- [ ] Konzept der neuen Tabellen und Felder inkl. Normalisierung der Datenbank
- [ ] Migration der Daten zu SQL-Datenbank auf Server
- [ ] Web-Apps: Admin und Leser

## Nicht Teil dieses Projekts

<!-- Was wird bewusst NICHT gemacht? Hilft gegen ausufernde Projekte. -->
- Einbindung von Daten anderer Bibliotheksbereiche

## Aktueller Stand

**Stand 02.10.2026:**

Entwurf Tabellen-Struktur steht. Bedarf erneuter Analyse und Überarbeitung mit Claude, insbesondere in Bezug auf Normalisierung.
Server-Recherche ausstehend. Entscheid zu Admin-Oberfläche aus Open-Source-Anbietern ausstehend.

## Wichtige Links

- Daten und Excel-Dateien: s.u. 
- Teams-Kanal: Apps und Technologien, Ordner [Datenbank E-Medien](https://unibe365.sharepoint.com/:f:/r/teams/unibe-1112.az/Freigegebene%20Dokumente/Apps%20und%20Technologien/Datenbank%20E-Medien?d=w45c7e2536b2b413ea8e8a1b37d07a483&csf=1&web=1&e=9407nl)
- Externe Quellen: 

## Entscheidungen

Siehe [docs/entscheidungen.md](docs/entscheidungen.md).

## Anleitung / Wie es funktioniert

<!--
  Für Projekte, die etwas Dauerhaftes hinterlassen (Ablauf, Skript, Auswertung):
  Wie benutzt man es? Was muss man regelmässig tun? Wo kann es schiefgehen?
  Bei Programmierprojekten: Voraussetzungen, Installation, Ausführen.
  So geschrieben, dass jemand ohne Vorwissen es nachvollziehen kann.
-->
**Ansehen der Webapp, Entwurf Leseansicht**

Ordner "webapp" herunterladen, im cmd zum Pfad navigieren, dann "python app.py". 

## Offene Fragen

- …
