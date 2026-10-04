# Handballtracker – Fehler und geplante Changes

Stand: 03.10.2026. Diese Datei dient als fortlaufende Fehler- und Änderungsliste.
Alle folgenden Punkte sind offen; die Dokumentation bedeutet keine Umsetzung.

## Übersicht

| ID | Typ | Priorität | Thema | Status |
|---|---|---|---|---|
| HB-001 | Change | Sehr dringend (Nutzer) | Automatischer Uhrstopp bei 30:00 und 60:00 | Offen |
| HB-002 | Fehler | Hoch (Vorschlag) | Rot/Blau sperrt den Spieler nicht dauerhaft | Offen |
| HB-003 | Fehler | Hoch (Vorschlag) | Zwei-Minuten-Sperre über andere Position umgehbar | Offen |
| HB-004 | Change | Normal (Vorschlag) | Assist unabhängig von einer Wurfeingabe erfassen | Offen |

## HB-001 – Automatischer Uhrstopp

- Ist: Die Spieluhr läuft über 30:00 und 60:00 hinaus; der Halbzeitwechsel ist manuell.
- Soll: Automatisch exakt bei 30:00 und 60:00 pausieren.
- Bereits vereinbart: Beim erneuten Start nach der ersten Halbzeit automatisch HZ 2 aktivieren, nicht schon beim Stopp.
- Prüfung: Grenzen auch bei verzögerten Timer-Aufrufen einhalten; Start der zweiten Halbzeit darf nicht sofort erneut bei 30:00 stoppen; kein automatisches Löschen oder Beenden des Spiels. Verhalten nach 60:00 für eine mögliche Fortsetzung separat festlegen.
- Priorität: Sehr dringend, ausdrücklich vom Nutzer hervorgehoben.

## HB-002 – Spieler mit Rot/Blau dauerhaft sperren

- Gemeldet: Spieler können trotz roter oder blauer Karte wieder eingesetzt werden.
- Soll: Betroffene Person für den Rest dieses Spiels auf allen Positionen sperren.
- Alle Einwechselwege berücksichtigen: Drag & Drop, Wechsel-Dialog, Torhüterposition und automatische Rückkehr.
- Ein zulässiger Ersatzspieler darf nach Ablauf der Mannschaftsstrafe eingesetzt werden.
- Prüfung: Sperre bleibt nach Neuladen erhalten. Rücknahme einer falsch erfassten Karte muss den korrekten Zustand wiederherstellen.

## HB-003 – Zwei-Minuten-Sperre an den Spieler binden

- Gemeldet: Die bisherige Position ist gesperrt, der bestrafte Spieler kann aber auf einer anderen Position eingesetzt werden.
- Soll: Der Spieler darf während seiner Strafzeit auf keiner Position eingesetzt werden, unabhängig vom Bedienweg.
- Erlaubt bleiben soll das Vormerken eines zulässigen Rückkehrers für die gesperrte Position per Drag & Drop oder Wechsel-Dialog.
- Prüfung: Kein sofortiger Einsatz über andere Position, TW-Position oder automatische Rückkehr aus anderer Vormerkung. Nach Ablauf wieder zulässig, sofern keine weitere Sperre besteht. Neuladen und Rücknahme der Strafe berücksichtigen.

## HB-004 – Assist ohne erzwungenen Wurf

- Ist: Assist öffnet eine Werferauswahl und erzwingt anschließend eine Wurfeingabe.
- Soll: Assist beim ausgewählten Vorlagengeber als eigenständige Aktion speichern, ohne nachträglichen Wurf zu erzwingen.
- Prüfung: Assist zählt genau einmal; kein zusätzlicher Wurf und kein Tor entstehen. Bestehende mit Würfen verknüpfte Assists bleiben auswertbar. Korrektur, Rückgängig, Cloud und Exporte berücksichtigen.

## Festlegungen für Anleitung und Bedienung

- Drag & Drop ist der primäre Wechselweg, auch zum Vormerken von Spielern bei Zeitstrafen.
- Das Wechsel-Fenster bleibt als Alternative bei Bedienproblemen erhalten.
- Kapitel 4.3 erklärt: Neuladen löscht ein gespeichertes Spiel nicht; über Teams > Mannschaft > Spiele > Weiter fortsetzen.
- Nach Neuladen ist die Uhr angehalten. Den gespeicherten Zeitstand prüfen: Die Uhr wird im aktuellen Code nicht sekündlich dauerhaft gesichert.
- Noch nicht umgesetzte Changes in der Anleitung nicht als bestehende Funktion beschreiben.

## Prüfung vor einem Push auf main

Bestehende Funktionen auf Regressionen prüfen, insbesondere Cloud-/Lokal-Zusammenführung, Erhalt von Teams und Spielen nach Neuladen, Wurferfassung, Wechsel, Strafen und Statistiken. Diese Vorgabe stammt aus der bisherigen Projektabstimmung.

## Vorlage für neue Einträge

### HB-XXX – Kurztitel

- Gemeldet am:
- Typ: Fehler / Change
- Priorität:
- Status: Offen / In Arbeit / Bereit zum Test / Erledigt
- Ist-Verhalten und Schritte zum Nachstellen:
- Erwartetes Verhalten:
- Betroffene Bereiche:
- Prüfkriterien:
- Umsetzung / Commit / PR:
- Testergebnis und Abschlussdatum:
