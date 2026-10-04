# Anleitung – Testversion

Die sechs Kapitel sind über „Anleitung“ in der oberen App-Leiste erreichbar. Die Anleitung öffnet sich über der aktuellen Ansicht; ein laufendes Spiel und seine Uhr bleiben aktiv. Screenshots lassen sich antippen und vergrößern. Die Cloud-Abbildung verwendet den neutralen Beispielcode 000000.

## Geprüft am 05.10.2026

- Alle sechs Kapitel bei 375 und 1024 Pixeln Breite ohne horizontalen Überlauf.
- Alle eingebundenen Screenshots, Bildvergrößerung und Kapitelverweise.
- Öffnen und Schließen während eines laufenden Spiels: Uhr läuft weiter, Spiel bleibt erhalten.
- Neuladen erhält Mannschaft und gespeicherten Spielstand.
- Cloud-Abgleich mit simuliertem älteren Cloud-Stand: lokale Mannschaft und neueres lokales Spiel bleiben erhalten, zusätzliche Cloud-Spiele werden ergänzt.
- Bestehende fünf Tests zur Wurferfassung erfolgreich.

Die Prüfung verwendet ausschließlich Testdaten und simulierte Cloud-Antworten. Es wurden keine echten Cloud-Spielstände verändert. Bekannte fachliche Fehler und geplante Änderungen stehen in `BUGS.md`.
