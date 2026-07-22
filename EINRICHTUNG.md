# Buch-Uhr iPad-PWA v17

Diese erste Begleitfassung verwendet dasselbe `Buch-Uhr.project.json` wie das Windows-Programm und greift über Microsoft Graph auf den OneDrive-Projektordner zu.

## Enthalten

- analoge Buch-Uhr mit Rastertiteln und blauen Fortschrittssegmenten
- Touch-Verschiebung von Dateien auf der Uhr
- leicht magnetisches Minutenraster, dazwischen sekundengenaue Positionierung
- Stehsatzanzeige einschließlich Trennstrichen
- TXT- und MD-Bearbeitung
- Word-Dateien in OneDrive beziehungsweise Word öffnen
- Rückgängig/Wiederherstellen innerhalb der aktuellen Sitzung
- OneDrive-Laden und -Speichern mit Konflikterkennung über ETags
- installierbare PWA mit Offline-App-Hülle

## 1. PWA veröffentlichen

Die Dateien müssen über **HTTPS** erreichbar sein. Eine einfache Möglichkeit ist GitHub Pages:

1. Neues GitHub-Repository anlegen, beispielsweise `buch-uhr-ipad`.
2. Den gesamten Inhalt dieses Ordners in das Repository hochladen.
3. In GitHub unter **Settings → Pages** die Veröffentlichung aus dem Hauptzweig aktivieren.
4. Die danach angezeigte HTTPS-Adresse notieren. Beispiel:
   `https://DEIN-NAME.github.io/buch-uhr-ipad/`

Alternativ kann derselbe statische Ordner über einen anderen HTTPS-Webspace bereitgestellt werden.

## 2. Microsoft-App registrieren

1. `https://entra.microsoft.com/` öffnen.
2. **App registrations → New registration**.
3. Name: `Buch-Uhr iPad`.
4. Bei unterstützten Kontotypen:
   - für ein privates Microsoft-/OneDrive-Konto eine Einstellung wählen, die persönliche Microsoft-Konten zulässt;
   - bei ausschließlich geschäftlichem OneDrive kann der eigene Mandant verwendet werden.
5. Unter **Authentication → Add a platform → Single-page application** die vollständige HTTPS-Adresse der PWA als Redirect URI eintragen.
6. Unter **API permissions → Microsoft Graph → Delegated permissions** hinzufügen:
   - `User.Read`
   - `Files.ReadWrite`
7. Auf der Übersichtsseite die **Application (client) ID** kopieren.
8. Es wird **kein Client Secret** benötigt und keines darf in die PWA eingetragen werden.

## 3. OneDrive-Projektordner

Der eingetragene Ordner ist relativ zum OneDrive-Stammordner, zum Beispiel:

```text
Thomas Buch
```

Er muss mindestens enthalten:

```text
Thomas Buch/
├── Buch-Uhr.project.json
└── Dateien/
```

Die Windows-Buch-Uhr sollte weiterhin direkt in diesem synchronisierten OneDrive-Ordner arbeiten.

## 4. Auf dem iPad einrichten

1. PWA-Adresse in Safari öffnen.
2. Client-ID, Mandant (`common` für private/persönliche Konten) und Projektordner eintragen.
3. Mit dem Microsoft-Konto anmelden, das den OneDrive-Ordner enthält.
4. In Safari **Teilen → Zum Home-Bildschirm → Als Web-App öffnen** aktivieren.

## Bedienung

- Datei oder blauen Fortschrittsabschnitt auf der Uhr ziehen: Position ändern.
- Antippen: Textansicht öffnen.
- TXT/MD: bearbeiten und speichern.
- DOCX: über **In OneDrive öffnen** an Word übergeben.
- `Tab` mit externer Tastatur: Seitenleisten ein-/ausblenden.
- `Strg/⌘+Z`, `Strg/⌘+Y`: lokale Sitzungshistorie.
- **Synchronisieren** lädt den aktuellen OneDrive-Zustand neu.

## Wichtige Grenze der ersten Fassung

Die Uhrdateien sind zuverlässig erreichbar, weil ihre Projektkopien im Unterordner `Dateien` liegen. Stehsatzdateien, die außerhalb des Projektordners liegen, werden anhand ihres Dateinamens in OneDrive gesucht. Bei mehreren gleichnamigen Dateien kann deshalb die falsche Datei gefunden werden. Eine spätere Windows-Version kann zusätzlich eindeutige OneDrive-relative Pfade im Projekt speichern.

## Lokaler PC-Test

`serve.py` starten:

```powershell
python serve.py
```

Dann `http://localhost:8080` öffnen und diese Adresse zusätzlich als SPA-Redirect-URI registrieren. Die installierbare iPad-Fassung benötigt trotzdem eine HTTPS-Veröffentlichung.


## Änderung in v2

- Microsoft-Anmeldebibliothek wird korrekt vor der App geladen.
- Zahnrad und Einstellungsdialog funktionieren.
- Offline-Cache auf Version 2 gesetzt.


## Änderung in v3

Die Microsoft-Anmeldung erfolgt nicht mehr in einem Popup. Die Buch-Uhr wird im selben Browserfenster zu Microsoft umgeleitet und danach automatisch wieder geöffnet. Das ist in Firefox und Safari zuverlässiger.


## Änderung in v4

Ein von einem früheren Popup- oder Redirect-Versuch zurückgebliebener
`interaction_in_progress`-Status wird automatisch entfernt. Die gespeicherten
Buch-Uhr-Einstellungen bleiben dabei erhalten.


## Änderung in v5

Projekt- und Textdateien werden direkt über den Microsoft-Graph-Endpunkt
`/content` geladen. Die PWA ist nicht mehr darauf angewiesen, dass OneDrive
eine temporäre `@microsoft.graph.downloadUrl` zurückgibt.


## Änderung in v6

- `N`: neue Textdatei
- `F2`: ausgewählte Datei umbenennen
- `Entf` oder Rücktaste: von der Uhr entfernen
- `+`, `-`, `0`: Zoom und Uhr einpassen
- sichtbare Schaltflächen für Touch
- langes Antippen oder Rechtsklick: Kontextmenü
- Rastertitel per Doppelklick oder Kontextmenü
- Normseiten und Farben im Kontextmenü


## Änderung in v7

- neue Textdatei besitzt ein eigenes Titelfeld
- einfacher Klick in der linken Dateiliste wählt nur aus
- Doppelklick öffnet die Datei in OneDrive
- `F2` wirkt auf die links ausgewählte Projektdatei
- Rechtsklick links öffnet das Kontextmenü
- `Tab` blendet nur die beiden Seitenleisten aus; die Uhr bleibt sichtbar


## Änderung in v8

- links und rechts: einfacher Klick wählt nur aus
- Doppelklick öffnet stets den internen Editor
- OneDrive öffnet nur noch über „In OneDrive öffnen“
- `F2` funktioniert auch im Stehsatz
- Zoom verändert nur die Uhr, nicht die Seitenleisten
- Dokumenttitel stehen weiter außerhalb der Uhr und sind kleiner
- `Tab` wird im Capture-Modus abgefangen und blendet nur die Seitenleisten aus


## Änderung in v9

- Endlosschleife in `applySidebarState()` behoben
- Uhr und Canvas werden wieder zuverlässig gerendert
- gespeicherter Zoom wird nach jedem Rendern angewendet
- doppelte Tab-Erkennung entfernt
- Seitenleistenstatus wird beim Start korrekt gesetzt


## Änderung in v10

- Dateititel werden in einem begrenzten mehrzeiligen Textfeld dargestellt
- lange Titel laufen nicht mehr in Uhr oder Grafik hinein
- Titel stehen weiter außerhalb des Uhrkreises
- `Tab` besitzt nur noch einen einzigen Handler im Capture-Modus
- `Tab` blendet ausschließlich beide Seitenleisten aus
- Canvas und Uhr bleiben sichtbar und werden nach dem Umschalten neu berechnet


## Änderung in v11

- `Strg` + Mausrad zoomt nur die Uhr
- Doppelklick auf das Uhrraster öffnet den Rastertitel-Dialog
- Doppelklick auf eine Datei öffnet den internen Dateieditor
- Datei-Doppelklicks werden nicht mehr an den Rastertitel weitergegeben
- OneDrive öffnet nur noch über den entsprechenden Button im Editor


## Änderung in v12

- Rastertitel stehen in begrenzten Textfeldern weiter innerhalb des Uhrkreises
- Datei-Doppelklick wird sicher vom Raster-Doppelklick getrennt
- auf dem iPad öffnet zweimaliges kurzes Antippen einer Datei den Editor
- Ziehen auf einer freien Canvasfläche verschiebt die gesamte Uhr
- Ziehen an einer Datei verschiebt weiterhin nur diese Datei auf dem Raster
- Taste `0` setzt Zoom und Canvasposition zurück


## Änderung in v13
- einfacher Klick wählt nur aus und verändert nie die Position
- Ziehen beginnt erst nach mehr als 10 Pixel Bewegung
- Doppelklick und Enter öffnen ausgewählte Dateien
- Rastertitel näher am Rand; Linien enden vor dem Text


## Änderung in v14
- Auswahl baut Elemente nicht neu auf
- Doppelklick öffnet Dateien an Uhr, links und rechts
- Rasterdialog ignoriert Dateielemente
- `E` schaltet wie `Tab` die Seitenleisten
- auf dem iPad übernimmt der Menüknopf oben links diese Funktion


## Änderung in v15

- bei einem einfachen Klick auf eine Uhrdatei wird die Dateiebene nicht mehr neu gezeichnet
- das angeklickte DOM-Element bleibt zwischen erstem und zweitem Klick erhalten
- Firefox erkennt den Doppelklick deshalb zuverlässig als Doppelklick auf die Datei
- der Rasterdialog erhält den zweiten Klick nicht mehr
- nur ein tatsächliches Ziehen verändert Position und Darstellung


## Änderung in v16

- Rasterdialog öffnet ausschließlich bei Doppelklick direkt auf einen sichtbaren Rasterstrich
- Doppelklick auf freien Canvas macht nichts
- Dateisymbole und Dateititel erhalten eigene Zeigerereignisse
- 15- und 45-Minuten-Titel bleiben mit fester Innenkante vollständig innerhalb des Uhrkreises
- Verbindungslinien enden vor dem Titelbereich


## Änderung in v17

- zwei Klicks auf dieselbe Uhrdatei innerhalb von 430 ms öffnen den Editor
- native Browser-Doppelklick-Erkennung wird für Uhrdateien nicht mehr verwendet
- dieselbe Öffnungsfunktion wie bei `Enter`
- Ziehen wird bei `pointerup`, `pointercancel`, verlorenem Pointer-Capture und Fensterfokusverlust sicher beendet
- sobald keine Maustaste mehr gedrückt ist, wird ein eventuell verbliebener Drag-Zustand verworfen
- nach einem echten Ziehen wird der nachfolgende Klick kurz unterdrückt
