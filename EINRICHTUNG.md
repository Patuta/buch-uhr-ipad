# Buch-Uhr iPad-PWA v31.4

Diese Begleitfassung verwendet dasselbe `Buch-Uhr.project.json` wie das Windows-Programm. Das Programm ist von den Projekten getrennt; die PWA greift über Microsoft Graph auf Projekte unter dem OneDrive-Projektstamm zu.

## Enthalten

- analoge Buch-Uhr mit Rastertiteln und blauen Fortschrittssegmenten
- Touch-Verschiebung von Dateien auf der Uhr
- leicht magnetisches Minutenraster, dazwischen sekundengenaue Positionierung
- Stehsatzanzeige einschließlich Trennstrichen
- projektintern ausschließlich TXT-Arbeitsdateien
- Bearbeiten und Speichern der Projekt-TXT-Dateien
- projektinterne Ordner `Uhr`, `Stehsatz` und `Papierkorb`
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

## 3. OneDrive-Projekte

Der bevorzugte Projektstamm ist:

```text
OneDrive/Buch-Uhr/
```

Darin liegt jedes Projekt getrennt, zum Beispiel:

```text
Buch-Uhr/
    Thomas Buch/
        Buch-Uhr.project.json
        Uhr/
        Stehsatz/
        Papierkorb/
    Bremen-Lexikon/
        Buch-Uhr.project.json
        Uhr/
        Stehsatz/
        Papierkorb/
```

Die PWA zeigt diese Projektordner beim Start an und kann neue Projekte anlegen. Die Windows-Version darf Projekte zusätzlich an jedem beliebigen lokalen oder synchronisierten Speicherort öffnen und speichern. Solche Projekte sind auf dem iPad nur erreichbar, wenn sie im OneDrive-Projektstamm liegen.

## 4. Auf dem iPad einrichten

1. PWA-Adresse in Safari öffnen.
2. Client-ID, Mandant (`common` für private/persönliche Konten) und Projektordner eintragen.
3. Mit dem Microsoft-Konto anmelden, das den OneDrive-Ordner enthält.
4. In Safari **Teilen → Zum Home-Bildschirm → Als Web-App öffnen** aktivieren.

## Bedienung

- Datei oder blauen Fortschrittsabschnitt auf der Uhr ziehen: Position ändern.
- Antippen: Textansicht öffnen.
- Projektdateien sind TXT-Dateien und können bearbeitet und gespeichert werden.
- Dateien auf der Uhr liegen im Ordner `Uhr` und heißen `HH.MM Dateiname.txt`.
- Verlässt eine Datei die Uhr, wird der Zeitpräfix entfernt und die TXT-Datei nach `Stehsatz`, `Papierkorb` oder ins Projekt-Root verschoben.
- `Tab` mit externer Tastatur: Seitenleisten ein-/ausblenden.
- `Strg/⌘+Z`, `Strg/⌘+Y`: lokale Sitzungshistorie.
- **Synchronisieren** lädt den aktuellen OneDrive-Zustand neu.

## Aktuelles Projektmodell

Die iPad-PWA v28 verwendet dasselbe vereinfachte Projektmodell wie die aktuelle Windows-Buchuhr:

- Im Projekt werden ausschließlich TXT-Arbeitsdateien verwaltet.
- Dateien, die auf der Uhr liegen, befinden sich im Unterordner `Uhr`.
- Eine Uhrdatei heißt beispielsweise `14.25 Brief an Johanne.txt`.
- Beim Verschieben auf der Uhr wird der Zeitpräfix entsprechend der Uhrposition geändert.
- Verlässt eine Datei die Uhr, wird der Zeitpräfix entfernt.
- `Uhr → Stehsatz` verschiebt die TXT-Datei nach `Stehsatz`.
- `Uhr → Papierkorb` verschiebt die TXT-Datei nach `Papierkorb`.
- `Uhr → Projekt-Root` verschiebt die TXT-Datei in den Projektordner.
- Eine neue Datei über `+` im Stehsatz wird unmittelbar als TXT-Datei im Ordner `Stehsatz` angelegt.
- Die Ordner `Uhr`, `Stehsatz` und `Papierkorb` sind Bestandteile des Projekts; die PWA-Programmdateien selbst gehören nicht in einen Projektordner.

Alte Projekte können beim Öffnen migriert werden. TXT- und MD-Bestände lassen sich auf dem iPad in das neue TXT-Modell übernehmen. Alte DOCX-Dateien können Safari nicht zuverlässig selbst in reinen Text umwandeln. Falls ein altes Projekt noch eine solche DOCX-Datei benötigt, sollte es einmal mit der aktuellen Windows-Buchuhr geöffnet und dort migriert werden.

## Lokaler PC-Test

`serve.py` starten:

```powershell
python serve.py
```

Dann `http://localhost:8080` öffnen und diese Adresse zusätzlich als SPA-Redirect-URI registrieren. Die installierbare iPad-Fassung benötigt trotzdem eine HTTPS-Veröffentlichung.


## Hinweis zu den folgenden Versionsnotizen

Die Abschnitte **„Änderung in v2“ bis „Änderung in v27“** dokumentieren den damaligen Entwicklungsstand. Dort können deshalb frühere Ordnernamen wie `Dateien`, frühere Dateiformate oder inzwischen ersetzte Bedienregeln genannt werden. Für die aktuelle Arbeitsweise gelten die Abschnitte **„OneDrive-Projekte“**, **„Bedienung“**, **„Aktuelles Projektmodell“** und **„Projektmodell ab v28“**.

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


## Änderung in v18

- Canvas-Ereignisse vollständig neu aufgebaut
- Dateisymbol und Dateititel besitzen eine gemeinsame unsichtbare Trefferfläche
- zwei Klicks beziehungsweise zweimaliges Antippen öffnen den internen Editor
- Ziehen beginnt erst nach klarer Bewegung
- Pointer-Capture wird erst beim tatsächlichen Ziehen gesetzt
- beim Loslassen, Abbruch oder Fokusverlust endet das Ziehen sicher
- beim Ziehen bleibt der ursprüngliche Winkelabstand zum Mauszeiger erhalten; Datei und Fortschrittsbogen springen nicht weg
- Kontextmenü mit der Maus nur per Rechtsklick
- Langdruck-Kontextmenü ausschließlich auf Touch
- Rasterdialog ausschließlich per Doppelklick auf einen Rasterstrich


## Änderung in v20

- jeder Rasterstrich besitzt eine deutlich größere unsichtbare Trefferfläche
- zweimaliges Antippen auf jedem Rasterstrich öffnet zuverlässig den Rastertitel
- einfacher Tipp links zeigt die Datei rechts als Textvorschau
- einfacher Tipp rechts zeigt die Datei links als Textvorschau
- ein weiterer Tipp auf eine andere Datei wechselt den Vorschauinhalt
- Doppeltipp auf Dateien links oder rechts öffnet den vollständigen Editor
- Tipp auf freien Canvas schließt die Vorschau


## Änderung in v21

- Doppeltipp in einen freien Bereich der rechten Stehsatzleiste erzeugt eine Trennlinie
- Doppelklick mit der Maus in den freien Bereich erzeugt ebenfalls eine Trennlinie
- Doppeltipp oder Doppelklick auf eine Trennlinie öffnet die Betitelung
- F2 beziehungsweise der Umbenennen-Knopf betitelt eine ausgewählte Trennlinie
- die Trennlinie behält beim Betiteln ihre Position im Stehsatz


## Änderung in v22

- Doppeltipp im Vorschautext schaltet das Fenster an Ort und Stelle in den Editiermodus
- Cursor wird anhand der Doppeltipp-Position in den Text gesetzt
- virtuelle Tastatur wird durch Fokussieren des Textfeldes geöffnet
- Speichern und Abbrechen erscheinen im Vorschaufenster
- langes Drücken auf eine Trennlinie aktiviert das Verschieben innerhalb des Stehsatzes
- nur der mittlere Titelbereich der Trennlinie ist per Doppeltipp editierbar
- rechte Leiste: deutliches Wischen nach links entfernt Textdatei oder Trennlinie aus dem Stehsatz
- linke Leiste: deutliches Wischen nach rechts löscht die Datei aus OneDrive und aus dem Projekt


## Änderung in v24

- Grundlage wieder v22; die global blockierende Touch-Ende-Regel aus v23 ist nicht enthalten
- Dateien in beiden Leisten lassen sich wieder auswählen und öffnen die gegenüberliegende Vorschau
- Safari-Zoom wird nur beim zweiten Tipp direkt im Vorschautext verhindert
- Dateien und Trennlinien in der rechten Leiste lassen sich nach 560 ms langem Drücken verschieben
- Verschieben funktioniert über die gesamte Höhe der rechten Leiste
- automatische Scrollbewegung am oberen und unteren Rand während des Verschiebens
- neue Reihenfolge wird beim Loslassen in der Projektdatei gespeichert
- der nach einem Ziehen erzeugte Klick wird unterdrückt


## Änderung in v25

- Projektzentrale statt fest eingetragenem Einzelprojekt
- bevorzugter OneDrive-Projektstamm `Buch-Uhr`
- Projekte können in der PWA neu angelegt und ausgewählt werden
- zuletzt verwendete Projekte stehen oben
- damaliger Stand: `Dateien` und `Stehsatz` wurden als projektinterne Ordner behandelt
- relative Projektpfade aus der Windows-Version werden direkt verstanden


## Änderung in v26

- PWA-interne Laufzeitfelder wie `_webItem` und `_relative` werden nicht mehr in `Buch-Uhr.project.json` gespeichert
- das gemeinsame Projektformat bleibt dadurch für die Windows-Version sauber und portabel


## Änderung in v28

- Projektformat bleibt `version: 2`.
- `title_color_rules` aus der Windows-Buchuhr werden gelesen und gespeichert.
- Signalwörter werden auf Uhr und im Stehsatz jeweils nur an ihrer Fundstelle farbig und fett dargestellt; mehrere Farben pro Titel sind möglich.
- Im Kontextmenü steht `Farbe` für die projektweiten Signalwort-Farbregeln zur Verfügung.
- Jedes Projekt besitzt den Ordner `Papierkorb`. Gelöschte Projektdateien werden dorthin verschoben; die Inhalte werden in der linken Dateiliste angezeigt.
- Wird eine Datei im Projekt-Papierkorb erneut gelöscht, wird sie endgültig aus OneDrive entfernt.


## Projektmodell ab v28

Die iPad-PWA verwendet dasselbe vereinfachte Projektmodell wie Buchuhr v70 unter Windows:

- Im Projekt werden ausschließlich TXT-Arbeitsdateien verwaltet.
- Der bisherige Ordner `Dateien` heißt `Uhr`.
- Eine Datei auf der Uhr liegt als `HH.MM Dateiname.txt` im Ordner `Uhr`.
- Verlässt eine Datei die Uhr, wird der Zeitpräfix entfernt.
- Uhr → Stehsatz verschiebt die TXT-Datei nach `Stehsatz`.
- Uhr → Papierkorb verschiebt die TXT-Datei nach `Papierkorb`.
- Uhr → linke Projektleiste verschiebt die TXT-Datei ins Projekt-Root.
- Eine neue Datei über `+` im Stehsatz wird direkt als TXT im Ordner `Stehsatz` angelegt.
- Die Titelfarbregeln aus v27 bleiben erhalten.

### Migration älterer Projekte

TXT- und MD-Bestände können auf dem iPad automatisch in das neue TXT-Modell übernommen werden.
Alte DOCX-Dateien lassen sich in Safari ohne zusätzlichen Word-Konverter nicht zuverlässig in reinen Text umwandeln.
Wenn beim Öffnen eines älteren Projekts noch eine solche DOCX-Datei übrig ist, öffne das Projekt einmal mit der aktuellen Windows-Buchuhr. Danach kann v28 das migrierte Projekt normal weiterverwenden.


## Änderung in v29

- Der OneDrive-Projektstamm wird als Elternordner der einzelnen Buch-Uhr-Projekte behandelt.
- Wird versehentlich direkt ein Projektordner mit `Buch-Uhr.project.json` als Projektstamm eingetragen, erscheint eine verständliche Fehlermeldung.
- Windows-Backslashes im Projektstamm werden automatisch in Graph-Pfade umgewandelt.
- Verbindungsfehler bleiben über den roten Status `Fehler` abrufbar; ein Klick zeigt die vollständige Microsoft-Graph-/Anmeldeantwort.
- Fehler direkt nach „Speichern und anmelden“ werden nicht mehr als unbehandelte Promise verworfen, sondern sichtbar angezeigt.


## Änderung in v30

- Dateititel an der Uhr werden näher an der Uhr platziert und nicht mehr am SVG-Rand abgeschnitten.
- Farbige Signalwörter bleiben erhalten, erzeugen aber keine zusätzlichen Umbruchstellen innerhalb eines Wortes mehr.
- Umbruch erfolgt nur noch zwischen ganzen Wörtern; bis zu drei Zeilen bleiben sichtbar.
- Titel oberhalb und unterhalb der Uhr werden mittig zum Dateisymbol ausgerichtet.


## Änderung in v31.3

- Uhr-Dateititel werden als natives SVG gerendert; `foreignObject`/HTML entfällt an dieser Stelle.
- Umbruch erfolgt nur zwischen ganzen Wörtern.
- Titel oben und unten stehen radial außerhalb des Dateisymbols und überlagern es nicht mehr.
- Der lokale Server kann über den neuen versteckten Starter ohne sichtbares CMD-Fenster betrieben werden.


## Änderung in v31.4

Nur Diagnose: Vor dem ersten Microsoft-Login zeigt die PWA Client-ID, Mandant, Authority, Authorize-Endpunkt, Redirect-URI, Redirect-Startseite, Scopes, aktuelle URL und Browserkennung an. Erst nach Bestätigung wird zu Microsoft weitergeleitet. Keine Änderung an OneDrive-, Projekt-, Synchronisations- oder Uhrlogik.
