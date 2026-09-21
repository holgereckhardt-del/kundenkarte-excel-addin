# Kundenkarte für Excel

Dieses Office-Add-in zeigt die Anschrift des in `OrderSP` ausgewählten Kunden in einem Aufgabenbereich von Excel.

## Funktionen

- liest Kundennummer, Name und Anschrift aus `OrderSP!W3:W7`
- aktualisiert sich automatisch nach der vorhandenen Kundensuche
- zeigt den Standort über OpenStreetMap
- öffnet Standort oder Route in Google Maps
- verändert die vorhandene Arbeitsmappe und ihre VBA-Makros nicht

## Installation

Die Excel-Installation erfolgt mit der Datei `manifest.xml`. Die Weboberfläche wird automatisch über GitHub Pages veröffentlicht.

## Datenschutz

Für die Standortsuche wird die Kundenanschrift an den OpenStreetMap-Geocodingdienst Nominatim übertragen. Die Add-in-Seite speichert keine Kundendaten. Bereits gefundene Koordinaten werden ausschließlich lokal im Add-in-Cache des Rechners gespeichert.
