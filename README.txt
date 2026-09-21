KUNDENKARTE FÜR ORDERTEMPLATE.XLSM
=================================

Funktion
--------
- Liest im Blatt OrderSP die Kundendaten aus W3:W7.
- Aktualisiert die Karte automatisch nach der Kundensuche.
- Zeigt den Standort im Excel-Aufgabenbereich.
- Öffnet auf Wunsch Google Maps oder die Routenplanung.
- Verändert die vorhandene Arbeitsmappe, Makros und Druckbereiche nicht.

Dateien
-------
- manifest.xml: fertiges Excel-Add-in-Manifest.
- manifest-template.xml: Vorlage für eine spätere Änderung der Webadresse.
- docs/: öffentlich bereitzustellende HTTPS-Dateien.

Bereitstellung
--------------
Die Weboberfläche wird über GitHub Pages veröffentlicht. Zur Installation wird
die fertige Datei manifest.xml in Excel als Office-Add-in bereitgestellt.

Datenschutz
-----------
Die Adresse des ausgewählten Kunden wird für die Standortsuche an den
OpenStreetMap-Geocodingdienst Nominatim übertragen. Die Add-in-Seite speichert
keine Kundendaten. Bereits gefundene Koordinaten werden lokal im Add-in-Cache
des verwendeten Rechners gespeichert.
