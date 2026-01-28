# VTiger Products Importer

Tampermonkey-Script zum Importieren von Produktdaten verschiedener Anbieter direkt in VTiger.

## Installation

### Voraussetzungen
- Browser: Chrome, Firefox, Edge oder Safari
- [Tampermonkey](https://www.tampermonkey.net/) Browser-Erweiterung

### Script installieren

**Methode 1: Direkt-Installation (empfohlen)**

Klicke auf diesen Link und bestaetigen die Installation in Tampermonkey:

[vtiger-importer.user.js installieren](https://github.com/HWW24-Office/products-importer/raw/main/vtiger-importer.user.js)

**Methode 2: Manuell**

1. Oeffne Tampermonkey im Browser
2. Klicke auf "Neues Script erstellen"
3. Kopiere den Inhalt von `vtiger-importer.user.js` hinein
4. Speichere mit Strg+S

## Benutzung

1. Oeffne VTiger: https://vtiger.hardwarewartung.com
2. Klicke auf den **"Importer"**-Button (in der Navigation oder unten rechts)
3. Waehle den gewuenschten Importer-Tab:
   - **Axians List** - Excel-Preislisten von Axians
   - **Technogroup List** - Excel-Preislisten von Technogroup
   - **Technogroup PDF** - PDF-Angebote von Technogroup
   - **Parkplace Excel** - Excel-Angebote von Park Place Technologies

## Funktionen

### Axians & Technogroup List
- Excel-Datei hochladen (Drag & Drop oder Klick)
- Produkte suchen und filtern nach Hersteller
- SLA auswaehlen
- Laufzeit und Preis-Multiplikator anpassen
- Land auswaehlen (DE/AT/CH)
- Warenkorb mit Bearbeitung
- CSV-Export fuer VTiger-Import

### Technogroup PDF
- PDF hochladen (Drag & Drop)
- Automatische Erkennung von Produkten, Seriennummern, SLA, Datum
- Globale Einstellungen fuer Hersteller, Land, SLA
- Editierbare Vorschau
- CSV-Export

### Parkplace Excel
- Excel-Angebot hochladen
- Automatische Erkennung der Tabellenstruktur
- Zusammenfuehrung gleicher Produkte
- Included-Items werden gruppiert
- CSV-Export

## Updates

Das Script aktualisiert sich automatisch, wenn eine neue Version verfuegbar ist.
Du kannst manuell nach Updates suchen: Tampermonkey > Dashboard > Tab "Installierte Scripts" > Auf Update pruefen.

## Changelog

### v1.0.0
- Initiale Version
- Axians List Importer
- Technogroup List Importer
- Technogroup PDF Importer
- Parkplace Excel Importer
