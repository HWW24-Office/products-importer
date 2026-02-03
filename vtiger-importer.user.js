// ==UserScript==
// @name         VTiger Products Importer
// @namespace    https://vtiger.hardwarewartung.com
// @version      1.10.4
// @description  Import-Tools fuer Axians, Parkplace, Technogroup direkt in VTiger
// @author       Hardwarewartung
// @match        https://vtiger.hardwarewartung.com/*
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.14.305/pdf.min.js
// @require      https://unpkg.com/msgreader@1.0.1/dist/MsgReader.js
// @updateURL    https://raw.githubusercontent.com/HWW24-Office/products-importer/main/vtiger-importer.user.js
// @downloadURL  https://raw.githubusercontent.com/HWW24-Office/products-importer/main/vtiger-importer.user.js
// ==/UserScript==

(function() {
    'use strict';

    const SCRIPT_VERSION = '1.10.4';
    console.log('[Products Importer] Version ' + SCRIPT_VERSION + ' geladen');

    // PDF.js Worker konfigurieren
    if (typeof pdfjsLib !== 'undefined') {
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.14.305/pdf.worker.min.js';
    }

    // ============================================
    // STYLES
    // ============================================
    const styles = `
        #importer-modal-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.6);
            z-index: 99999;
        }
        #importer-modal {
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 90%;
            max-width: 1200px;
            height: 85%;
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            z-index: 100000;
            display: flex;
            flex-direction: column;
            overflow: hidden;
        }
        #importer-modal-header {
            background: #1d8d9f;
            color: white;
            padding: 15px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        #importer-modal-header h2 {
            margin: 0;
            font-size: 18px;
        }
        #importer-modal-close {
            background: none;
            border: none;
            color: white;
            font-size: 24px;
            cursor: pointer;
            padding: 0 5px;
        }
        #importer-modal-close:hover {
            opacity: 0.8;
        }
        #importer-tabs {
            display: flex;
            background: #f5f5f5;
            border-bottom: 1px solid #ddd;
        }
        .importer-tab {
            padding: 12px 20px;
            cursor: pointer;
            border: none;
            background: none;
            font-size: 14px;
            color: #555;
            border-bottom: 3px solid transparent;
            transition: all 0.2s;
        }
        .importer-tab:hover {
            background: #e8e8e8;
        }
        .importer-tab.active {
            color: #1d8d9f;
            border-bottom-color: #1d8d9f;
            background: #fff;
        }
        #importer-content {
            flex: 1;
            overflow-y: auto;
            padding: 20px;
        }
        .importer-panel {
            display: none;
        }
        .importer-panel.active {
            display: block;
        }
        /* Form Styles */
        .imp-form-group {
            margin-bottom: 15px;
        }
        .imp-form-group label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
        }
        .imp-form-group input,
        .imp-form-group select,
        .imp-form-group button {
            width: 100%;
            padding: 8px;
            box-sizing: border-box;
        }
        .imp-form-group button {
            background-color: #1d8d9f;
            color: white;
            border: none;
            cursor: pointer;
            margin-top: 5px;
        }
        .imp-form-group button:hover {
            background-color: #166c79;
        }
        .imp-drop-zone {
            border: 2px dashed #1d8d9f;
            padding: 20px;
            text-align: center;
            margin: 10px 0;
            cursor: pointer;
            transition: background 0.2s;
        }
        .imp-drop-zone:hover,
        .imp-drop-zone.hover {
            background-color: #f0f9fa;
        }
        .imp-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
            font-size: 12px;
        }
        .imp-table th,
        .imp-table td {
            border: 1px solid #ddd;
            padding: 6px;
            text-align: left;
        }
        .imp-table th {
            background: #f5f5f5;
        }
        .imp-table .editable {
            background: #fffef0;
        }
        .imp-row-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
        }
        .imp-btn-danger {
            background-color: #d9534f !important;
        }
        .imp-btn-danger:hover {
            background-color: #c9302c !important;
        }
        .imp-hidden {
            display: none;
        }
        .orange-product {
            background-color: #ffcc99;
        }
        .missing-price {
            color: red;
        }
    `;

    // Style einfuegen
    const styleEl = document.createElement('style');
    styleEl.textContent = styles;
    document.head.appendChild(styleEl);

    // ============================================
    // MODAL HTML
    // ============================================
    const modalHTML = `
        <div id="importer-modal-overlay">
            <div id="importer-modal">
                <div id="importer-modal-header">
                    <h2>Products Importer <span id="importer-version" style="font-size:11px;color:#888;font-weight:normal;"></span></h2>
                    <button id="importer-modal-close">&times;</button>
                </div>
                <div id="importer-tabs">
                    <button class="importer-tab active" data-panel="axians">Axians List</button>
                    <button class="importer-tab" data-panel="axians-email">Axians Email</button>
                    <button class="importer-tab" data-panel="itris-email">ITRIS Email</button>
                    <button class="importer-tab" data-panel="technogroup">Technogroup List</button>
                    <button class="importer-tab" data-panel="technogroup-pdf">Technogroup PDF</button>
                    <button class="importer-tab" data-panel="parkplace">Parkplace Excel</button>
                    <button class="importer-tab" data-panel="parkplace-pdf">Parkplace PDF</button>
                    <button class="importer-tab" data-panel="dis-pdf">DIS PDF</button>
                    <button class="importer-tab" data-panel="ids-pdf">IDS PDF</button>
                </div>
                <div id="importer-content">
                    <!-- Axians Panel -->
                    <div class="importer-panel active" id="panel-axians"></div>
                    <!-- Axians Email Panel -->
                    <div class="importer-panel" id="panel-axians-email"></div>
                    <!-- ITRIS Email Panel -->
                    <div class="importer-panel" id="panel-itris-email"></div>
                    <!-- Technogroup Panel -->
                    <div class="importer-panel" id="panel-technogroup"></div>
                    <!-- Technogroup PDF Panel -->
                    <div class="importer-panel" id="panel-technogroup-pdf"></div>
                    <!-- Parkplace Panel -->
                    <div class="importer-panel" id="panel-parkplace"></div>
                    <!-- Parkplace PDF Panel -->
                    <div class="importer-panel" id="panel-parkplace-pdf"></div>
                    <!-- DIS PDF Panel -->
                    <div class="importer-panel" id="panel-dis-pdf"></div>
                    <!-- IDS PDF Panel -->
                    <div class="importer-panel" id="panel-ids-pdf"></div>
                </div>
            </div>
        </div>
    `;

    // Modal einfuegen
    document.body.insertAdjacentHTML('beforeend', modalHTML);

    // Version im Header anzeigen
    const versionSpan = document.getElementById('importer-version');
    if (versionSpan) {
        versionSpan.textContent = 'v' + SCRIPT_VERSION;
    }

    // ============================================
    // FLOATING BUTTON (nur im Products-Modul)
    // ============================================
    function addFloatingButton() {
        // Pruefen ob wir im Products-Modul sind
        function isProductsModule() {
            const url = window.location.href.toLowerCase();
            return url.includes('module=products') ||
                   url.includes('module%3dproducts') ||
                   url.includes('/products/');
        }

        // Button erstellen
        function createButton() {
            if (document.getElementById('open-importer-btn')) return; // Bereits vorhanden

            const floatBtn = document.createElement('button');
            floatBtn.id = 'open-importer-btn';
            floatBtn.textContent = 'Importer';
            floatBtn.style.cssText = `
                position: fixed;
                bottom: 20px;
                right: 20px;
                z-index: 99998;
                background: #1d8d9f;
                color: white;
                border: none;
                padding: 12px 20px;
                border-radius: 5px;
                cursor: pointer;
                font-size: 14px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.2);
                transition: background 0.2s, transform 0.2s;
            `;
            floatBtn.addEventListener('mouseenter', () => {
                floatBtn.style.background = '#166c79';
                floatBtn.style.transform = 'scale(1.05)';
            });
            floatBtn.addEventListener('mouseleave', () => {
                floatBtn.style.background = '#1d8d9f';
                floatBtn.style.transform = 'scale(1)';
            });
            floatBtn.addEventListener('click', (e) => {
                e.preventDefault();
                openModal();
            });
            document.body.appendChild(floatBtn);
        }

        // Button entfernen
        function removeButton() {
            const btn = document.getElementById('open-importer-btn');
            if (btn) btn.remove();
        }

        // Button anzeigen/verstecken basierend auf Modul
        function updateButtonVisibility() {
            if (isProductsModule()) {
                createButton();
            } else {
                removeButton();
            }
        }

        // Initial pruefen
        updateButtonVisibility();

        // Bei URL-Aenderungen pruefen (fuer Single-Page-Navigation)
        let lastUrl = location.href;
        new MutationObserver(() => {
            if (location.href !== lastUrl) {
                lastUrl = location.href;
                updateButtonVisibility();
            }
        }).observe(document.body, { subtree: true, childList: true });

        // Auch auf popstate hoeren (Browser-Navigation)
        window.addEventListener('popstate', updateButtonVisibility);
    }

    // ============================================
    // MODAL FUNKTIONEN
    // ============================================
    const overlay = document.getElementById('importer-modal-overlay');
    const closeBtn = document.getElementById('importer-modal-close');
    const tabs = document.querySelectorAll('.importer-tab');

    function openModal() {
        overlay.style.display = 'block';
    }

    function closeModal() {
        overlay.style.display = 'none';
    }

    closeBtn.addEventListener('click', closeModal);
    overlay.addEventListener('click', (e) => {
        if (e.target === overlay) closeModal();
    });

    // Tab-Wechsel
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            document.querySelectorAll('.importer-panel').forEach(p => p.classList.remove('active'));
            document.getElementById('panel-' + tab.dataset.panel).classList.add('active');
        });
    });

    // ============================================
    // GEMEINSAME HILFSFUNKTIONEN
    // ============================================
    const countryMapping = {
        "DE": "Deutschland",
        "AT": "Österreich",
        "CH": "Schweiz"
    };

    // ============================================
    // LAENDER-NORMALISIERUNG UND UEBERSETZUNG
    // ============================================
    const countryNormalization = {
        // USA-Varianten
        'united states': 'USA', 'united states of america': 'USA', 'us': 'USA', 'u.s.': 'USA', 'u.s.a.': 'USA',
        'vereinigte staaten': 'USA', 'vereinigte staaten von amerika': 'USA',
        // UK-Varianten
        'united kingdom': 'UK', 'great britain': 'UK', 'gb': 'UK', 'großbritannien': 'UK', 'grossbritannien': 'UK', 'england': 'UK',
        // UAE-Varianten
        'united arab emirates': 'UAE', 'vereinigte arabische emirate': 'UAE', 'vae': 'UAE', 'v.a.e.': 'UAE',
        // Deutschland-Varianten
        'germany': 'Germany', 'deutschland': 'Germany', 'de': 'Germany', 'ger': 'Germany',
        // Österreich-Varianten
        'austria': 'Austria', 'oesterreich': 'Austria', 'österreich': 'Austria', 'at': 'Austria', 'aut': 'Austria',
        // Schweiz-Varianten
        'switzerland': 'Switzerland', 'schweiz': 'Switzerland', 'suisse': 'Switzerland', 'svizzera': 'Switzerland', 'ch': 'Switzerland',
        // Weitere Laender
        'france': 'France', 'frankreich': 'France', 'fr': 'France',
        'netherlands': 'Netherlands', 'niederlande': 'Netherlands', 'holland': 'Netherlands', 'nl': 'Netherlands',
        'belgium': 'Belgium', 'belgien': 'Belgium', 'be': 'Belgium',
        'spain': 'Spain', 'spanien': 'Spain', 'es': 'Spain',
        'italy': 'Italy', 'italien': 'Italy', 'it': 'Italy',
        'poland': 'Poland', 'polen': 'Poland', 'pl': 'Poland',
        'czech republic': 'Czech Republic', 'tschechien': 'Czech Republic', 'tschechische republik': 'Czech Republic', 'cz': 'Czech Republic',
        'hungary': 'Hungary', 'ungarn': 'Hungary', 'hu': 'Hungary',
        'romania': 'Romania', 'rumaenien': 'Romania', 'rumänien': 'Romania', 'ro': 'Romania',
        'sweden': 'Sweden', 'schweden': 'Sweden', 'se': 'Sweden',
        'denmark': 'Denmark', 'daenemark': 'Denmark', 'dänemark': 'Denmark', 'dk': 'Denmark',
        'norway': 'Norway', 'norwegen': 'Norway', 'no': 'Norway',
        'finland': 'Finland', 'finnland': 'Finland', 'fi': 'Finland'
    };

    const countryTranslations = {
        // Normalisierter Name -> { en: englisch, de: deutsch }
        'USA': { en: 'USA', de: 'USA' },
        'UK': { en: 'UK', de: 'UK' },
        'UAE': { en: 'UAE', de: 'UAE' },
        'Germany': { en: 'Germany', de: 'Deutschland' },
        'Austria': { en: 'Austria', de: 'Österreich' },
        'Switzerland': { en: 'Switzerland', de: 'Schweiz' },
        'France': { en: 'France', de: 'Frankreich' },
        'Netherlands': { en: 'Netherlands', de: 'Niederlande' },
        'Belgium': { en: 'Belgium', de: 'Belgien' },
        'Spain': { en: 'Spain', de: 'Spanien' },
        'Italy': { en: 'Italy', de: 'Italien' },
        'Poland': { en: 'Poland', de: 'Polen' },
        'Czech Republic': { en: 'Czech Republic', de: 'Tschechien' },
        'Hungary': { en: 'Hungary', de: 'Ungarn' },
        'Romania': { en: 'Romania', de: 'Rumaenien' },
        'Sweden': { en: 'Sweden', de: 'Schweden' },
        'Denmark': { en: 'Denmark', de: 'Daenemark' },
        'Norway': { en: 'Norway', de: 'Norwegen' },
        'Finland': { en: 'Finland', de: 'Finnland' }
    };

    // Normalisiert Laendernamen auf standardisierte Form
    function normalizeCountry(countryName) {
        if (!countryName) return countryName;
        const trimmed = countryName.trim();
        const lower = trimmed.toLowerCase();

        // Bereits normalisiert?
        if (countryTranslations[trimmed]) return trimmed;

        // Suche in Normalisierungstabelle
        if (countryNormalization[lower]) return countryNormalization[lower];

        // Keine Uebereinstimmung - original zurueckgeben
        return trimmed;
    }

    // Gibt den Laendernamen in der gewuenschten Sprache zurueck
    function getCountryForLanguage(countryName, language) {
        const normalized = normalizeCountry(countryName);
        const translation = countryTranslations[normalized];
        if (translation) {
            return translation[language] || normalized;
        }
        return countryName; // Unbekanntes Land - unuebersetzt lassen
    }

    function setupDropZone(dropZone, fileInput) {
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('hover');
        });
        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('hover');
        });
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('hover');
            fileInput.files = e.dataTransfer.files;
            fileInput.dispatchEvent(new Event('change'));
        });
        dropZone.addEventListener('click', () => fileInput.click());
    }

    function downloadCSV(rows, filename) {
        const bom = "\uFEFF";
        const csvData = new Blob([bom + rows.join('\n')], { type: 'text/csv;charset=utf-8;' });
        const csvUrl = URL.createObjectURL(csvData);
        const link = document.createElement('a');
        link.href = csvUrl;
        link.download = filename;
        link.click();
        URL.revokeObjectURL(csvUrl);
    }

    // ============================================
    // SPRACH-UMSCHALTUNG (DE <-> EN)
    // ============================================
    const descriptionTranslations = {
        // Description-Texte DE -> EN
        'inkl.:': 'incl.:',
        'Service Ende:': 'Service End:',
        'Laufzeit:': 'Duration:',
        'Monate': 'months',
        // EN -> DE
        'incl.:': 'inkl.:',
        'Service End:': 'Service Ende:',
        'Duration:': 'Laufzeit:',
        'months': 'Monate'
    };

    function toggleLanguage(tableId, countryInputClass, currentLang) {
        const newLang = currentLang === 'de' ? 'en' : 'de';
        const tbody = document.querySelector(`#${tableId} tbody`);
        if (!tbody) return newLang;

        // Country-Inputs uebersetzen (verwendet neue Normalisierung)
        if (countryInputClass) {
            document.querySelectorAll(`.${countryInputClass}`).forEach(input => {
                const val = input.value.trim();
                input.value = getCountryForLanguage(val, newLang);
            });
        }

        // Description-Spalte uebersetzen (normalerweise Spalte 8)
        tbody.querySelectorAll('tr').forEach(row => {
            const descCell = row.cells[8];
            if (descCell) {
                let text = descCell.innerHTML;
                // Ersetze Description-Begriffe je nach Richtung
                Object.entries(descriptionTranslations).forEach(([from, to]) => {
                    if (currentLang === 'de') {
                        // DE -> EN
                        if (from === 'inkl.:' || from === 'Service Ende:' || from === 'Laufzeit:' || from === 'Monate') {
                            text = text.split(from).join(to);
                        }
                    } else {
                        // EN -> DE
                        if (from === 'incl.:' || from === 'Service End:' || from === 'Duration:' || from === 'months') {
                            text = text.split(from).join(to);
                        }
                    }
                });

                // Auch Laendernamen in Description uebersetzen
                Object.keys(countryTranslations).forEach(normalized => {
                    const trans = countryTranslations[normalized];
                    if (currentLang === 'de' && trans.de !== trans.en) {
                        // DE -> EN: z.B. "Deutschland" -> "Germany"
                        text = text.split(trans.de).join(trans.en);
                    } else if (currentLang === 'en' && trans.de !== trans.en) {
                        // EN -> DE: z.B. "Germany" -> "Deutschland"
                        text = text.split(trans.en).join(trans.de);
                    }
                });

                descCell.innerHTML = text;
            }

            // Country-Spalte (contenteditable) uebersetzen - falls keine Input-Felder
            const countryCell = row.cells[11];
            if (countryCell && !countryCell.querySelector('input')) {
                const val = countryCell.textContent.trim();
                countryCell.textContent = getCountryForLanguage(val, newLang);
            }
        });

        return newLang;
    }

    // ============================================
    // MSG-DATEI LESEN (Outlook E-Mails)
    // ============================================
    async function readMsgFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const arrayBuffer = e.target.result;

                    // MsgReader kann als Modul oder global verfuegbar sein
                    let MsgReaderClass = window.MsgReader;
                    if (!MsgReaderClass && typeof MsgReader !== 'undefined') {
                        MsgReaderClass = MsgReader;
                    }
                    if (!MsgReaderClass) {
                        throw new Error('MsgReader Bibliothek nicht geladen. Bitte Seite neu laden.');
                    }

                    const msgReader = new MsgReaderClass(arrayBuffer);
                    const fileData = msgReader.getFileData();

                    // DEBUG: Zeige MsgReader-Methoden und alle Daten
                    console.log('=== MsgReader Object ===');
                    console.log('MsgReader methods:', Object.getOwnPropertyNames(Object.getPrototypeOf(msgReader)));
                    console.log('MsgReader own keys:', Object.keys(msgReader));

                    // Untersuche das ds (DataStream) Objekt
                    if (msgReader.ds) {
                        console.log('msgReader.ds keys:', Object.keys(msgReader.ds));
                    }

                    // Untersuche fileData genauer
                    if (msgReader.fileData) {
                        console.log('msgReader.fileData keys:', Object.keys(msgReader.fileData));
                        // Suche nach allen String-Properties mit Inhalt
                        for (const key of Object.keys(msgReader.fileData)) {
                            const val = msgReader.fileData[key];
                            if (typeof val === 'string' && val.length > 50) {
                                console.log(`msgReader.fileData.${key} (${val.length} chars):`, val.substring(0, 300));
                            }
                        }
                    }
                    console.log('========================');

                    // DEBUG: Zeige alle verfuegbaren Properties
                    console.log('=== MSG FileData DEBUG ===');
                    console.log('Alle Keys:', Object.keys(fileData));
                    for (const key of Object.keys(fileData)) {
                        const val = fileData[key];
                        if (typeof val === 'string') {
                            console.log(`${key} (string, ${val.length} chars):`, val.substring(0, 200));
                        } else if (val instanceof Uint8Array || val instanceof ArrayBuffer) {
                            console.log(`${key} (binary, ${val.byteLength || val.length} bytes)`);
                        } else if (Array.isArray(val)) {
                            console.log(`${key} (array, ${val.length} items):`, val);
                        } else {
                            console.log(`${key} (${typeof val}):`, val);
                        }
                    }
                    // DEBUG: Zeige Attachment-Details
                    if (fileData.attachments && fileData.attachments.length > 0) {
                        console.log('=== ATTACHMENTS DEBUG ===');
                        fileData.attachments.forEach((att, idx) => {
                            console.log(`Attachment ${idx} - ALLE KEYS:`, att);
                            // Versuche Attachment-Content zu laden
                            try {
                                const attContent = msgReader.getAttachment(idx);
                                if (attContent) {
                                    console.log(`  getAttachment(${idx}):`, {
                                        keys: Object.keys(attContent),
                                        contentLength: attContent.content ? attContent.content.length : 0
                                    });
                                }
                            } catch (e) {
                                console.log(`  getAttachment(${idx}) error:`, e.message);
                            }
                        });
                        console.log('=========================');
                    }

                    // Suche nach eingebetteten Nachrichten oder RTF-Body
                    console.log('=== SUCHE NACH BODY-ALTERNATIVEN ===');
                    // Manchmal ist der Body im RTF-Format komprimiert
                    if (fileData.body && typeof fileData.body === 'object') {
                        console.log('Body ist ein Objekt mit Keys:', Object.keys(fileData.body));
                        // Versuche alle numerischen Indizes zu lesen
                        const bodyBytes = [];
                        for (let i = 0; i < 100 && fileData.body[i] !== undefined; i++) {
                            bodyBytes.push(fileData.body[i]);
                        }
                        console.log('Body bytes (erste 100):', bodyBytes);
                    }
                    console.log('=====================================');
                    console.log('==========================');

                    // E-Mail-Daten extrahieren (mit Typ-Pruefung)
                    let bodyText = fileData.body || '';
                    // Falls body kein String ist, versuche zu konvertieren
                    if (typeof bodyText !== 'string') {
                        // ZUERST auf binaere Daten pruefen (ArrayBuffer, Uint8Array)
                        if (bodyText instanceof ArrayBuffer) {
                            try {
                                const decoder = new TextDecoder('utf-8');
                                bodyText = decoder.decode(bodyText);
                            } catch (e) {
                                bodyText = '';
                            }
                        } else if (bodyText instanceof Uint8Array || (bodyText && bodyText.buffer instanceof ArrayBuffer)) {
                            try {
                                const decoder = new TextDecoder('utf-8');
                                bodyText = decoder.decode(bodyText);
                            } catch (e) {
                                bodyText = '';
                            }
                        } else if (bodyText && typeof bodyText === 'object' && bodyText.length !== undefined) {
                            // Array-aehnliches Objekt (z.B. Array von Bytes)
                            try {
                                const uint8 = new Uint8Array(bodyText);
                                const decoder = new TextDecoder('utf-8');
                                bodyText = decoder.decode(uint8);
                            } catch (e) {
                                bodyText = '';
                            }
                        } else if (bodyText && bodyText.toString) {
                            bodyText = bodyText.toString();
                        } else {
                            bodyText = '';
                        }
                    }

                    let bodyHTML = fileData.bodyHTML || '';
                    if (typeof bodyHTML !== 'string') {
                        // Gleiche Logik fuer HTML
                        if (bodyHTML instanceof ArrayBuffer || bodyHTML instanceof Uint8Array) {
                            try {
                                const decoder = new TextDecoder('utf-8');
                                bodyHTML = decoder.decode(bodyHTML);
                            } catch (e) {
                                bodyHTML = '';
                            }
                        } else {
                            bodyHTML = '';
                        }
                    }

                    const result = {
                        subject: (fileData.subject || '').toString(),
                        from: (fileData.senderName || fileData.senderEmail || '').toString(),
                        body: bodyText,
                        bodyHTML: bodyHTML,
                        attachments: fileData.attachments || []
                    };

                    // Wenn HTML vorhanden, versuche Text zu extrahieren
                    if (!result.body && result.bodyHTML) {
                        const temp = document.createElement('div');
                        temp.innerHTML = result.bodyHTML;
                        result.body = temp.textContent || temp.innerText || '';
                    }

                    // Finale Sicherheitspruefung: result.body MUSS ein String sein
                    if (typeof result.body !== 'string') {
                        // Binaere Daten zuerst pruefen
                        if (result.body instanceof ArrayBuffer || result.body instanceof Uint8Array) {
                            try {
                                const decoder = new TextDecoder('utf-8');
                                result.body = decoder.decode(result.body);
                            } catch (e) {
                                result.body = '';
                            }
                        } else if (result.body && typeof result.body === 'object' && result.body.length !== undefined) {
                            try {
                                const uint8 = new Uint8Array(result.body);
                                const decoder = new TextDecoder('utf-8');
                                result.body = decoder.decode(uint8);
                            } catch (e) {
                                result.body = '';
                            }
                        } else {
                            result.body = '';
                        }
                    }

                    console.log('MSG-Datei gelesen:', result.subject);
                    console.log('Body Laenge:', result.body.length);
                    if (result.body && result.body.length > 0) {
                        console.log('Body Vorschau:', result.body.substring(0, 500) + '...');
                    }

                    resolve(result);
                } catch (error) {
                    console.error('Fehler beim Lesen der MSG-Datei:', error);
                    reject(error);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    // Parkplace-Daten aus E-Mail-Text extrahieren
    function parseParkplaceFromEmail(emailBody) {
        const lines = emailBody.split('\n').map(l => l.trim()).filter(l => l);
        const dataRows = [];

        // Bekannte OEMs
        const knownOEMs = ['NetApp', 'Dell', 'HP', 'HPE', 'IBM', 'Cisco', 'EMC', 'Fujitsu', 'Lenovo', 'Sun', 'Oracle', 'Hitachi', 'Pure Storage', 'Nimble'];
        const datePattern = /(\d{2}-[A-Za-z]{3}-\d{4})/g;
        const pricePattern = /€\s*([\d.,]+)/;
        const lineNumPattern = /^(\d+\.\d+\.?\d*\.?\d*)\s+/;
        const slaPattern = /(\d+x\d+x\w+)/i;

        for (const line of lines) {
            const lineMatch = line.match(lineNumPattern);
            if (!lineMatch) continue;
            if (line.includes('Grand Total')) break;

            const lineNum = lineMatch[1];
            let oem = 'N/A';
            for (const m of knownOEMs) {
                if (line.includes(m)) { oem = m; break; }
            }

            let total = '0';
            const priceMatch = line.match(pricePattern);
            if (priceMatch) total = '€' + priceMatch[1];
            else if (line.toLowerCase().includes('included')) total = 'Included';

            const dates = line.match(datePattern) || [];
            const startDate = dates[0] || '';
            const endDate = dates[1] || '';

            let sla = 'N/A';
            const slaMatch = line.match(slaPattern);
            if (slaMatch) sla = slaMatch[1];

            // Seriennummer
            let serial = '';
            const serialCandidates = line.match(/\b([A-Z0-9]{8,20})\b/gi) || [];
            for (const c of serialCandidates) {
                if (!/\d{2}-[A-Za-z]{3}-\d{4}/.test(c) && !/\d+x\d+x/i.test(c)) {
                    serial = c; break;
                }
            }

            // Location
            let location = '';
            const locMatch = line.match(/([A-Za-z\s]+,\s*[A-Za-z]+)\s+\d{2}-/);
            if (locMatch) location = locMatch[1].trim();

            // Produktname
            let productName = 'N/A';
            if (oem !== 'N/A') {
                const oemIdx = line.indexOf(oem);
                let afterOem = line.substring(oemIdx + oem.length).trim();
                const stopPatterns = [/Parts Tech/i, /ParkView/i, /\d+x\d+x\w+/i, /\d{2}-[A-Za-z]{3}-\d{4}/];
                for (const pattern of stopPatterns) {
                    const match = afterOem.search(pattern);
                    if (match > 0) { afterOem = afterOem.substring(0, match).trim(); break; }
                }
                productName = afterOem.replace(/Parts Tech & Labor/gi, '').trim() || 'N/A';
            }

            dataRows.push({
                line: lineNum, oem, productName, sla, serial,
                qty: 1, location, startDate, endDate, total
            });
        }

        return dataRows;
    }

    // ============================================
    // AXIANS IMPORTER
    // ============================================
    function initAxians() {
        const panel = document.getElementById('panel-axians');
        panel.innerHTML = `
            <h3>Axians List Importer</h3>
            <div class="imp-row-grid">
                <div class="imp-form-group">
                    <label>Excel Datei:</label>
                    <input type="file" id="axians-file" accept=".xlsx,.xls" class="imp-hidden">
                    <div class="imp-drop-zone" id="axians-dropzone">Datei hierher ziehen oder klicken</div>
                    <button id="axians-process">Datei verarbeiten</button>
                </div>
                <div class="imp-form-group">
                    <label>Produkt suchen:</label>
                    <input type="text" id="axians-search" placeholder="Suchbegriff...">
                </div>
                <div class="imp-form-group">
                    <label>Hersteller:</label>
                    <select id="axians-manufacturer"><option value="all">Alle Hersteller</option></select>
                </div>
            </div>
            <div class="imp-row-grid">
                <div class="imp-form-group">
                    <label>Produkte:</label>
                    <select id="axians-products" multiple style="height:150px;"></select>
                </div>
                <div>
                    <div class="imp-form-group">
                        <label>SLA:</label>
                        <select id="axians-sla"></select>
                    </div>
                    <div class="imp-form-group">
                        <label>Laufzeit (Monate):</label>
                        <input type="number" id="axians-duration" value="12">
                    </div>
                    <div class="imp-form-group">
                        <label>Preis-Multiplikator:</label>
                        <input type="number" id="axians-multiplier" value="1.84" step="0.01">
                    </div>
                    <div class="imp-form-group">
                        <label>Land:</label>
                        <select id="axians-country">
                            <option value="DE">Deutschland</option>
                            <option value="AT">Österreich</option>
                            <option value="CH">Schweiz</option>
                        </select>
                    </div>
                </div>
            </div>
            <div class="imp-form-group">
                <button id="axians-add">Zum Warenkorb hinzufuegen</button>
                <button id="axians-clear" class="imp-btn-danger">Warenkorb leeren</button>
            </div>
            <h4>Warenkorb</h4>
            <div style="overflow-x:auto;">
                <table class="imp-table" id="axians-table">
                    <thead>
                        <tr>
                            <th>Product Name</th>
                            <th>Active</th>
                            <th>Manufacturer</th>
                            <th>Category</th>
                            <th>Vendor</th>
                            <th>Unit Price</th>
                            <th>Stock</th>
                            <th>Handler</th>
                            <th>Description</th>
                            <th>Purchase Cost</th>
                            <th>SLA</th>
                            <th>Country</th>
                            <th>Duration</th>
                            <th>Aktion</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
            <button id="axians-download">CSV speichern</button>
        `;

        const slaMappings = {
            "5x9x NBD Response": "5x9xNBD",
            "5x9x NBD FIX": "5x9xNBD fix",
            "5x13x 4 h Response": "5x13x4",
            "5x13x 4 h FIX": "5x13x4 fix",
            "5x13x 6 h FIX": "5x13x6 fix",
            "5x13x 8 h FIX": "5x13x8 fix",
            "5x13x 12 h FIX": "5x13x12 fix",
            "5x13xNBD FIX": "5x13xNBD fix",
            "7x24x4 h Response": "7x24x4",
            "7x24x4 h FIX": "7x24x4 fix",
            "7x24x6 h FIX": "7x24x6 fix",
            "7x24x8 h FIX": "7x24x8 fix",
            "7x24x12 h FIX": "7x24x12 fix"
        };
        const validTabs = ['HP', 'Dell', 'Fujitsu', 'NetApp', 'SUN - Oracle', 'Workstations_Notebooks_Desktops', 'IBM', 'Sonstiges'];
        const excludedEntries = ['Es gelten unsere Angebotsbedingungen', 'Bei Preisanfragen benoetigen wir die IBM Partnummer !!!'];

        let productData = [];
        let filteredProductData = [];
        let cart = [];

        const fileInput = document.getElementById('axians-file');
        const dropZone = document.getElementById('axians-dropzone');
        setupDropZone(dropZone, fileInput);

        document.getElementById('axians-process').addEventListener('click', () => {
            const file = fileInput.files[0];
            if (!file) { alert('Bitte eine Datei hochladen.'); return; }

            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                const productSelector = document.getElementById('axians-products');
                const manufacturerFilter = document.getElementById('axians-manufacturer');
                const slaSelector = document.getElementById('axians-sla');

                productSelector.innerHTML = '';
                slaSelector.innerHTML = '';
                manufacturerFilter.innerHTML = '<option value="all">Alle Hersteller</option>';
                productData = [];

                const sheetNames = workbook.SheetNames.filter(name => validTabs.includes(name.trim()));

                sheetNames.forEach(sheetName => {
                    const sheet = workbook.Sheets[sheetName];
                    if (!sheet) return;
                    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                    if (rows.length === 0) return;

                    const mappingKeys = Object.keys(slaMappings);
                    const slaRowIndex = rows.findIndex(row =>
                        row.some(cell => typeof cell === 'string' && mappingKeys.includes(cell.trim()))
                    );
                    if (slaRowIndex < 0) return;

                    const slaHeaders = rows[slaRowIndex].slice(1, 14);
                    if (slaSelector.options.length === 0) {
                        slaHeaders.forEach((sla, idx) => {
                            const mappedSla = slaMappings[sla] || sla;
                            const option = document.createElement('option');
                            option.value = idx + 1;
                            option.textContent = mappedSla;
                            slaSelector.appendChild(option);
                        });
                    }

                    rows.slice(slaRowIndex + 1).forEach(row => {
                        let productName = row[0] ? row[0].toString().trim() : '';
                        if ((sheetName === 'IBM' || sheetName === 'Sonstiges') && row[1]) {
                            productName += ` - ${row[1].toString().trim()}`;
                        }
                        if (!productName || excludedEntries.includes(productName)) return;

                        const prices = (sheetName === 'IBM' || sheetName === 'Sonstiges')
                            ? row.slice(2, 15) : row.slice(1, 14);
                        const manufacturer = sheetName === 'HP' ? 'HPE' : sheetName;

                        productData.push({ name: productName, prices, manufacturer });

                        if (!Array.from(manufacturerFilter.options).some(opt => opt.value === manufacturer)) {
                            const option = document.createElement('option');
                            option.value = manufacturer;
                            option.textContent = manufacturer;
                            manufacturerFilter.appendChild(option);
                        }
                    });
                });

                updateProductSelector('');
                dropZone.textContent = file.name;
            };
            reader.readAsArrayBuffer(file);
        });

        function updateProductSelector(searchTerm) {
            const productSelector = document.getElementById('axians-products');
            const selectedManufacturer = document.getElementById('axians-manufacturer').value;
            productSelector.innerHTML = '';

            filteredProductData = productData.filter(p =>
                p.name.toLowerCase().includes(searchTerm) &&
                (selectedManufacturer === 'all' || p.manufacturer === selectedManufacturer)
            );

            filteredProductData.forEach((product, index) => {
                const option = document.createElement('option');
                option.value = index;
                option.textContent = `${product.name} (${product.manufacturer})`;
                productSelector.appendChild(option);
            });
        }

        document.getElementById('axians-search').addEventListener('input', (e) => {
            updateProductSelector(e.target.value.toLowerCase());
        });
        document.getElementById('axians-manufacturer').addEventListener('change', () => {
            updateProductSelector(document.getElementById('axians-search').value.toLowerCase());
        });

        document.getElementById('axians-add').addEventListener('click', addToCart);
        document.getElementById('axians-products').addEventListener('dblclick', addToCart);

        function addToCart() {
            const productIndices = Array.from(document.getElementById('axians-products').selectedOptions).map(o => o.value);
            const slaIndex = document.getElementById('axians-sla').value;
            const slaText = document.getElementById('axians-sla').selectedOptions[0]?.text || '';
            const duration = document.getElementById('axians-duration').value;
            const country = countryMapping[document.getElementById('axians-country').value];
            const priceMultiplier = parseFloat(document.getElementById('axians-multiplier').value);

            productIndices.forEach(idx => {
                const product = filteredProductData[idx];
                const price = product.prices[slaIndex - 1];
                if (!price || isNaN(price)) return;

                let unitPriceValue = parseFloat(price) * priceMultiplier * duration;
                if (country === "Schweiz") unitPriceValue *= 1.45;
                const unitPrice = unitPriceValue.toFixed(1);

                let purchaseCost = parseFloat(price) * duration;
                if (country === "Schweiz") purchaseCost *= 1.45;
                purchaseCost = purchaseCost.toFixed(2);

                const isDuplicate = cart.some(item =>
                    item.name === product.name && item.purchaseCost === purchaseCost &&
                    item.duration === duration && item.country === country && item.sla === slaText
                );
                if (isDuplicate) { alert('Produkt bereits im Warenkorb.'); return; }

                cart.push({
                    name: product.name, active: 1, manufacturer: product.manufacturer,
                    category: 'Wartung', vendor: 'Axians IT-Infrastructure Services GmbH',
                    unitPrice, qtyInStock: 999, handler: 'Team Wartung',
                    description: 'S/N:\nService Start:\nService Ende:',
                    purchaseCost, sla: slaText, country, duration, listPrice: 1
                });
                updateTable();
            });
        }

        function updateTable() {
            const tbody = document.querySelector('#axians-table tbody');
            tbody.innerHTML = '';
            cart.forEach((item, index) => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td contenteditable="true" class="editable">${item.name}</td>
                    <td>${item.active}</td>
                    <td contenteditable="true" class="editable">${item.manufacturer}</td>
                    <td>${item.category}</td>
                    <td>${item.vendor}</td>
                    <td contenteditable="true" class="editable">${item.unitPrice}</td>
                    <td>${item.qtyInStock}</td>
                    <td>${item.handler}</td>
                    <td contenteditable="true" class="editable">${item.description.replace(/\n/g, '<br>')}</td>
                    <td>${item.purchaseCost}</td>
                    <td>${item.sla}</td>
                    <td>${item.country}</td>
                    <td>${item.duration}</td>
                    <td><button onclick="this.closest('tr').remove(); window.axiansCart.splice(${index}, 1);" class="imp-btn-danger">X</button></td>
                `;
                tbody.appendChild(row);
            });
            window.axiansCart = cart;
        }

        document.getElementById('axians-clear').addEventListener('click', () => {
            cart = [];
            updateTable();
        });

        document.getElementById('axians-download').addEventListener('click', () => {
            const headers = ["Product Name", "Product Active", "Manufacturer", "Product Category", "Vendor Name", "Unit Price", "Qty. in Stock", "Handler", "Description", "Purchase Cost", "SLA", "Country", "Duration in months", "Listenpreis"];
            const csvRows = [headers.join(';')];
            cart.forEach(item => {
                csvRows.push([
                    item.name, item.active, item.manufacturer, item.category, item.vendor,
                    item.unitPrice, item.qtyInStock, item.handler,
                    `"${item.description.replace(/\n/g, '\n')}"`,
                    item.purchaseCost, item.sla, item.country, item.duration, item.listPrice
                ].join(';'));
            });
            downloadCSV(csvRows, 'crm_import_axians.csv');
        });
    }

    // ============================================
    // AXIANS EMAIL IMPORTER
    // ============================================
    function initAxiansEmail() {
        const panel = document.getElementById('panel-axians-email');
        panel.innerHTML = `
            <h3>Axians Email Importer</h3>
            <p style="color:#666;margin-bottom:15px;">Liest Angebote direkt aus Axians E-Mails (.msg Dateien)</p>
            <div class="imp-row-grid">
                <div class="imp-form-group">
                    <label>Email (.msg Datei):</label>
                    <input type="file" id="axians-email-file" accept=".msg" class="imp-hidden">
                    <div class="imp-drop-zone" id="axians-email-dropzone">MSG-Datei hierher ziehen oder klicken</div>
                    <button id="axians-email-process">Email verarbeiten</button>
                </div>
                <div>
                    <div class="imp-form-group">
                        <label>Laufzeit (Monate):</label>
                        <input type="number" id="axians-email-duration" value="12">
                    </div>
                    <div class="imp-form-group">
                        <label>Preis-Multiplikator:</label>
                        <input type="number" id="axians-email-multiplier" value="1.84" step="0.01">
                    </div>
                    <div class="imp-form-group">
                        <label>Land:</label>
                        <input type="text" id="axians-email-country" value="Deutschland" placeholder="z.B. Deutschland, Österreich...">
                    </div>
                </div>
            </div>
            <div id="axians-email-info" style="margin:15px 0;padding:10px;background:#f5f5f5;border-radius:4px;display:none;">
                <strong>Email-Info:</strong>
                <div id="axians-email-subject"></div>
                <div id="axians-email-reference"></div>
                <div id="axians-email-date"></div>
            </div>
            <div id="axians-email-parsed" style="margin:15px 0;display:none;">
                <h4>Gefundene Produkte:</h4>
                <div id="axians-email-products-list"></div>
                <button id="axians-email-add-all" style="margin-top:10px;">Alle zum Warenkorb hinzufügen</button>
            </div>
            <div class="imp-form-group">
                <button id="axians-email-clear" class="imp-btn-danger">Warenkorb leeren</button>
            </div>
            <h4>Warenkorb</h4>
            <div style="overflow-x:auto;">
                <table class="imp-table" id="axians-email-table">
                    <thead>
                        <tr>
                            <th>Product Name</th>
                            <th>Active</th>
                            <th>Manufacturer</th>
                            <th>Category</th>
                            <th>Vendor</th>
                            <th>Unit Price</th>
                            <th>Stock</th>
                            <th>Handler</th>
                            <th>Description</th>
                            <th>Purchase Cost</th>
                            <th>SLA</th>
                            <th>Country</th>
                            <th>Duration</th>
                            <th>Aktion</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
            <button id="axians-email-download">CSV speichern</button>
        `;

        // SLA-Mapping von Email-Format zu VTiger-Format
        const slaMappings = {
            "5/9/NBD/-": "5x9xNBD",
            "5/9/NBD": "5x9xNBD",
            "5x9xNBD": "5x9xNBD",
            "7/24/4/-": "7x24x4",
            "7/24/4": "7x24x4",
            "7x24x4": "7x24x4",
            "5/13/4/-": "5x13x4",
            "5/13/4": "5x13x4",
            "7/24/6/-": "7x24x6",
            "7/24/6": "7x24x6"
        };

        let parsedProducts = [];
        let cart = [];

        const fileInput = document.getElementById('axians-email-file');
        const dropZone = document.getElementById('axians-email-dropzone');
        setupDropZone(dropZone, fileInput);

        // E-Mail-Text parsen
        function parseAxiansEmail(body, subject) {
            const products = [];

            // Sicherstellen dass body ein String ist
            if (typeof body !== 'string') {
                body = body ? String(body) : '';
            }

            const lines = body.split('\n');

            // Referenznummer extrahieren
            const refMatch = body.match(/Referenz-Nr\.?:\s*(#\d+-\d+[a-z]*\d*)/i) ||
                            subject.match(/(#\d+-\d+[a-z]*\d*)/i);
            const reference = refMatch ? refMatch[1] : '';

            // Standort extrahieren (z.B. "Standort: Österreich")
            const standortMatch = body.match(/Standort:\s*([^\n\r]+)/i);
            let standort = standortMatch ? standortMatch[1].trim() : '';
            // Bereinige Standort (entferne evtl. nachfolgende Texte)
            if (standort) {
                standort = standort.split(/\s{2,}/)[0].trim(); // Stoppe bei mehreren Leerzeichen
            }
            console.log('Gefundener Standort:', standort);

            // Seriennummern extrahieren (verschiedene Formate)
            const serialNumbers = [];
            // Pattern 1: "S/N: XXXXX" oder "S/N XXXXX"
            const snMatches = body.matchAll(/S\/N[:\s]+([A-Z0-9\-]+)/gi);
            for (const match of snMatches) {
                if (match[1] && match[1].length >= 4 && !serialNumbers.includes(match[1])) {
                    serialNumbers.push(match[1]);
                }
            }
            // Pattern 2: "Seriennummer: XXXXX" oder "Seriennummer XXXXX"
            const serienMatches = body.matchAll(/Seriennummer[:\s]+([A-Z0-9\-]+)/gi);
            for (const match of serienMatches) {
                if (match[1] && match[1].length >= 4 && !serialNumbers.includes(match[1])) {
                    serialNumbers.push(match[1]);
                }
            }
            // Pattern 3: "Serial: XXXXX" oder "Serial Number: XXXXX"
            const serialMatches = body.matchAll(/Serial(?:\s+Number)?[:\s]+([A-Z0-9\-]+)/gi);
            for (const match of serialMatches) {
                if (match[1] && match[1].length >= 4 && !serialNumbers.includes(match[1])) {
                    serialNumbers.push(match[1]);
                }
            }
            console.log('Gefundene Seriennummern:', serialNumbers);

            // 1. Zuerst echte Produktnamen aus E-Mail-Verlauf extrahieren
            // Pattern: "Storage 1: Seagate 4825..." oder "Server 1: Dell PowerEdge..."
            const productMappings = {};
            const mappingPattern = /(Storage|Server|System|Gerät|Device|Unit)\s*(\d+)\s*[:]\s*([^\n\r(]+)/gi;
            let mappingMatch;
            while ((mappingMatch = mappingPattern.exec(body)) !== null) {
                const key = (mappingMatch[1] + ' ' + mappingMatch[2]).toLowerCase().trim();
                let realName = mappingMatch[3].trim();
                // Entferne trailing "(PDF ...)" oder aehnliches
                realName = realName.replace(/\s*\(PDF[^)]*\)\s*$/i, '').trim();
                if (realName.length > 3) {
                    productMappings[key] = realName;
                }
            }
            console.log('Gefundene Produkt-Mappings:', productMappings);

            // 2. Angebot parsen
            let currentProductKey = '';
            let currentProductName = '';

            for (let i = 0; i < lines.length; i++) {
                const line = lines[i].trim();

                // Stopp bei Fußzeile
                if (line.includes('Bitte geben Sie bei weiterer Korrespondenz') ||
                    line.includes('Es gelten die allgemeinen Wartungsvertragsbedingungen')) {
                    break;
                }

                // Pruefen ob es eine Produktzeile ist (z.B. "Storage 1 gem. Anfrage")
                const productLineMatch = line.match(/^(Storage|Server|System|Gerät|Device|Unit)\s*(\d+)\s*(gem\.|gemäß|laut)?/i);
                if (productLineMatch) {
                    currentProductKey = (productLineMatch[1] + ' ' + productLineMatch[2]).toLowerCase().trim();
                    // Versuche echten Namen aus Mapping zu holen
                    currentProductName = productMappings[currentProductKey] || line;
                }

                // Prüfen ob es eine SLA-Zeile ist
                const slaMatch = line.match(/Full Service SLA:\s*([0-9x\/]+(?:NBD)?)\/?-?\s*\(.*?\)\s*([\d.,]+)\s*[€]/i);

                if (slaMatch) {
                    const slaRaw = slaMatch[1].replace(/\//g, 'x').replace(/x-$/, '').replace(/xNBD$/, 'xNBD').trim();
                    const priceStr = slaMatch[2].replace('.', '').replace(',', '.');
                    const price = parseFloat(priceStr);

                    // Wenn kein aktueller Produktname, aus vorherigen Zeilen suchen
                    let productName = currentProductName;
                    if (!productName) {
                        for (let j = i - 1; j >= Math.max(0, i - 5); j--) {
                            const prevLine = lines[j].trim();
                            if (prevLine && !prevLine.includes('Full Service') && !prevLine.includes('€') &&
                                !prevLine.includes('Angebot') && prevLine.length > 2 && prevLine.length < 200) {
                                productName = prevLine;
                                break;
                            }
                        }
                    }

                    if (productName && !isNaN(price)) {
                        // SLA normalisieren
                        let sla = slaRaw;
                        if (slaRaw.includes('NBD') || slaRaw.includes('nbd')) {
                            sla = '5x9xNBD';
                        } else if (slaRaw.match(/7.*24.*4/)) {
                            sla = '7x24x4';
                        } else if (slaRaw.match(/5.*13.*4/)) {
                            sla = '5x13x4';
                        }

                        products.push({
                            name: productName,
                            sla: sla,
                            pricePerMonth: price,
                            reference: reference
                        });
                    }
                }
            }

            return { products, reference, standort, serialNumbers };
        }

        // Hersteller aus Produktname extrahieren
        function extractManufacturer(productName) {
            const manufacturers = {
                'Fujitsu': ['Fujitsu', 'Primergy', 'Eternus', 'PY '],
                'HPE': ['HP ', 'HPE', 'ProLiant', 'Proliant', '3PAR', 'Nimble'],
                'Dell': ['Dell', 'PowerEdge', 'PowerVault', 'EqualLogic', 'Compellent'],
                'NetApp': ['NetApp', 'FAS', 'AFF'],
                'IBM': ['IBM', 'System x', 'Storwize', 'FlashSystem'],
                'Lenovo': ['Lenovo', 'ThinkSystem', 'ThinkServer'],
                'Cisco': ['Cisco', 'UCS'],
                'Oracle': ['Oracle', 'Sun '],
                'Seagate': ['Seagate'],
                'Wortmann': ['Wortmann', 'Terra']
            };

            for (const [mfr, keywords] of Object.entries(manufacturers)) {
                for (const kw of keywords) {
                    if (productName.toLowerCase().includes(kw.toLowerCase())) {
                        return mfr;
                    }
                }
            }
            return 'Sonstiges';
        }

        // Variablen fuer geparste E-Mail-Daten
        let parsedStandort = '';
        let parsedSerialNumbers = [];

        document.getElementById('axians-email-process').addEventListener('click', async () => {
            const file = fileInput.files[0];
            if (!file) {
                alert('Bitte eine .msg Datei auswählen.');
                return;
            }

            try {
                const emailData = await readMsgFile(file);
                const { products, reference, standort, serialNumbers } = parseAxiansEmail(emailData.body, emailData.subject);

                // Geparste Daten speichern
                parsedStandort = standort;
                parsedSerialNumbers = serialNumbers;

                // Info anzeigen
                document.getElementById('axians-email-info').style.display = 'block';
                document.getElementById('axians-email-subject').textContent = 'Betreff: ' + emailData.subject;
                document.getElementById('axians-email-reference').textContent = 'Referenz: ' + (reference || 'Nicht gefunden');
                document.getElementById('axians-email-date').textContent = 'Von: ' + emailData.from;

                // Standort in das Country-Feld eintragen (wenn gefunden)
                if (standort) {
                    document.getElementById('axians-email-country').value = standort;
                }

                // Seriennummern anzeigen (wenn gefunden)
                if (serialNumbers.length > 0) {
                    document.getElementById('axians-email-date').textContent += ' | S/N: ' + serialNumbers.join(', ');
                }

                if (products.length === 0) {
                    alert('Keine Produkte in der E-Mail gefunden. Bitte prüfen Sie, ob es sich um eine Axians-Angebots-E-Mail handelt.');
                    document.getElementById('axians-email-parsed').style.display = 'none';
                    return;
                }

                parsedProducts = products;

                // Gefundene Produkte anzeigen
                const listDiv = document.getElementById('axians-email-products-list');
                listDiv.innerHTML = products.map((p, idx) => `
                    <div style="padding:8px;margin:5px 0;background:#fff;border:1px solid #ddd;border-radius:4px;">
                        <label style="display:flex;align-items:center;gap:10px;">
                            <input type="checkbox" class="axians-email-product-check" data-idx="${idx}" checked>
                            <span><strong>${p.name}</strong> - SLA: ${p.sla} - ${p.pricePerMonth.toFixed(2)} €/Monat</span>
                        </label>
                    </div>
                `).join('');

                document.getElementById('axians-email-parsed').style.display = 'block';
                dropZone.textContent = file.name;

            } catch (error) {
                console.error('Fehler beim Verarbeiten der E-Mail:', error);
                alert('Fehler beim Verarbeiten der E-Mail: ' + error.message);
            }
        });

        document.getElementById('axians-email-add-all').addEventListener('click', () => {
            const checkboxes = document.querySelectorAll('.axians-email-product-check:checked');
            const duration = parseInt(document.getElementById('axians-email-duration').value) || 12;
            const multiplier = parseFloat(document.getElementById('axians-email-multiplier').value) || 1.84;
            const country = document.getElementById('axians-email-country').value.trim() || 'Deutschland';

            checkboxes.forEach(cb => {
                const idx = parseInt(cb.dataset.idx);
                const product = parsedProducts[idx];

                let unitPriceValue = product.pricePerMonth * multiplier * duration;
                const unitPrice = unitPriceValue.toFixed(1);

                let purchaseCost = product.pricePerMonth * duration;
                purchaseCost = purchaseCost.toFixed(2);

                const manufacturer = extractManufacturer(product.name);

                // Duplikat-Check
                const isDuplicate = cart.some(item =>
                    item.name === product.name && item.sla === product.sla &&
                    item.duration === duration && item.country === country
                );

                if (!isDuplicate) {
                    // Description mit Seriennummern erstellen
                    const snText = parsedSerialNumbers.length > 0
                        ? parsedSerialNumbers.join(', ')
                        : '';
                    const description = `S/N: ${snText}\nService Start: tba\nService Ende: tba`;

                    cart.push({
                        name: product.name,
                        active: 1,
                        manufacturer: manufacturer,
                        category: 'Wartung',
                        vendor: 'Axians IT-Infrastructure Services GmbH',
                        unitPrice,
                        qtyInStock: 999,
                        handler: 'Team Wartung',
                        description: description,
                        purchaseCost,
                        sla: product.sla,
                        country,
                        duration,
                        listPrice: 1,
                        pricePerMonth: product.pricePerMonth // Speichern fuer Neuberechnung
                    });
                }
            });

            updateTable();
        });

        function updateTable() {
            const tbody = document.querySelector('#axians-email-table tbody');
            tbody.innerHTML = '';
            cart.forEach((item, index) => {
                const row = document.createElement('tr');
                row.dataset.index = index;
                row.innerHTML = `
                    <td contenteditable="true" class="editable" data-field="name">${item.name}</td>
                    <td data-field="active">${item.active}</td>
                    <td contenteditable="true" class="editable" data-field="manufacturer">${item.manufacturer}</td>
                    <td data-field="category">${item.category}</td>
                    <td data-field="vendor">${item.vendor}</td>
                    <td contenteditable="true" class="editable" data-field="unitPrice">${item.unitPrice}</td>
                    <td data-field="qtyInStock">${item.qtyInStock}</td>
                    <td data-field="handler">${item.handler}</td>
                    <td contenteditable="true" class="editable" data-field="description">${item.description.replace(/\n/g, '<br>')}</td>
                    <td contenteditable="true" class="editable" data-field="purchaseCost">${item.purchaseCost}</td>
                    <td contenteditable="true" class="editable" data-field="sla">${item.sla}</td>
                    <td contenteditable="true" class="editable" data-field="country">${item.country}</td>
                    <td contenteditable="true" class="editable" data-field="duration">${item.duration}</td>
                    <td><button class="imp-btn-danger axians-email-remove" data-index="${index}">X</button></td>
                `;
                tbody.appendChild(row);
            });
            window.axiansEmailCart = cart;

            // Event-Listener fuer Remove-Buttons
            tbody.querySelectorAll('.axians-email-remove').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    const idx = parseInt(e.target.dataset.index);
                    cart.splice(idx, 1);
                    updateTable();
                });
            });
        }

        // Funktion zum Auslesen der aktuellen Tabellenwerte
        function getTableData() {
            const rows = document.querySelectorAll('#axians-email-table tbody tr');
            const data = [];
            rows.forEach(row => {
                const item = {
                    name: row.querySelector('[data-field="name"]')?.textContent || '',
                    active: row.querySelector('[data-field="active"]')?.textContent || '1',
                    manufacturer: row.querySelector('[data-field="manufacturer"]')?.textContent || '',
                    category: row.querySelector('[data-field="category"]')?.textContent || 'Wartung',
                    vendor: row.querySelector('[data-field="vendor"]')?.textContent || '',
                    unitPrice: row.querySelector('[data-field="unitPrice"]')?.textContent || '0',
                    qtyInStock: row.querySelector('[data-field="qtyInStock"]')?.textContent || '999',
                    handler: row.querySelector('[data-field="handler"]')?.textContent || '',
                    description: (row.querySelector('[data-field="description"]')?.innerHTML || '').replace(/<br\s*\/?>/gi, '\n'),
                    purchaseCost: row.querySelector('[data-field="purchaseCost"]')?.textContent || '0',
                    sla: row.querySelector('[data-field="sla"]')?.textContent || '',
                    country: row.querySelector('[data-field="country"]')?.textContent || '',
                    duration: row.querySelector('[data-field="duration"]')?.textContent || '12',
                    listPrice: 1
                };
                data.push(item);
            });
            return data;
        }

        // Preise neu berechnen (ohne Warenkorb zu leeren)
        function recalculatePrices() {
            const duration = parseInt(document.getElementById('axians-email-duration').value) || 12;
            const multiplier = parseFloat(document.getElementById('axians-email-multiplier').value) || 1.84;

            cart.forEach(item => {
                if (item.pricePerMonth) {
                    item.unitPrice = (item.pricePerMonth * multiplier * duration).toFixed(1);
                    item.purchaseCost = (item.pricePerMonth * duration).toFixed(2);
                    item.duration = duration;
                }
            });
            updateTable();
        }

        // Event-Listener fuer Neuberechnung
        document.getElementById('axians-email-duration').addEventListener('change', recalculatePrices);
        document.getElementById('axians-email-multiplier').addEventListener('change', recalculatePrices);

        document.getElementById('axians-email-clear').addEventListener('click', () => {
            cart = [];
            updateTable();
        });

        document.getElementById('axians-email-download').addEventListener('click', () => {
            // Werte aus der Tabelle lesen (inkl. manueller Aenderungen)
            const tableData = getTableData();

            const headers = ["Product Name", "Product Active", "Manufacturer", "Product Category", "Vendor Name", "Unit Price", "Qty. in Stock", "Handler", "Description", "Purchase Cost", "SLA", "Country", "Duration in months", "Listenpreis"];
            const csvRows = [headers.join(';')];
            tableData.forEach(item => {
                csvRows.push([
                    item.name, item.active, item.manufacturer, item.category, item.vendor,
                    item.unitPrice, item.qtyInStock, item.handler,
                    `"${item.description}"`,
                    item.purchaseCost, item.sla, item.country, item.duration, item.listPrice
                ].join(';'));
            });
            downloadCSV(csvRows, 'crm_import_axians_email.csv');
        });
    }

    // ============================================
    // ITRIS EMAIL IMPORTER
    // ============================================
    function initItrisEmail() {
        const panel = document.getElementById('panel-itris-email');
        panel.innerHTML = `
            <h3>ITRIS Email Importer</h3>
            <p style="color:#666;margin-bottom:15px;">Liest Angebote direkt aus ITRIS E-Mails (.msg Dateien)</p>
            <div class="imp-row-grid">
                <div class="imp-form-group">
                    <label>Email (.msg Datei):</label>
                    <input type="file" id="itris-email-file" accept=".msg" class="imp-hidden">
                    <div class="imp-drop-zone" id="itris-email-dropzone">MSG-Datei hierher ziehen oder klicken</div>
                    <button id="itris-email-process">Email verarbeiten</button>
                </div>
                <div>
                    <div class="imp-form-group">
                        <label>SLA auswählen:</label>
                        <select id="itris-email-sla">
                            <option value="7x24x4">7x24x4h</option>
                            <option value="5x9x4">5x9x4h</option>
                            <option value="5x9xNBD" selected>5x9xNBD</option>
                        </select>
                    </div>
                    <div class="imp-form-group">
                        <label>Laufzeit (Monate):</label>
                        <input type="number" id="itris-email-duration" value="12">
                    </div>
                    <div class="imp-form-group">
                        <label>Preis-Multiplikator:</label>
                        <input type="number" id="itris-email-multiplier" value="1.84" step="0.01">
                    </div>
                    <div class="imp-form-group">
                        <label>Land:</label>
                        <input type="text" id="itris-email-country" value="Deutschland" placeholder="z.B. Deutschland, Österreich...">
                    </div>
                </div>
            </div>
            <div id="itris-email-info" style="margin:15px 0;padding:10px;background:#f5f5f5;border-radius:4px;display:none;">
                <strong>Email-Info:</strong>
                <div id="itris-email-subject"></div>
                <div id="itris-email-angebot"></div>
                <div id="itris-email-standort"></div>
            </div>
            <div id="itris-email-parsed" style="margin:15px 0;display:none;">
                <h4>Gefundene Produkte:</h4>
                <div id="itris-email-products-list"></div>
                <button id="itris-email-add-all" style="margin-top:10px;">Ausgewählte zum Warenkorb hinzufügen</button>
            </div>
            <div class="imp-form-group">
                <button id="itris-email-clear" class="imp-btn-danger">Warenkorb leeren</button>
            </div>
            <h4>Warenkorb</h4>
            <div style="overflow-x:auto;">
                <table class="imp-table" id="itris-email-table">
                    <thead>
                        <tr>
                            <th>Product Name</th>
                            <th>Active</th>
                            <th>Manufacturer</th>
                            <th>Category</th>
                            <th>Vendor</th>
                            <th>Unit Price</th>
                            <th>Stock</th>
                            <th>Handler</th>
                            <th>Description</th>
                            <th>Purchase Cost</th>
                            <th>SLA</th>
                            <th>Country</th>
                            <th>Duration</th>
                            <th>Aktion</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
            <button id="itris-email-download">CSV speichern</button>
        `;

        let parsedProducts = [];
        let cart = [];
        let parsedStandort = '';
        let parsedAngebotsnummer = '';

        const fileInput = document.getElementById('itris-email-file');
        const dropZone = document.getElementById('itris-email-dropzone');
        setupDropZone(dropZone, fileInput);

        // ITRIS E-Mail-Text parsen
        function parseItrisEmail(body, subject) {
            const products = [];

            // Sicherstellen dass body ein String ist
            if (typeof body !== 'string') {
                body = body ? String(body) : '';
            }

            // Angebotsnummer extrahieren (z.B. "Angebotsnummer: W-77111904")
            const angebotMatch = body.match(/Angebotsnummer:\s*([A-Z0-9\-]+)/i);
            const angebotsnummer = angebotMatch ? angebotMatch[1].trim() : '';

            // Standort extrahieren (aus der Tabelle: "Standort: Mainz")
            const standortMatch = body.match(/Standort:\s*([^\n\r\t]+)/i);
            let standort = standortMatch ? standortMatch[1].trim() : '';
            // Bereinigen (stoppe bei Tab oder mehreren Leerzeichen)
            standort = standort.split(/\t|\s{2,}/)[0].trim();

            // ITRIS Format: Daten sind auf separaten Zeilen
            // Zeilen aufteilen (behalte \r\n als Trenner)
            const lines = body.split(/\r?\n/).map(l => l.trim()).filter(l => l);

            // Suche nach Seriennummern (alphanumerisch, 6+ Zeichen, allein auf einer Zeile oder am Anfang)
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];

                // Seriennummer-Kandidat (alphanumerisch, 6-12 Zeichen, keine Preise)
                if (/^[A-Z0-9]{6,12}$/.test(line) && !line.includes('€')) {
                    const serialNumber = line;

                    // Suche die 3 Preise nach der Seriennummer
                    const prices = [];
                    for (let j = i + 1; j < Math.min(i + 10, lines.length); j++) {
                        const priceLine = lines[j];
                        const priceMatch = priceLine.match(/^([\d.,]+)\s*€$/);
                        if (priceMatch) {
                            prices.push(parseFloat(priceMatch[1].replace('.', '').replace(',', '.')));
                        }
                        if (prices.length >= 3) break;
                        // Stoppe wenn Gesamt erreicht
                        if (priceLine.startsWith('Gesamt')) break;
                    }

                    // Suche Produktinfo vor der Seriennummer
                    let productName = '';
                    let typ = '';
                    let bezeichnung = '';

                    for (let j = i - 1; j >= Math.max(0, i - 10); j--) {
                        const prevLine = lines[j];

                        // Typ/Modell (z.B. "TS7420 G3")
                        if (/^[A-Z]{2}[A-Z0-9]+(\s+[A-Z0-9]+)?$/.test(prevLine) && !typ) {
                            typ = prevLine;
                        }
                        // Bezeichnung (laenger, enthaelt Beschreibung)
                        if (prevLine.length > 20 && !prevLine.includes('€') && !bezeichnung) {
                            bezeichnung = prevLine.split(/Kaufpreis/i)[0].trim();
                        }
                        // Stoppe bei Header-Zeilen
                        if (prevLine.includes('5x9xNBD') || prevLine.includes('7x24x4h')) break;
                    }

                    if (prices.length >= 3 && (typ || bezeichnung)) {
                        productName = typ ? (bezeichnung ? typ + ' - ' + bezeichnung : typ) : bezeichnung;

                        products.push({
                            name: productName,
                            typ: typ,
                            serialNumber: serialNumber,
                            price7x24x4: prices[0],
                            price5x9x4: prices[1],
                            price5x9xNBD: prices[2]
                        });
                    }
                }
            }

            // Alternative: Suche im gesamten Body nach S/N Pattern gefolgt von Preisen
            if (products.length === 0) {
                // Pattern: S/N gefolgt von 3 Preisen (mit beliebigem Whitespace)
                const regex = /([A-Z0-9]{6,12})\s*\r?\n\s*([\d.,]+)\s*€\s*\r?\n\s*([\d.,]+)\s*€\s*\r?\n\s*([\d.,]+)\s*€/gi;
                let match;
                while ((match = regex.exec(body)) !== null) {
                    const serialNumber = match[1];

                    // Ueberspringe wenn S/N-Kandidat ein bekanntes Keyword ist
                    if (/^(Gesamt|Monate|Monat|Total|EUR)$/i.test(serialNumber)) {
                        continue;
                    }

                    const price1 = parseFloat(match[2].replace('.', '').replace(',', '.'));
                    const price2 = parseFloat(match[3].replace('.', '').replace(',', '.'));
                    const price3 = parseFloat(match[4].replace('.', '').replace(',', '.'));

                    // Finde Produktname vor der S/N
                    const beforeMatch = body.substring(0, match.index);
                    const beforeLines = beforeMatch.split(/\r?\n/).map(l => l.trim()).filter(l => l);

                    let productName = '';
                    let typ = '';
                    for (let j = beforeLines.length - 1; j >= Math.max(0, beforeLines.length - 10); j--) {
                        const prevLine = beforeLines[j];
                        // Stoppe bei Gesamt-Zeilen
                        if (/^Gesamt/i.test(prevLine)) break;

                        if (/^[A-Z]{2}[A-Z0-9]+(\s+[A-Z0-9]+)?$/.test(prevLine) && !typ) {
                            typ = prevLine;
                        }
                        if (prevLine.length > 15 && !prevLine.includes('€') && !productName && !prevLine.match(/^[0-9]+$/)) {
                            productName = prevLine.split(/Kaufpreis/i)[0].trim();
                        }
                        if (typ && productName) break;
                    }

                    products.push({
                        name: typ ? (productName ? typ + ' - ' + productName : typ) : (productName || 'Unbekannt'),
                        typ: typ,
                        serialNumber: serialNumber,
                        price7x24x4: price1,
                        price5x9x4: price2,
                        price5x9xNBD: price3
                    });
                }
            }

            console.log('ITRIS gefundene Produkte:', products);
            return { products, angebotsnummer, standort };
        }

        // Hersteller aus Produktname/Typ extrahieren
        function extractManufacturer(productName) {
            const manufacturers = {
                'TERRA': ['TERRA', 'TGERRA', 'TS7', 'TS5'],
                'Fujitsu': ['Fujitsu', 'Primergy', 'Eternus', 'PY '],
                'HPE': ['HP ', 'HPE', 'ProLiant', 'Proliant', '3PAR', 'Nimble', 'DL3', 'DL5', 'ML3'],
                'Dell': ['Dell', 'PowerEdge', 'PowerVault', 'EqualLogic', 'Compellent', 'R6', 'R7', 'T6'],
                'NetApp': ['NetApp', 'FAS', 'AFF'],
                'IBM': ['IBM', 'System x', 'Storwize', 'FlashSystem'],
                'Lenovo': ['Lenovo', 'ThinkSystem', 'ThinkServer'],
                'Cisco': ['Cisco', 'UCS'],
                'Supermicro': ['Supermicro', 'Super', 'SMC']
            };

            for (const [mfr, keywords] of Object.entries(manufacturers)) {
                for (const kw of keywords) {
                    if (productName.toUpperCase().includes(kw.toUpperCase())) {
                        return mfr;
                    }
                }
            }
            return 'Sonstiges';
        }

        document.getElementById('itris-email-process').addEventListener('click', async () => {
            const file = fileInput.files[0];
            if (!file) {
                alert('Bitte eine .msg Datei auswählen.');
                return;
            }

            try {
                const emailData = await readMsgFile(file);

                // Debug: Zeige Body-Inhalt
                console.log('=== ITRIS EMAIL DEBUG ===');
                console.log('Body Typ:', typeof emailData.body);
                console.log('Body Laenge:', emailData.body ? emailData.body.length : 0);
                console.log('Body Inhalt (erste 2000 Zeichen):', emailData.body ? emailData.body.substring(0, 2000) : 'LEER');
                console.log('Body HTML vorhanden:', !!emailData.bodyHTML);
                if (emailData.bodyHTML) {
                    console.log('Body HTML (erste 2000 Zeichen):', emailData.bodyHTML.substring(0, 2000));
                }
                console.log('=========================');

                const { products, angebotsnummer, standort } = parseItrisEmail(emailData.body, emailData.subject);

                parsedStandort = standort;
                parsedAngebotsnummer = angebotsnummer;

                // Info anzeigen
                document.getElementById('itris-email-info').style.display = 'block';
                document.getElementById('itris-email-subject').textContent = 'Betreff: ' + emailData.subject;
                document.getElementById('itris-email-angebot').textContent = 'Angebotsnummer: ' + (angebotsnummer || 'Nicht gefunden');
                document.getElementById('itris-email-standort').textContent = 'Standort: ' + (standort || 'Nicht gefunden');

                // Standort in Land-Feld eintragen
                if (standort) {
                    document.getElementById('itris-email-country').value = standort;
                }

                if (products.length === 0) {
                    alert('Keine Produkte in der E-Mail gefunden. Bitte prüfen Sie, ob es sich um eine ITRIS-Angebots-E-Mail handelt.');
                    document.getElementById('itris-email-parsed').style.display = 'none';
                    return;
                }

                parsedProducts = products;

                // Gefundene Produkte anzeigen mit allen SLA-Preisen
                const listDiv = document.getElementById('itris-email-products-list');
                listDiv.innerHTML = products.map((p, idx) => `
                    <div style="padding:8px;margin:5px 0;background:#fff;border:1px solid #ddd;border-radius:4px;">
                        <label style="display:flex;align-items:center;gap:10px;">
                            <input type="checkbox" class="itris-email-product-check" data-idx="${idx}" checked>
                            <span>
                                <strong>${p.name}</strong><br>
                                <small>S/N: ${p.serialNumber} | 7x24x4: ${p.price7x24x4.toFixed(2)}€ | 5x9x4: ${p.price5x9x4.toFixed(2)}€ | 5x9xNBD: ${p.price5x9xNBD.toFixed(2)}€</small>
                            </span>
                        </label>
                    </div>
                `).join('');

                document.getElementById('itris-email-parsed').style.display = 'block';
                dropZone.textContent = file.name;

            } catch (error) {
                console.error('Fehler beim Verarbeiten der E-Mail:', error);
                alert('Fehler beim Verarbeiten der E-Mail: ' + error.message);
            }
        });

        document.getElementById('itris-email-add-all').addEventListener('click', () => {
            const checkboxes = document.querySelectorAll('.itris-email-product-check:checked');
            const selectedSla = document.getElementById('itris-email-sla').value;
            const duration = parseInt(document.getElementById('itris-email-duration').value) || 12;
            const multiplier = parseFloat(document.getElementById('itris-email-multiplier').value) || 1.84;
            const country = document.getElementById('itris-email-country').value.trim() || 'Deutschland';

            checkboxes.forEach(cb => {
                const idx = parseInt(cb.dataset.idx);
                const product = parsedProducts[idx];

                // Preis basierend auf ausgewähltem SLA
                let pricePerMonth;
                switch (selectedSla) {
                    case '7x24x4': pricePerMonth = product.price7x24x4; break;
                    case '5x9x4': pricePerMonth = product.price5x9x4; break;
                    case '5x9xNBD': pricePerMonth = product.price5x9xNBD; break;
                    default: pricePerMonth = product.price5x9xNBD;
                }

                const unitPrice = (pricePerMonth * multiplier * duration).toFixed(1);
                const purchaseCost = (pricePerMonth * duration).toFixed(2);
                const manufacturer = extractManufacturer(product.name);

                // Duplikat-Check
                const isDuplicate = cart.some(item =>
                    item.name === product.name && item.sla === selectedSla &&
                    item.duration === duration && item.country === country
                );

                if (!isDuplicate) {
                    const description = `S/N: ${product.serialNumber}\nService Start: tba\nService Ende: tba`;

                    cart.push({
                        name: product.name,
                        active: 1,
                        manufacturer: manufacturer,
                        category: 'Wartung',
                        vendor: 'ITRIS GmbH',
                        unitPrice,
                        qtyInStock: 999,
                        handler: 'Team Wartung',
                        description: description,
                        purchaseCost,
                        sla: selectedSla,
                        country,
                        duration,
                        listPrice: 1,
                        pricePerMonth: pricePerMonth
                    });
                }
            });

            updateTable();
        });

        function updateTable() {
            const tbody = document.querySelector('#itris-email-table tbody');
            tbody.innerHTML = '';
            cart.forEach((item, index) => {
                const row = document.createElement('tr');
                row.dataset.index = index;
                row.innerHTML = `
                    <td contenteditable="true" class="editable" data-field="name">${item.name}</td>
                    <td data-field="active">${item.active}</td>
                    <td contenteditable="true" class="editable" data-field="manufacturer">${item.manufacturer}</td>
                    <td data-field="category">${item.category}</td>
                    <td data-field="vendor">${item.vendor}</td>
                    <td contenteditable="true" class="editable" data-field="unitPrice">${item.unitPrice}</td>
                    <td data-field="qtyInStock">${item.qtyInStock}</td>
                    <td data-field="handler">${item.handler}</td>
                    <td contenteditable="true" class="editable" data-field="description">${item.description.replace(/\n/g, '<br>')}</td>
                    <td contenteditable="true" class="editable" data-field="purchaseCost">${item.purchaseCost}</td>
                    <td contenteditable="true" class="editable" data-field="sla">${item.sla}</td>
                    <td contenteditable="true" class="editable" data-field="country">${item.country}</td>
                    <td contenteditable="true" class="editable" data-field="duration">${item.duration}</td>
                    <td><button class="imp-btn-danger itris-email-remove" data-index="${index}">X</button></td>
                `;
                tbody.appendChild(row);
            });
            window.itrisEmailCart = cart;

            tbody.querySelectorAll('.itris-email-remove').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    const idx = parseInt(e.target.dataset.index);
                    cart.splice(idx, 1);
                    updateTable();
                });
            });
        }

        function getTableData() {
            const rows = document.querySelectorAll('#itris-email-table tbody tr');
            const data = [];
            rows.forEach(row => {
                const item = {
                    name: row.querySelector('[data-field="name"]')?.textContent || '',
                    active: row.querySelector('[data-field="active"]')?.textContent || '1',
                    manufacturer: row.querySelector('[data-field="manufacturer"]')?.textContent || '',
                    category: row.querySelector('[data-field="category"]')?.textContent || 'Wartung',
                    vendor: row.querySelector('[data-field="vendor"]')?.textContent || '',
                    unitPrice: row.querySelector('[data-field="unitPrice"]')?.textContent || '0',
                    qtyInStock: row.querySelector('[data-field="qtyInStock"]')?.textContent || '999',
                    handler: row.querySelector('[data-field="handler"]')?.textContent || '',
                    description: (row.querySelector('[data-field="description"]')?.innerHTML || '').replace(/<br\s*\/?>/gi, '\n'),
                    purchaseCost: row.querySelector('[data-field="purchaseCost"]')?.textContent || '0',
                    sla: row.querySelector('[data-field="sla"]')?.textContent || '',
                    country: row.querySelector('[data-field="country"]')?.textContent || '',
                    duration: row.querySelector('[data-field="duration"]')?.textContent || '12',
                    listPrice: 1
                };
                data.push(item);
            });
            return data;
        }

        function recalculatePrices() {
            const duration = parseInt(document.getElementById('itris-email-duration').value) || 12;
            const multiplier = parseFloat(document.getElementById('itris-email-multiplier').value) || 1.84;

            cart.forEach(item => {
                if (item.pricePerMonth) {
                    item.unitPrice = (item.pricePerMonth * multiplier * duration).toFixed(1);
                    item.purchaseCost = (item.pricePerMonth * duration).toFixed(2);
                    item.duration = duration;
                }
            });
            updateTable();
        }

        document.getElementById('itris-email-duration').addEventListener('change', recalculatePrices);
        document.getElementById('itris-email-multiplier').addEventListener('change', recalculatePrices);

        document.getElementById('itris-email-clear').addEventListener('click', () => {
            cart = [];
            updateTable();
        });

        document.getElementById('itris-email-download').addEventListener('click', () => {
            const tableData = getTableData();
            const headers = ["Product Name", "Product Active", "Manufacturer", "Product Category", "Vendor Name", "Unit Price", "Qty. in Stock", "Handler", "Description", "Purchase Cost", "SLA", "Country", "Duration in months", "Listenpreis"];
            const csvRows = [headers.join(';')];
            tableData.forEach(item => {
                csvRows.push([
                    item.name, item.active, item.manufacturer, item.category, item.vendor,
                    item.unitPrice, item.qtyInStock, item.handler,
                    `"${item.description}"`,
                    item.purchaseCost, item.sla, item.country, item.duration, item.listPrice
                ].join(';'));
            });
            downloadCSV(csvRows, 'crm_import_itris_email.csv');
        });
    }

    // ============================================
    // TECHNOGROUP LIST IMPORTER
    // ============================================
    function initTechnogroup() {
        const panel = document.getElementById('panel-technogroup');
        panel.innerHTML = `
            <h3>Technogroup List Importer</h3>
            <div class="imp-row-grid">
                <div class="imp-form-group">
                    <label>Excel Datei:</label>
                    <input type="file" id="tg-file" accept=".xlsx,.xls" class="imp-hidden">
                    <div class="imp-drop-zone" id="tg-dropzone">Datei hierher ziehen oder klicken</div>
                    <button id="tg-process">Datei verarbeiten</button>
                </div>
                <div class="imp-form-group">
                    <label>Produkt suchen:</label>
                    <input type="text" id="tg-search" placeholder="Suchbegriff...">
                </div>
                <div class="imp-form-group">
                    <label>Hersteller:</label>
                    <select id="tg-manufacturer"><option value="all">Alle Hersteller</option></select>
                </div>
            </div>
            <div class="imp-row-grid">
                <div class="imp-form-group">
                    <label>Produkte:</label>
                    <select id="tg-products" multiple style="height:150px;"></select>
                </div>
                <div>
                    <div class="imp-form-group">
                        <label>SLA:</label>
                        <select id="tg-sla"></select>
                    </div>
                    <div class="imp-form-group">
                        <label>Laufzeit (Monate):</label>
                        <input type="number" id="tg-duration" value="12">
                    </div>
                    <div class="imp-form-group">
                        <label>Preis-Multiplikator:</label>
                        <input type="number" id="tg-multiplier" value="1.84" step="0.01">
                    </div>
                    <div class="imp-form-group">
                        <label>Land:</label>
                        <select id="tg-country">
                            <option value="DE">Deutschland</option>
                            <option value="AT">Österreich</option>
                            <option value="CH">Schweiz</option>
                        </select>
                    </div>
                </div>
            </div>
            <div class="imp-form-group">
                <button id="tg-add">Zum Warenkorb hinzufuegen</button>
                <button id="tg-clear" class="imp-btn-danger">Warenkorb leeren</button>
            </div>
            <h4>Warenkorb</h4>
            <div style="overflow-x:auto;">
                <table class="imp-table" id="tg-table">
                    <thead>
                        <tr>
                            <th>Product Name</th>
                            <th>Active</th>
                            <th>Manufacturer</th>
                            <th>Category</th>
                            <th>Vendor</th>
                            <th>Unit Price</th>
                            <th>Stock</th>
                            <th>Handler</th>
                            <th>Description</th>
                            <th>Purchase Cost</th>
                            <th>SLA</th>
                            <th>Country</th>
                            <th>Duration</th>
                            <th>Aktion</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
            <button id="tg-download">CSV speichern</button>
        `;

        const slaMappings = {
            "5x9x NBD Response": "5x9xNBD",
            "5x9x NBD FIX": "5x9xNBD fix",
            "5x13x 4 h Response": "5x13x4",
            "5x13x 4 h FIX": "5x13x4 fix",
            "5x13x 6 h FIX": "5x13x6 fix",
            "5x13x 8 h FIX": "5x13x8 fix",
            "5x13x 12 h FIX": "5x13x12 fix",
            "5x13x NBD fix": "5x13xNBD fix",
            "7x24x4 h Response": "7x24x4",
            "7x24x4 h FIX": "7x24x4 fix",
            "7x24x6 h FIX": "7x24x6 fix",
            "7x24x8 h FIX": "7x24x8 fix",
            "7x24x12 h FIX": "7x24x12 fix"
        };
        const validTabs = ['HP', 'Dell', 'Fujitsu', 'NetApp', 'IBM', 'EMC', 'Brocade'];

        let productData = [];
        let filteredProductData = [];
        let cart = [];

        const fileInput = document.getElementById('tg-file');
        const dropZone = document.getElementById('tg-dropzone');
        setupDropZone(dropZone, fileInput);

        document.getElementById('tg-process').addEventListener('click', () => {
            const file = fileInput.files[0];
            if (!file) { alert('Bitte eine Datei hochladen.'); return; }

            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                const productSelector = document.getElementById('tg-products');
                const manufacturerFilter = document.getElementById('tg-manufacturer');
                const slaSelector = document.getElementById('tg-sla');

                productSelector.innerHTML = '';
                slaSelector.innerHTML = '';
                manufacturerFilter.innerHTML = '<option value="all">Alle Hersteller</option>';
                productData = [];

                validTabs.forEach(sheetName => {
                    const sheet = workbook.Sheets[sheetName];
                    if (!sheet) return;
                    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                    const slaHeaders = rows[5]?.slice(1, 14) || [];
                    if (slaSelector.options.length === 0) {
                        slaHeaders.forEach((sla, idx) => {
                            const mappedSla = slaMappings[sla] || sla;
                            const option = document.createElement('option');
                            option.value = idx + 1;
                            option.textContent = mappedSla;
                            slaSelector.appendChild(option);
                        });
                    }

                    rows.slice(8).forEach(row => {
                        const productName = (sheetName === 'IBM' || sheetName === 'Brocade') && row[1]
                            ? `${row[0]} - ${row[1]}` : row[0];
                        if (!productName || productName.trim() === '') return;

                        const priceColumns = (sheetName === 'IBM' || sheetName === 'Brocade') ? [2, 7] : [1, 5];
                        const prices = row.slice(...priceColumns).map(p => p || 0);
                        const manufacturer = sheetName === 'HP' ? 'HPE' : sheetName;

                        productData.push({
                            name: productName.trim(),
                            prices,
                            manufacturer,
                            isMissingPrice: prices.includes(0)
                        });

                        if (!Array.from(manufacturerFilter.options).some(o => o.value === manufacturer)) {
                            const option = document.createElement('option');
                            option.value = manufacturer;
                            option.textContent = manufacturer;
                            manufacturerFilter.appendChild(option);
                        }
                    });
                });

                updateProductSelector('');
                dropZone.textContent = file.name;
            };
            reader.readAsArrayBuffer(file);
        });

        function updateProductSelector(searchTerm) {
            const productSelector = document.getElementById('tg-products');
            const selectedManufacturer = document.getElementById('tg-manufacturer').value;
            productSelector.innerHTML = '';

            filteredProductData = productData.filter(p =>
                p.name.toLowerCase().includes(searchTerm) &&
                (selectedManufacturer === 'all' || p.manufacturer === selectedManufacturer)
            );

            filteredProductData.forEach((product, index) => {
                const option = document.createElement('option');
                option.value = index;
                option.textContent = `${product.name} (${product.manufacturer})`;
                if (product.isMissingPrice) option.className = 'missing-price';
                productSelector.appendChild(option);
            });
        }

        document.getElementById('tg-search').addEventListener('input', (e) => {
            updateProductSelector(e.target.value.toLowerCase());
        });
        document.getElementById('tg-manufacturer').addEventListener('change', () => {
            updateProductSelector(document.getElementById('tg-search').value.toLowerCase());
        });

        document.getElementById('tg-add').addEventListener('click', addToCart);
        document.getElementById('tg-products').addEventListener('dblclick', addToCart);

        function addToCart() {
            const productIndices = Array.from(document.getElementById('tg-products').selectedOptions).map(o => o.value);
            const slaIndex = document.getElementById('tg-sla').value;
            const slaText = document.getElementById('tg-sla').selectedOptions[0]?.text || '';
            const duration = document.getElementById('tg-duration').value;
            const country = countryMapping[document.getElementById('tg-country').value];
            const priceMultiplier = parseFloat(document.getElementById('tg-multiplier').value);

            productIndices.forEach(idx => {
                const product = filteredProductData[idx];
                const price = product.prices[slaIndex - 1];
                if (!price || isNaN(price)) return;

                const unitPrice = (parseFloat(price) * priceMultiplier * duration).toFixed(1);
                const purchaseCost = (parseFloat(price) * duration).toFixed(2);

                const isDuplicate = cart.some(item =>
                    item.name === product.name && item.purchaseCost === purchaseCost &&
                    item.duration === duration && item.country === country && item.sla === slaText
                );
                if (isDuplicate) { alert('Produkt bereits im Warenkorb.'); return; }

                cart.push({
                    name: product.name, active: 1, manufacturer: product.manufacturer,
                    category: 'Wartung', vendor: 'Technogroup IT-Service GmbH',
                    unitPrice, qtyInStock: 999, handler: 'Team Wartung',
                    description: 'S/N:\nService Start:\nService Ende:',
                    purchaseCost, sla: slaText, country, duration, listPrice: 1
                });
                updateTable();
            });
        }

        function updateTable() {
            const tbody = document.querySelector('#tg-table tbody');
            tbody.innerHTML = '';
            cart.forEach((item, index) => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td contenteditable="true" class="editable">${item.name}</td>
                    <td>${item.active}</td>
                    <td contenteditable="true" class="editable">${item.manufacturer}</td>
                    <td>${item.category}</td>
                    <td>${item.vendor}</td>
                    <td contenteditable="true" class="editable">${item.unitPrice}</td>
                    <td>${item.qtyInStock}</td>
                    <td>${item.handler}</td>
                    <td contenteditable="true" class="editable">${item.description.replace(/\n/g, '<br>')}</td>
                    <td>${item.purchaseCost}</td>
                    <td>${item.sla}</td>
                    <td>${item.country}</td>
                    <td>${item.duration}</td>
                    <td><button onclick="this.closest('tr').remove(); window.tgCart.splice(${index}, 1);" class="imp-btn-danger">X</button></td>
                `;
                tbody.appendChild(row);
            });
            window.tgCart = cart;
        }

        document.getElementById('tg-clear').addEventListener('click', () => {
            cart = [];
            updateTable();
        });

        document.getElementById('tg-download').addEventListener('click', () => {
            const headers = ["Product Name", "Product Active", "Manufacturer", "Product Category", "Vendor Name", "Unit Price", "Qty. in Stock", "Handler", "Description", "Purchase Cost", "SLA", "Country", "Duration in months", "Listenpreis"];
            const csvRows = [headers.join(';')];
            cart.forEach(item => {
                csvRows.push([
                    item.name, item.active, item.manufacturer, item.category, item.vendor,
                    item.unitPrice, item.qtyInStock, item.handler,
                    `"${item.description}"`,
                    item.purchaseCost, item.sla, item.country, item.duration, item.listPrice
                ].join(';'));
            });
            downloadCSV(csvRows, 'crm_import_technogroup.csv');
        });
    }

    // ============================================
    // TECHNOGROUP PDF IMPORTER
    // ============================================
    function initTechnogroupPDF() {
        const panel = document.getElementById('panel-technogroup-pdf');
        panel.innerHTML = `
            <h3>Technogroup PDF Importer</h3>
            <div class="imp-form-group">
                <input type="file" id="tgpdf-file" accept="application/pdf" class="imp-hidden">
                <div class="imp-drop-zone" id="tgpdf-dropzone">PDF hierher ziehen oder klicken</div>
            </div>
            <div class="imp-row-grid">
                <div class="imp-form-group">
                    <label>Multiplikator:</label>
                    <input type="number" id="tgpdf-multiplier" value="1.84" step="0.01">
                    <button id="tgpdf-update-price">Unit Price aktualisieren</button>
                </div>
                <div class="imp-form-group">
                    <label>Manufacturer:</label>
                    <input type="text" id="tgpdf-manufacturer" placeholder="Hersteller">
                    <button id="tgpdf-apply-manufacturer">Anwenden</button>
                </div>
                <div class="imp-form-group">
                    <label>Land:</label>
                    <input type="text" id="tgpdf-country" value="Deutschland">
                    <button id="tgpdf-apply-country">Anwenden</button>
                </div>
                <div class="imp-form-group">
                    <label>SLA:</label>
                    <input type="text" id="tgpdf-sla" placeholder="Globales SLA">
                    <button id="tgpdf-apply-sla">Anwenden</button>
                </div>
            </div>
            <h4>CSV Vorschau</h4>
            <div style="overflow-x:auto;">
                <table class="imp-table" id="tgpdf-table">
                    <thead>
                        <tr>
                            <th>Product Name</th>
                            <th>Active</th>
                            <th>Manufacturer</th>
                            <th>Category</th>
                            <th>Vendor</th>
                            <th>Unit Price</th>
                            <th>Stock</th>
                            <th>Handler</th>
                            <th>Description</th>
                            <th>Purchase Cost</th>
                            <th>SLA</th>
                            <th>Country</th>
                            <th>Duration</th>
                            <th>Aktion</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
            <div class="imp-form-group" style="margin-top:10px; display:flex; gap:10px;">
                <button id="tgpdf-download">CSV speichern</button>
                <button id="tgpdf-lang-toggle">Sprache: DE → EN</button>
            </div>
        `;

        let globalParsedData = [];
        let tgpdfCurrentLang = 'de';

        const fileInput = document.getElementById('tgpdf-file');
        const dropZone = document.getElementById('tgpdf-dropzone');
        setupDropZone(dropZone, fileInput);

        fileInput.addEventListener('change', async () => {
            const file = fileInput.files[0];
            if (!file || file.type !== 'application/pdf') {
                alert('Bitte eine PDF-Datei auswaehlen.');
                return;
            }
            dropZone.textContent = file.name;
            await processPdf(file);
        });

        async function processPdf(file) {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

            let fullText = '';
            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                const page = await pdf.getPage(pageNum);
                const content = await page.getTextContent();
                const lines = {};
                content.items.forEach(item => {
                    const [,,,, x, y] = item.transform;
                    const yKey = Math.round(y * 10);
                    if (!lines[yKey]) lines[yKey] = [];
                    lines[yKey].push({ x, str: item.str });
                });
                Object.keys(lines).map(k => parseInt(k)).sort((a, b) => b - a).forEach(yKey => {
                    fullText += lines[yKey].sort((a, b) => a.x - b.x).map(i => i.str).join(' ') + '\n';
                });
                fullText += '\n';
            }

            const formatted = formatRawData(fullText.trim());
            globalParsedData = parseFormattedData(formatted);
            generateTable(globalParsedData);
        }

        function formatRawData(rawData) {
            rawData = rawData.replace(/^[\s\S]*?(?=\n?\d+\.\s)/, '');
            rawData = rawData.replace(/^Vertragsnummer:.*$/gm, '');
            rawData = rawData.replace(/Pos Artikelnummer SLA Service Start -\nService Ende\nStueck Einzelpreis \/\nMonat\nGesamtpreis \/\nMonat\nGesamtpreis \/\nLaufzeit/g, '');
            rawData = rawData.replace(/Wartungsangebot/g, '');
            rawData = rawData.replace(/^Technogroup IT-Service GmbH.*$/gm, '');
            rawData = rawData.replace(/^Angebotsnummer:.*$/gm, '');
            rawData = rawData.replace(/^Version:.*$/gm, '');
            rawData = rawData.replace(/^Euro Summe Netto.*$/gm, '');
            rawData = rawData.replace(/\n\s*\n/g, '\n').trim();

            let cleanedData = rawData.trim()
                .replace(/\n\s+/g, '\n')
                .replace(/(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})/g, '$1 - $2')
                .replace(/\n+/g, '\n')
                .replace(/(SN:)\s*\n/g, '$1 n.a.\n')
                .replace(/(Serial Number:)\s*\n/gi, '$1 n.a.\n')
                .replace(/(Seriennummer:)\s*\n/gi, '$1 n.a.\n');

            return cleanedData.replace(/(\d\.\s)(.*?)(?=\d\.\s|$)/gs, '$1$2\n\n').trim();
        }

        function parseFormattedData(formattedData) {
            const blocks = formattedData.trim().split(/\n{2,}/);
            const flatList = [];

            blocks.forEach(blockText => {
                let block = blockText.replace(/^\s*\d+\.\s*/, '').trim();
                const text = block.split('\n').map(l => l.trim()).filter(l => l.length > 0).join(' ');

                const snMatch = text.match(/(?:SN:|Serial Number:|Seriennummer:)\s*([^\s]+)/i);
                const serial = snMatch ? snMatch[1] : 'n.a.';

                const dateMatches = [...text.matchAll(/\d{2}\.\d{2}\.\d{4}/g)].map(m => m[0]);
                let serviceStart = 'N/A', serviceEnde = 'N/A';
                if (dateMatches.length >= 2) {
                    serviceStart = dateMatches[0];
                    serviceEnde = dateMatches[1];
                }

                let sla = 'N/A';
                // SLA mapping - support both orders (CTI NBD and NBD CTI)
                if (/13x5\s*(?:CTI\s*)?NBD(?:\s*CTI)?/i.test(text)) sla = '5x9xNBD';
                else if (/24x7\s*(?:CTI\s*)?NBD(?:\s*CTI)?/i.test(text)) sla = '7x24xNBD';
                else if (/13x5\s*(?:CTI\s*)?(?:4h?|04)(?:\s*CTI)?/i.test(text)) sla = '5x9x4';
                else if (/24x7\s*(?:CTI\s*)?(?:4h?|04)(?:\s*CTI)?/i.test(text)) sla = '7x24x4';
                // Additional patterns for Evernex format
                else if (/5x9\s*NBD/i.test(text)) sla = '5x9xNBD';
                else if (/7x24\s*NBD/i.test(text)) sla = '7x24xNBD';
                else if (/5x9\s*4/i.test(text)) sla = '5x9x4';
                else if (/7x24\s*4/i.test(text)) sla = '7x24x4';

                const priceMatches = [...text.matchAll(/(\d+[.,]\d{2}\s?€)/g)].map(m => m[1]);
                const einzelpreisMonat = priceMatches.length >= 1 ? priceMatches[0] : 'N/A';

                let productName = '';
                const nameEndIndex = text.search(/\d+x\d+/i);
                if (nameEndIndex > 0) {
                    productName = text.substring(0, nameEndIndex).trim();
                } else {
                    const fallbackMatch = text.match(/^(.*?)\s+(?:SN:|\d+[.,]\d{2}\s?€)/);
                    productName = fallbackMatch ? fallbackMatch[1].trim() : text.trim();
                }

                const durationInMonths = calculateDuration(serviceStart, serviceEnde);

                flatList.push({
                    artikelnummer: productName,
                    sla, serviceStart, serviceEnde,
                    einzelpreisMonat, serial, durationInMonths
                });
            });

            // Gruppierung
            const groupedMap = {};
            flatList.forEach(item => {
                const key = `${item.artikelnummer}|${item.sla}|${item.serviceStart}|${item.serviceEnde}|${item.einzelpreisMonat}`;
                if (!groupedMap[key]) {
                    groupedMap[key] = { ...item, seriennummern: [], count: 0 };
                }
                if (item.serial && item.serial !== 'n.a.') {
                    groupedMap[key].seriennummern.push(item.serial);
                }
                groupedMap[key].count++;
            });

            return Object.values(groupedMap);
        }

        function calculateDuration(start, end) {
            if (start === 'N/A' || end === 'N/A') return 12;
            const [d1, m1, y1] = start.split('.');
            const [d2, m2, y2] = end.split('.');
            const dtStart = new Date(`${y1}-${m1}-${d1}`);
            const dtEnd = new Date(`${y2}-${m2}-${d2}`);
            let months = (dtEnd.getFullYear() - dtStart.getFullYear()) * 12 + (dtEnd.getMonth() - dtStart.getMonth());
            if (dtEnd.getDate() >= 15) months++;
            return months > 0 ? months : 1;
        }

        function generateTable(data) {
            const tbody = document.querySelector('#tgpdf-table tbody');
            const multiplier = parseFloat(document.getElementById('tgpdf-multiplier').value) || 1.84;
            const countryInput = document.getElementById('tgpdf-country').value || 'Deutschland';
            const manufacturer = document.getElementById('tgpdf-manufacturer').value || '';
            // Normalisiere das Land
            const country = getCountryForLanguage(normalizeCountry(countryInput), 'de');
            tbody.innerHTML = '';

            data.forEach((item, index) => {
                const unitValue = parseFloat(item.einzelpreisMonat.replace(/\s+/g, '').replace(',', '.').replace('€', '').trim()) || 0;
                const purchaseCost = (unitValue * item.durationInMonths).toFixed(2);
                const unitPrice = (purchaseCost * multiplier).toFixed(2);
                const description = `S/N: ${item.seriennummern.join(', ') || 'n.a.'}\nService Start: ${item.serviceStart}\nService Ende: ${item.serviceEnde}`;

                const row = document.createElement('tr');
                row.innerHTML = `
                    <td contenteditable="true">${item.artikelnummer}</td>
                    <td>1</td>
                    <td>
                        <input type="text" value="${manufacturer}" class="tgpdf-manufacturer-input" style="width:calc(100% - 25px);">
                        <button class="tgpdf-search-btn" title="Manufacturer suchen" style="width:22px;padding:2px;">🔍</button>
                    </td>
                    <td>Wartung</td>
                    <td>Technogroup IT-Service GmbH</td>
                    <td contenteditable="true">${unitPrice}</td>
                    <td>999</td>
                    <td>Team Wartung</td>
                    <td contenteditable="true" style="white-space:pre-wrap;">${description}</td>
                    <td contenteditable="true">${purchaseCost}</td>
                    <td><input type="text" value="${item.sla}" class="tgpdf-sla-input" style="width:100%;"></td>
                    <td><input type="text" value="${country}" class="tgpdf-country-input" style="width:100%;"></td>
                    <td contenteditable="true">${item.durationInMonths}</td>
                    <td><button onclick="this.closest('tr').remove();" class="imp-btn-danger">X</button></td>
                `;
                tbody.appendChild(row);
            });

            // Event-Handler fuer Such-Buttons
            tbody.querySelectorAll('.tgpdf-search-btn').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    const row = e.target.closest('tr');
                    const productName = row.cells[0].textContent.trim();
                    if (productName && productName !== 'N/A') {
                        const searchUrl = `https://www.google.com/search?q=${encodeURIComponent(productName + ' manufacturer')}`;
                        window.open(searchUrl, '_blank');
                    }
                });
            });
        }

        document.getElementById('tgpdf-apply-manufacturer').addEventListener('click', () => {
            const val = document.getElementById('tgpdf-manufacturer').value;
            document.querySelectorAll('.tgpdf-manufacturer-input').forEach(i => i.value = val);
        });
        document.getElementById('tgpdf-apply-country').addEventListener('click', () => {
            const val = document.getElementById('tgpdf-country').value;
            document.querySelectorAll('.tgpdf-country-input').forEach(i => i.value = val);
        });
        document.getElementById('tgpdf-apply-sla').addEventListener('click', () => {
            const val = document.getElementById('tgpdf-sla').value;
            document.querySelectorAll('.tgpdf-sla-input').forEach(i => i.value = val);
        });
        document.getElementById('tgpdf-update-price').addEventListener('click', () => {
            const multiplier = parseFloat(document.getElementById('tgpdf-multiplier').value) || 1.84;
            document.querySelectorAll('#tgpdf-table tbody tr').forEach(row => {
                const purchaseCost = parseFloat(row.cells[9].textContent.replace(',', '.')) || 0;
                row.cells[5].textContent = (purchaseCost * multiplier).toFixed(2);
            });
        });

        document.getElementById('tgpdf-download').addEventListener('click', () => {
            const headers = ["Product Name", "Product Active", "Manufacturer", "Product Category", "Vendor Name", "Unit Price", "Qty. in Stock", "Handler", "Description", "Purchase Cost", "SLA", "Country", "Duration in months"];
            const csvRows = [headers.join(';')];

            document.querySelectorAll('#tgpdf-table tbody tr').forEach(row => {
                const cells = row.cells;
                csvRows.push([
                    cells[0].textContent, cells[1].textContent,
                    cells[2].querySelector('input').value,
                    cells[3].textContent, cells[4].textContent,
                    cells[5].textContent, cells[6].textContent, cells[7].textContent,
                    `"${cells[8].textContent}"`,
                    cells[9].textContent,
                    cells[10].querySelector('input').value,
                    cells[11].querySelector('input').value,
                    cells[12].textContent
                ].join(';'));
            });
            downloadCSV(csvRows, 'vtiger_import_tg_pdf.csv');
        });

        document.getElementById('tgpdf-lang-toggle').addEventListener('click', () => {
            tgpdfCurrentLang = toggleLanguage('tgpdf-table', 'tgpdf-country-input', tgpdfCurrentLang);
            document.getElementById('tgpdf-lang-toggle').textContent =
                tgpdfCurrentLang === 'de' ? 'Sprache: DE → EN' : 'Sprache: EN → DE';
        });
    }

    // ============================================
    // PARKPLACE IMPORTER
    // ============================================
    function initParkplace() {
        const panel = document.getElementById('panel-parkplace');
        panel.innerHTML = `
            <h3>Parkplace Excel Importer</h3>
            <div class="imp-row-grid">
                <div class="imp-form-group">
                    <label>Excel Datei (.xlsx):</label>
                    <input type="file" id="pp-file" accept=".xlsx" class="imp-hidden">
                    <div class="imp-drop-zone" id="pp-dropzone">Datei hierher ziehen oder klicken</div>
                </div>
                <div class="imp-form-group">
                    <label>Multiplikator:</label>
                    <input type="number" id="pp-multiplier" value="1.84" step="0.01">
                    <button id="pp-process">Excel verarbeiten</button>
                </div>
            </div>
            <h4>Ausgabe-Vorschau</h4>
            <div style="overflow-x:auto;">
                <table class="imp-table" id="pp-table">
                    <thead>
                        <tr>
                            <th>Product Name</th>
                            <th>Active</th>
                            <th>Manufacturer</th>
                            <th>Category</th>
                            <th>Vendor</th>
                            <th>Unit Price</th>
                            <th>Stock</th>
                            <th>Handler</th>
                            <th>Description</th>
                            <th>Purchase Cost</th>
                            <th>SLA</th>
                            <th>Country</th>
                            <th>Duration</th>
                            <th>Merged</th>
                            <th>Aktion</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
            <div class="imp-form-group" style="margin-top:10px; display:flex; gap:10px;">
                <button id="pp-download">CSV herunterladen</button>
                <button id="pp-lang-toggle">Sprache: DE → EN</button>
            </div>
        `;

        let ppCurrentLang = 'de';
        const fileInput = document.getElementById('pp-file');
        const dropZone = document.getElementById('pp-dropzone');
        setupDropZone(dropZone, fileInput);

        fileInput.addEventListener('change', () => {
            if (fileInput.files[0]) {
                dropZone.textContent = fileInput.files[0].name;
            }
        });

        document.getElementById('pp-process').addEventListener('click', () => {
            const file = fileInput.files[0];
            if (!file) { alert('Bitte eine Excel-Datei auswaehlen.'); return; }

            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false })
                    .map(row => row.map(cell => cell == null ? '' : cell.toString()));

                const headerIndex = findHeaderIndex(sheetData);
                if (headerIndex < 0) {
                    alert('Keine Tabellenkopfzeile gefunden.');
                    return;
                }

                const headerRow = sheetData[headerIndex].map(c => c.trim().toUpperCase());
                const headerMapping = {};
                headerRow.forEach((val, idx) => headerMapping[val] = idx);

                processDataRows(sheetData.slice(headerIndex + 1), headerMapping);
            };
            reader.readAsArrayBuffer(file);
        });

        function findHeaderIndex(sheetData) {
            for (let i = 0; i < sheetData.length; i++) {
                const row = sheetData[i];
                if (row.length < 2) continue;
                for (let j = 0; j < row.length - 1; j++) {
                    const c0 = (row[j] || '').toString().trim().toUpperCase();
                    const c1 = (row[j + 1] || '').toString().trim().toUpperCase();
                    if ((c0 === 'LINE' || c0 === 'LINIE') && (c1 === 'OEM' || c1 === 'HERSTELLER')) {
                        return i;
                    }
                }
            }
            return -1;
        }

        function formatDate(dateString) {
            if (!dateString) return 'n.a.';
            const str = dateString.toString().trim();
            const dt = new Date(str);
            if (!isNaN(dt)) {
                return `${String(dt.getDate()).padStart(2, '0')}.${String(dt.getMonth() + 1).padStart(2, '0')}.${dt.getFullYear()}`;
            }
            return 'n.a.';
        }

        function calculateDuration(start, end) {
            if (start === 'n.a.' || end === 'n.a.') return 'n.a.';
            const [d1, m1, y1] = start.split('.');
            const [d2, m2, y2] = end.split('.');
            const dtStart = new Date(`${y1}-${m1}-${d1}`);
            const dtEnd = new Date(`${y2}-${m2}-${d2}`);
            let months = (dtEnd.getFullYear() - dtStart.getFullYear()) * 12 + (dtEnd.getMonth() - dtStart.getMonth());
            if (dtEnd.getDate() >= 15) months++;
            return months;
        }

        function processDataRows(rows, header) {
            const mergedRows = [];
            const multiplier = parseFloat(document.getElementById('pp-multiplier').value) || 1.84;

            for (let i = 0; i < rows.length; i++) {
                const cells = rows[i];
                const lineVal = (cells[header['LINE']] || '').trim();
                if (!lineVal) break;

                const purchaseCostText = (cells[header['TOTAL']] || '').trim();
                let rawProductName = (cells[header['PRODUCT DESCRIPTON']] || '').trim()
                    .replace(/-?\s*ParkView Supported/gi, '')
                    .replace(/-?\s*ParkView Support/gi, '')
                    .replace(/•/g, '').trim();

                if (purchaseCostText.toLowerCase() === 'included') {
                    if (rawProductName.toLowerCase().includes('parkview')) continue;
                    const qty = parseInt((cells[header['QTY']] || '1').trim() || '1', 10);
                    const serialNumber = (cells[header['SERIAL NUMBER']] || '').trim();
                    if (mergedRows.length > 0) {
                        const lastRow = mergedRows[mergedRows.length - 1];
                        lastRow.includedItems[rawProductName] = (lastRow.includedItems[rawProductName] || 0) + qty;
                        if (serialNumber) lastRow.serialNumbers.push(serialNumber);
                    }
                } else {
                    const manufacturer = (cells[header['OEM']] || 'N/A').trim() || 'N/A';
                    const serialNumber = (cells[header['SERIAL NUMBER']] || '').trim();
                    const sla = (cells[header['SERVICE LEVEL (SLA)']] || 'N/A').trim() || 'N/A';
                    const location = (cells[header['LOCATION']] || '').trim();
                    const country = location.includes(',') ? location.split(',').pop().trim() : location || 'N/A';
                    const startDate = formatDate((cells[header['START DATE']] || '').trim());
                    const endDate = formatDate((cells[header['END DATE']] || '').trim());

                    let cp = purchaseCostText.replace(/[^0-9,.\-]/g, '').trim();
                    let numericValue = 0;
                    if (cp.match(/^\d{1,3}(\.\d{3})*,\d{2}$/)) {
                        numericValue = parseFloat(cp.replace(/\./g, '').replace(',', '.'));
                    } else if (cp.match(/^\d{1,3}(,\d{3})*\.\d{2}$/)) {
                        numericValue = parseFloat(cp.replace(/,/g, ''));
                    } else if (cp.includes(',')) {
                        numericValue = parseFloat(cp.replace(',', '.'));
                    } else {
                        numericValue = parseFloat(cp);
                    }
                    if (isNaN(numericValue)) numericValue = 0;

                    mergedRows.push({
                        productName: rawProductName || 'N/A',
                        manufacturer,
                        serialNumbers: serialNumber && serialNumber.toUpperCase() !== 'N.A.' ? [serialNumber] : [],
                        sla, country, startDate, endDate,
                        purchaseCost: numericValue.toFixed(2),
                        unitPrice: (numericValue * multiplier).toFixed(2),
                        includedItems: {}, count: 1
                    });
                }
            }

            // Gruppierung
            const finalMap = {};
            mergedRows.forEach(row => {
                const incParts = Object.entries(row.includedItems).map(([n, q]) => `${q}x${n}`).sort().join('|');
                const key = `${row.productName}|${row.manufacturer}|${row.unitPrice}|${row.purchaseCost}|${row.sla}|${row.country}|${row.startDate}|${row.endDate}|${incParts}`;
                if (finalMap[key]) {
                    finalMap[key].serialNumbers = finalMap[key].serialNumbers.concat(row.serialNumbers);
                    finalMap[key].count += row.count;
                } else {
                    finalMap[key] = { ...row, serialNumbers: row.serialNumbers.slice() };
                }
            });

            generateTable(Object.values(finalMap));
        }

        function generateTable(data) {
            const tbody = document.querySelector('#pp-table tbody');
            tbody.innerHTML = '';

            data.forEach(item => {
                const validSerials = item.serialNumbers.filter(sn => sn.toLowerCase() !== 'n.a.' && sn.trim() !== '');
                const descLines = [];
                if (validSerials.length > 0) descLines.push(`S/N: ${validSerials.join(', ')}`);
                const inclNames = Object.keys(item.includedItems);
                if (inclNames.length > 0) {
                    descLines.push('incl.:');
                    inclNames.forEach(name => descLines.push(`${item.includedItems[name]}x ${name}`));
                }
                descLines.push(`Service Start: ${item.startDate}`);
                descLines.push(`Service Ende: ${item.endDate}`);
                const description = descLines.join('\n');
                const duration = calculateDuration(item.startDate, item.endDate);

                const row = document.createElement('tr');
                row.innerHTML = `
                    <td contenteditable="true">${item.productName}</td>
                    <td>1</td>
                    <td contenteditable="true">${item.manufacturer}</td>
                    <td>Wartung</td>
                    <td>Park Place Technologies GmbH</td>
                    <td contenteditable="true">${item.unitPrice}</td>
                    <td>999</td>
                    <td>Team Wartung</td>
                    <td contenteditable="true" style="white-space:pre-wrap;">${description}</td>
                    <td contenteditable="true">${item.purchaseCost}</td>
                    <td contenteditable="true">${item.sla}</td>
                    <td contenteditable="true">${item.country}</td>
                    <td contenteditable="true">${duration}</td>
                    <td>${item.count}</td>
                    <td><button onclick="this.closest('tr').remove();" class="imp-btn-danger">X</button></td>
                `;
                tbody.appendChild(row);
            });
        }

        document.getElementById('pp-download').addEventListener('click', () => {
            const headers = ["Product Name", "Product Active", "Manufacturer", "Product Category", "Vendor Name", "Unit Price", "Qty. in Stock", "Handler", "Description", "Purchase Cost", "SLA", "Country", "Duration in months"];
            const csvRows = [headers.join(';')];

            document.querySelectorAll('#pp-table tbody tr').forEach(row => {
                const cells = row.cells;
                csvRows.push([
                    cells[0].textContent, cells[1].textContent, cells[2].textContent,
                    cells[3].textContent, cells[4].textContent, cells[5].textContent,
                    cells[6].textContent, cells[7].textContent,
                    `"${cells[8].textContent}"`,
                    cells[9].textContent, cells[10].textContent, cells[11].textContent, cells[12].textContent
                ].join(';'));
            });
            downloadCSV(csvRows, 'bereinigte_daten_parkplace.csv');
        });

        document.getElementById('pp-lang-toggle').addEventListener('click', () => {
            ppCurrentLang = toggleLanguage('pp-table', null, ppCurrentLang);
            document.getElementById('pp-lang-toggle').textContent =
                ppCurrentLang === 'de' ? 'Sprache: DE → EN' : 'Sprache: EN → DE';
        });
    }

    // ============================================
    // PARKPLACE PDF IMPORTER
    // ============================================
    function initParkplacePDF() {
        const panel = document.getElementById('panel-parkplace-pdf');
        panel.innerHTML = `
            <h3>Parkplace PDF / E-Mail Importer</h3>
            <div class="imp-form-group">
                <input type="file" id="pppdf-file" accept="application/pdf,.msg" class="imp-hidden">
                <div class="imp-drop-zone" id="pppdf-dropzone">PDF oder MSG-Datei hierher ziehen oder klicken</div>
            </div>
            <div class="imp-row-grid">
                <div class="imp-form-group">
                    <label>Multiplikator:</label>
                    <input type="number" id="pppdf-multiplier" value="1.84" step="0.01">
                    <button id="pppdf-update-price">Unit Price aktualisieren</button>
                </div>
                <div class="imp-form-group">
                    <label>Land ueberschreiben:</label>
                    <input type="text" id="pppdf-country" placeholder="Leer = aus PDF">
                    <button id="pppdf-apply-country">Anwenden</button>
                </div>
                <div class="imp-form-group">
                    <label>SLA ueberschreiben:</label>
                    <input type="text" id="pppdf-sla" placeholder="Leer = aus PDF">
                    <button id="pppdf-apply-sla">Anwenden</button>
                </div>
            </div>
            <h4>CSV Vorschau</h4>
            <div style="overflow-x:auto;">
                <table class="imp-table" id="pppdf-table">
                    <thead>
                        <tr>
                            <th>Product Name</th>
                            <th>Active</th>
                            <th>Manufacturer</th>
                            <th>Category</th>
                            <th>Vendor</th>
                            <th>Unit Price</th>
                            <th>Stock</th>
                            <th>Handler</th>
                            <th>Description</th>
                            <th>Purchase Cost</th>
                            <th>SLA</th>
                            <th>Country</th>
                            <th>Duration</th>
                            <th>Merged</th>
                            <th>Aktion</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
            <div class="imp-form-group" style="margin-top:10px; display:flex; gap:10px;">
                <button id="pppdf-download">CSV herunterladen</button>
                <button id="pppdf-lang-toggle">Sprache: DE → EN</button>
            </div>
        `;

        let pppdfCurrentLang = 'de';
        const fileInput = document.getElementById('pppdf-file');
        const dropZone = document.getElementById('pppdf-dropzone');
        setupDropZone(dropZone, fileInput);

        fileInput.addEventListener('change', async () => {
            const file = fileInput.files[0];
            if (!file) {
                alert('Bitte eine Datei auswaehlen.');
                return;
            }

            dropZone.textContent = file.name;

            // Dateityp erkennen
            const fileName = file.name.toLowerCase();
            if (fileName.endsWith('.msg')) {
                await processParkplaceMsg(file);
            } else if (fileName.endsWith('.pdf') || file.type === 'application/pdf') {
                await processParkplacePdf(file);
            } else {
                alert('Bitte eine PDF- oder MSG-Datei auswaehlen.');
            }
        });

        async function processParkplaceMsg(file) {
            try {
                const emailData = await readMsgFile(file);
                const dataRows = parseParkplaceFromEmail(emailData.body);

                if (dataRows.length === 0) {
                    alert('Keine Parkplace-Daten in der E-Mail gefunden.\\nBitte pruefen Sie das Format.');
                    console.log('E-Mail Body:', emailData.body);
                    return;
                }

                // Verarbeitung wie bei PDF
                const multiplier = parseFloat(document.getElementById('pppdf-multiplier').value) || 1.84;
                const mergedRows = [];

                for (const row of dataRows) {
                    const isIncluded = row.total.toLowerCase() === 'included';

                    if (isIncluded) {
                        if (mergedRows.length > 0) {
                            const lastRow = mergedRows[mergedRows.length - 1];
                            lastRow.includedItems[row.productName] = (lastRow.includedItems[row.productName] || 0) + row.qty;
                            if (row.serial) lastRow.serialNumbers.push(row.serial);
                        }
                    } else {
                        let numericValue = 0;
                        const priceStr = row.total.replace(/[^0-9,.\-]/g, '').trim();
                        if (priceStr.includes(',')) {
                            numericValue = parseFloat(priceStr.replace('.', '').replace(',', '.'));
                        } else {
                            numericValue = parseFloat(priceStr);
                        }
                        if (isNaN(numericValue)) numericValue = 0;

                        let country = 'N/A';
                        if (row.location) {
                            const parts = row.location.split(',');
                            country = parts[parts.length - 1].trim();
                        }

                        mergedRows.push({
                            productName: row.productName,
                            manufacturer: row.oem,
                            serialNumbers: row.serial ? [row.serial] : [],
                            sla: row.sla,
                            country,
                            startDate: row.startDate,
                            endDate: row.endDate,
                            purchaseCost: numericValue.toFixed(2),
                            unitPrice: (numericValue * multiplier).toFixed(2),
                            includedItems: {},
                            count: 1
                        });
                    }
                }

                // Gruppieren
                const finalMap = {};
                mergedRows.forEach(row => {
                    const incParts = Object.entries(row.includedItems).map(([n, q]) => `${q}x${n}`).sort().join('|');
                    const key = `${row.productName}|${row.manufacturer}|${row.unitPrice}|${row.sla}|${row.country}|${row.startDate}|${row.endDate}|${incParts}`;
                    if (finalMap[key]) {
                        finalMap[key].serialNumbers = finalMap[key].serialNumbers.concat(row.serialNumbers);
                        finalMap[key].count += row.count;
                    } else {
                        finalMap[key] = { ...row, serialNumbers: row.serialNumbers.slice() };
                    }
                });

                generateParkplacePdfTable(Object.values(finalMap));
            } catch (error) {
                alert('Fehler beim Lesen der MSG-Datei: ' + error.message);
                console.error(error);
            }
        }

        async function processParkplacePdf(file) {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

            // Text zeilenweise extrahieren mit y-Koordinaten
            let allLines = [];

            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                const page = await pdf.getPage(pageNum);
                const content = await page.getTextContent();
                const items = content.items;

                // Gruppieren nach y-Koordinate
                const lineMap = {};
                items.forEach(item => {
                    const y = Math.round(item.transform[5]);
                    const x = Math.round(item.transform[4]);
                    if (!lineMap[y]) lineMap[y] = [];
                    lineMap[y].push({ x, text: item.str });
                });

                // Zeilen sortieren und zusammenfuegen
                Object.keys(lineMap)
                    .map(Number)
                    .sort((a, b) => b - a)
                    .forEach(y => {
                        const lineItems = lineMap[y].sort((a, b) => a.x - b.x);
                        const lineText = lineItems.map(i => i.text).join(' ').trim();
                        if (lineText) allLines.push(lineText);
                    });
            }

            console.log('Extrahierte Zeilen:', allLines);

            // Bekannte OEMs
            const knownOEMs = ['NetApp', 'Dell', 'HP', 'HPE', 'IBM', 'Cisco', 'EMC', 'Fujitsu', 'Lenovo', 'Sun', 'Oracle', 'Hitachi', 'Pure Storage', 'Nimble'];

            // Datenzeilen parsen
            const dataRows = [];
            const datePattern = /(\d{2}-[A-Za-z]{3}-\d{4})/g;
            const pricePattern = /€([\d.,]+)/;
            const lineNumPattern = /^(\d+\.\d+\.?\d*\.?\d*)\s+/;
            const slaPattern = /(\d+x\d+x\w+)/i;
            const serialPattern = /([A-Z0-9]{8,})/i;

            for (const line of allLines) {
                // Zeile muss mit LINE-Nummer beginnen
                const lineMatch = line.match(lineNumPattern);
                if (!lineMatch) continue;

                // Stopp bei Grand Total
                if (line.includes('Grand Total') || line.includes('Raw Grand Total')) break;

                const lineNum = lineMatch[1];
                const restOfLine = line.substring(lineMatch[0].length);

                // OEM finden
                let oem = 'N/A';
                for (const m of knownOEMs) {
                    if (restOfLine.includes(m)) {
                        oem = m;
                        break;
                    }
                }

                // Preis finden
                let total = '0';
                const priceMatch = line.match(pricePattern);
                if (priceMatch) {
                    total = '€' + priceMatch[1];
                } else if (line.toLowerCase().includes('included')) {
                    total = 'Included';
                }

                // Datum finden
                const dates = line.match(datePattern) || [];
                const startDate = dates[0] || '';
                const endDate = dates[1] || '';

                // SLA finden
                let sla = 'N/A';
                const slaMatch = line.match(slaPattern);
                if (slaMatch) sla = slaMatch[1];

                // Seriennummer finden (nach OEM, vor Datum, 8+ Zeichen)
                let serial = '';
                const afterOemIdx = oem !== 'N/A' ? restOfLine.indexOf(oem) + oem.length : 0;
                const searchArea = restOfLine.substring(afterOemIdx);
                const serialCandidates = searchArea.match(/\b([A-Z0-9]{8,20})\b/gi) || [];
                // Filtere Datumsformate und SLA raus
                for (const c of serialCandidates) {
                    if (!/\d{2}-[A-Za-z]{3}-\d{4}/.test(c) && !/\d+x\d+x/i.test(c)) {
                        serial = c;
                        break;
                    }
                }

                // Location finden (Stadt, Land)
                let location = '';
                const locMatch = line.match(/([A-Za-z\s]+,\s*[A-Za-z]+)\s+\d{2}-/);
                if (locMatch) location = locMatch[1].trim();

                // Produktname extrahieren
                let productName = 'N/A';
                if (oem !== 'N/A') {
                    const oemIdx = restOfLine.indexOf(oem);
                    let afterOem = restOfLine.substring(oemIdx + oem.length).trim();

                    // Entferne bekannte Suffixe und finde Ende
                    const stopPatterns = [
                        /Parts Tech.*$/i,
                        /ParkView.*$/i,
                        /\d+x\d+x\w+/i,
                        /\d{2}-[A-Za-z]{3}-\d{4}/,
                        /[A-Z0-9]{10,}/,
                        /\d+\s+(sepaf|[a-z]+-\d)/i
                    ];

                    for (const pattern of stopPatterns) {
                        const match = afterOem.search(pattern);
                        if (match > 0) {
                            afterOem = afterOem.substring(0, match).trim();
                            break;
                        }
                    }

                    productName = afterOem.replace(/Parts Tech & Labor/gi, '').replace(/ParkView Supported/gi, '').trim() || 'N/A';
                }

                // QTY finden (einzelne Ziffer, oft nach Seriennummer)
                let qty = 1;
                const qtyMatch = line.match(/\s(\d)\s+[a-z]/i);
                if (qtyMatch) qty = parseInt(qtyMatch[1]);

                console.log('Parsed row:', { lineNum, oem, productName, sla, serial, startDate, endDate, total });

                dataRows.push({
                    line: lineNum,
                    oem,
                    productName,
                    sla,
                    serial,
                    qty,
                    location,
                    startDate,
                    endDate,
                    total
                });
            }

            if (dataRows.length === 0) {
                alert('Keine Datenzeilen gefunden. Bitte pruefen Sie das PDF-Format.');
                return;
            }

            // Verarbeiten: Hauptzeilen und Included-Items
            const multiplier = parseFloat(document.getElementById('pppdf-multiplier').value) || 1.84;
            const mergedRows = [];

            for (let i = 0; i < dataRows.length; i++) {
                const row = dataRows[i];
                const isIncluded = row.total.toLowerCase() === 'included';

                if (isIncluded) {
                    // Zu letzter Hauptzeile hinzufuegen
                    if (mergedRows.length > 0) {
                        const lastRow = mergedRows[mergedRows.length - 1];
                        const itemName = row.productName;
                        lastRow.includedItems[itemName] = (lastRow.includedItems[itemName] || 0) + row.qty;
                        if (row.serial) lastRow.serialNumbers.push(row.serial);
                    }
                } else {
                    // Preis parsen
                    let numericValue = 0;
                    const priceStr = row.total.replace(/[^0-9,.\-]/g, '').trim();
                    if (priceStr.match(/^\d{1,3}(\.\d{3})*,\d{2}$/)) {
                        numericValue = parseFloat(priceStr.replace(/\./g, '').replace(',', '.'));
                    } else if (priceStr.match(/^\d{1,3}(,\d{3})*\.\d{2}$/)) {
                        numericValue = parseFloat(priceStr.replace(/,/g, ''));
                    } else if (priceStr.includes(',')) {
                        numericValue = parseFloat(priceStr.replace(',', '.'));
                    } else {
                        numericValue = parseFloat(priceStr);
                    }
                    if (isNaN(numericValue)) numericValue = 0;

                    // Land aus Location
                    let country = 'N/A';
                    if (row.location) {
                        const parts = row.location.split(',');
                        country = parts[parts.length - 1].trim();
                    }

                    mergedRows.push({
                        productName: row.productName,
                        manufacturer: row.oem,
                        serialNumbers: row.serial ? [row.serial] : [],
                        sla: row.sla,
                        country,
                        startDate: row.startDate,
                        endDate: row.endDate,
                        purchaseCost: numericValue.toFixed(2),
                        unitPrice: (numericValue * multiplier).toFixed(2),
                        includedItems: {},
                        count: 1
                    });
                }
            }

            // Gruppieren nach gleichen Eigenschaften
            const finalMap = {};
            mergedRows.forEach(row => {
                const incParts = Object.entries(row.includedItems).map(([n, q]) => `${q}x${n}`).sort().join('|');
                const key = `${row.productName}|${row.manufacturer}|${row.unitPrice}|${row.purchaseCost}|${row.sla}|${row.country}|${row.startDate}|${row.endDate}|${incParts}`;
                if (finalMap[key]) {
                    finalMap[key].serialNumbers = finalMap[key].serialNumbers.concat(row.serialNumbers);
                    finalMap[key].count += row.count;
                    Object.entries(row.includedItems).forEach(([n, q]) => {
                        finalMap[key].includedItems[n] = (finalMap[key].includedItems[n] || 0) + q;
                    });
                } else {
                    finalMap[key] = { ...row, serialNumbers: row.serialNumbers.slice(), includedItems: { ...row.includedItems } };
                }
            });

            generateParkplacePdfTable(Object.values(finalMap));
        }

        function formatDate(dateStr) {
            // DD-MMM-YYYY -> DD.MM.YYYY
            if (!dateStr) return 'n.a.';
            const months = { Jan: '01', Feb: '02', Mar: '03', Apr: '04', May: '05', Jun: '06', Jul: '07', Aug: '08', Sep: '09', Oct: '10', Nov: '11', Dec: '12' };
            const match = dateStr.match(/(\d{2})-([A-Za-z]{3})-(\d{4})/);
            if (match) {
                return `${match[1]}.${months[match[2]] || '01'}.${match[3]}`;
            }
            return dateStr;
        }

        function calculateDuration(start, end) {
            if (!start || !end || start === 'n.a.' || end === 'n.a.') return 12;
            const startFmt = formatDate(start);
            const endFmt = formatDate(end);
            const [d1, m1, y1] = startFmt.split('.');
            const [d2, m2, y2] = endFmt.split('.');
            const dtStart = new Date(`${y1}-${m1}-${d1}`);
            const dtEnd = new Date(`${y2}-${m2}-${d2}`);
            if (isNaN(dtStart) || isNaN(dtEnd)) return 12;
            let months = (dtEnd.getFullYear() - dtStart.getFullYear()) * 12 + (dtEnd.getMonth() - dtStart.getMonth());
            if (dtEnd.getDate() >= 15) months++;
            return months > 0 ? months : 1;
        }

        function generateParkplacePdfTable(data) {
            const tbody = document.querySelector('#pppdf-table tbody');
            tbody.innerHTML = '';

            data.forEach(item => {
                const validSerials = item.serialNumbers.filter(sn => sn && sn.toLowerCase() !== 'n.a.' && sn.trim() !== '');
                const descLines = [];
                if (validSerials.length > 0) descLines.push(`S/N: ${validSerials.join(', ')}`);
                const inclNames = Object.keys(item.includedItems);
                if (inclNames.length > 0) {
                    descLines.push('incl.:');
                    inclNames.forEach(name => descLines.push(`${item.includedItems[name]}x ${name}`));
                }
                descLines.push(`Service Start: ${formatDate(item.startDate)}`);
                descLines.push(`Service Ende: ${formatDate(item.endDate)}`);
                const description = descLines.join('\n');
                const duration = calculateDuration(item.startDate, item.endDate);

                const row = document.createElement('tr');
                row.innerHTML = `
                    <td contenteditable="true">${item.productName}</td>
                    <td>1</td>
                    <td contenteditable="true">${item.manufacturer}</td>
                    <td>Wartung</td>
                    <td>Park Place Technologies GmbH</td>
                    <td contenteditable="true">${item.unitPrice}</td>
                    <td>999</td>
                    <td>Team Wartung</td>
                    <td contenteditable="true" style="white-space:pre-wrap;">${description}</td>
                    <td contenteditable="true">${item.purchaseCost}</td>
                    <td><input type="text" value="${item.sla}" class="pppdf-sla-input" style="width:100%;"></td>
                    <td><input type="text" value="${item.country}" class="pppdf-country-input" style="width:100%;"></td>
                    <td contenteditable="true">${duration}</td>
                    <td>${item.count}</td>
                    <td><button onclick="this.closest('tr').remove();" class="imp-btn-danger">X</button></td>
                `;
                tbody.appendChild(row);
            });
        }

        document.getElementById('pppdf-apply-country').addEventListener('click', () => {
            const val = document.getElementById('pppdf-country').value;
            if (val) document.querySelectorAll('.pppdf-country-input').forEach(i => i.value = val);
        });
        document.getElementById('pppdf-apply-sla').addEventListener('click', () => {
            const val = document.getElementById('pppdf-sla').value;
            if (val) document.querySelectorAll('.pppdf-sla-input').forEach(i => i.value = val);
        });
        document.getElementById('pppdf-update-price').addEventListener('click', () => {
            const multiplier = parseFloat(document.getElementById('pppdf-multiplier').value) || 1.84;
            document.querySelectorAll('#pppdf-table tbody tr').forEach(row => {
                const purchaseCost = parseFloat(row.cells[9].textContent.replace(',', '.')) || 0;
                row.cells[5].textContent = (purchaseCost * multiplier).toFixed(2);
            });
        });

        document.getElementById('pppdf-download').addEventListener('click', () => {
            const headers = ["Product Name", "Product Active", "Manufacturer", "Product Category", "Vendor Name", "Unit Price", "Qty. in Stock", "Handler", "Description", "Purchase Cost", "SLA", "Country", "Duration in months"];
            const csvRows = [headers.join(';')];

            document.querySelectorAll('#pppdf-table tbody tr').forEach(row => {
                const cells = row.cells;
                csvRows.push([
                    cells[0].textContent, cells[1].textContent, cells[2].textContent,
                    cells[3].textContent, cells[4].textContent, cells[5].textContent,
                    cells[6].textContent, cells[7].textContent,
                    `"${cells[8].textContent}"`,
                    cells[9].textContent,
                    cells[10].querySelector('input').value,
                    cells[11].querySelector('input').value,
                    cells[12].textContent
                ].join(';'));
            });
            downloadCSV(csvRows, 'parkplace_pdf_import.csv');
        });

        document.getElementById('pppdf-lang-toggle').addEventListener('click', () => {
            pppdfCurrentLang = toggleLanguage('pppdf-table', 'pppdf-country-input', pppdfCurrentLang);
            document.getElementById('pppdf-lang-toggle').textContent =
                pppdfCurrentLang === 'de' ? 'Sprache: DE → EN' : 'Sprache: EN → DE';
        });
    }

    // ============================================
    // DIS PDF IMPORTER
    // ============================================
    function initDisPDF() {
        const panel = document.getElementById('panel-dis-pdf');
        panel.innerHTML = `
            <h3>DIS PDF Importer</h3>
            <div class="imp-form-group">
                <input type="file" id="dispdf-file" accept="application/pdf" class="imp-hidden">
                <div class="imp-drop-zone" id="dispdf-dropzone">DIS-PDF hierher ziehen oder klicken</div>
            </div>
            <div class="imp-row-grid">
                <div class="imp-form-group">
                    <label>Multiplikator:</label>
                    <input type="number" id="dispdf-multiplier" value="1.84" step="0.01">
                    <button id="dispdf-update-price">Unit Price aktualisieren</button>
                </div>
                <div class="imp-form-group">
                    <label>Manufacturer:</label>
                    <input type="text" id="dispdf-manufacturer" placeholder="Hersteller">
                    <button id="dispdf-apply-manufacturer">Anwenden</button>
                </div>
                <div class="imp-form-group">
                    <label>Land:</label>
                    <input type="text" id="dispdf-country" value="Deutschland">
                    <button id="dispdf-apply-country">Anwenden</button>
                </div>
                <div class="imp-form-group">
                    <label>SLA:</label>
                    <input type="text" id="dispdf-sla" placeholder="Globales SLA">
                    <button id="dispdf-apply-sla">Anwenden</button>
                </div>
            </div>
            <h4>CSV Vorschau</h4>
            <div style="overflow-x:auto;">
                <table class="imp-table" id="dispdf-table">
                    <thead>
                        <tr>
                            <th>Product Name</th>
                            <th>Active</th>
                            <th>Manufacturer</th>
                            <th>Category</th>
                            <th>Vendor</th>
                            <th>Unit Price</th>
                            <th>Stock</th>
                            <th>Handler</th>
                            <th>Description</th>
                            <th>Purchase Cost</th>
                            <th>SLA</th>
                            <th>Country</th>
                            <th>Duration</th>
                            <th>Aktion</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
            <div class="imp-form-group" style="margin-top:10px; display:flex; gap:10px;">
                <button id="dispdf-download">CSV speichern</button>
                <button id="dispdf-lang-toggle">Sprache: DE → EN</button>
            </div>
        `;

        let dispdfCurrentLang = 'de';
        let dispdfParsedData = [];

        const fileInput = document.getElementById('dispdf-file');
        const dropZone = document.getElementById('dispdf-dropzone');
        setupDropZone(dropZone, fileInput);

        fileInput.addEventListener('change', async () => {
            const file = fileInput.files[0];
            if (!file || file.type !== 'application/pdf') {
                alert('Bitte eine PDF-Datei auswaehlen.');
                return;
            }
            dropZone.textContent = file.name;
            await processDisPdf(file);
        });

        async function processDisPdf(file) {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

            let fullText = '';
            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                const page = await pdf.getPage(pageNum);
                const content = await page.getTextContent();
                const lines = {};
                content.items.forEach(item => {
                    const [,,,, x, y] = item.transform;
                    const yKey = Math.round(y * 10);
                    if (!lines[yKey]) lines[yKey] = [];
                    lines[yKey].push({ x, str: item.str });
                });
                Object.keys(lines).map(k => parseInt(k)).sort((a, b) => b - a).forEach(yKey => {
                    fullText += lines[yKey].sort((a, b) => a.x - b.x).map(i => i.str).join(' ') + '\n';
                });
                fullText += '\n';
            }

            dispdfParsedData = parseDisPdf(fullText.trim());
            generateDisTable(dispdfParsedData);
        }

        function parseDisPdf(rawText) {
            const items = [];
            const lines = rawText.split('\n').map(l => l.trim()).filter(l => l);

            // Helper: Parse German price format (170,30 or 1.234,56)
            const parseGermanPrice = (str) => {
                if (!str) return 0;
                let clean = str.replace(/[€EUR\s]/g, '').trim();
                if (clean.includes(',')) {
                    clean = clean.replace(/\./g, '').replace(',', '.');
                }
                return parseFloat(clean) || 0;
            };

            // Extract Laufzeit - Format: "*** Laufzeit: 01.04.2023 - 30.04.2024 ***"
            let serviceStart = 'tba';
            let serviceEnde = 'tba';
            let globalDuration = 12;

            const laufzeitMatch = rawText.match(/\*{3}\s*Laufzeit:\s*(\d{1,2}\.\d{1,2}\.\d{4})\s*-\s*(\d{1,2}\.\d{1,2}\.\d{4})\s*\*{3}/i);
            if (laufzeitMatch) {
                serviceStart = laufzeitMatch[1];
                serviceEnde = laufzeitMatch[2];
                // Berechne Dauer in Monaten
                const [d1, m1, y1] = serviceStart.split('.').map(Number);
                const [d2, m2, y2] = serviceEnde.split('.').map(Number);
                globalDuration = (y2 - y1) * 12 + (m2 - m1) + (d2 >= d1 ? 1 : 0);
                if (globalDuration <= 0) globalDuration = 12;
            } else {
                // Fallback: "Laufzeit: X Monate"
                const durationMatch = rawText.match(/Laufzeit:\s*(\d+)\s*(?:Monate?|Monat)/i);
                if (durationMatch) {
                    globalDuration = parseInt(durationMatch[1], 10);
                }
            }

            // Extract SLA - Format: "Servicezeiten: Mo.- Fr. 08.00-17.00 Uhr (5x9)" + "Reaktionszeit vor Ort: NBD"
            let globalSla = 'tba';
            const serviceZeitenMatch = rawText.match(/Servicezeiten:[^\n]*\((\d+x\d+)\)/i);
            const reaktionszeitMatch = rawText.match(/Reaktionszeit\s*(?:vor\s*Ort)?:\s*(\w+)/i);

            if (serviceZeitenMatch || reaktionszeitMatch) {
                let zeitFenster = serviceZeitenMatch ? serviceZeitenMatch[1].toLowerCase() : '';
                let reaktion = reaktionszeitMatch ? reaktionszeitMatch[1].toLowerCase() : '';

                if (zeitFenster.includes('7x24') || zeitFenster.includes('24x7')) {
                    if (reaktion.includes('4h') || reaktion === '4') globalSla = '7x24x4';
                    else globalSla = '7x24xNBD';
                } else if (zeitFenster.includes('5x9') || zeitFenster.includes('9x5')) {
                    if (reaktion.includes('4h') || reaktion === '4') globalSla = '5x9x4';
                    else globalSla = '5x9xNBD';
                } else if (reaktion) {
                    if (reaktion.includes('nbd')) globalSla = '5x9xNBD';
                    else if (reaktion.includes('4h')) globalSla = '5x9x4';
                }
            }

            // Extract Country - "AT 1200 Wien" oder "DE 85326 München" oder Lieferadresse
            let country = 'Deutschland';
            const countryCodeMatch = rawText.match(/\b(AT|DE|CH)\s+\d{4,5}\s+\w+/i);
            if (countryCodeMatch) {
                const code = countryCodeMatch[1].toUpperCase();
                if (code === 'AT') country = 'Österreich';
                else if (code === 'CH') country = 'Schweiz';
                else country = 'Deutschland';
            }
            // Lieferadresse hat Vorrang
            const lieferMatch = rawText.match(/Lieferadresse[:\s]+[^\n]*\n[^\n]*(AT|DE|CH)\s+\d{4,5}/i);
            if (lieferMatch) {
                const code = lieferMatch[1].toUpperCase();
                if (code === 'AT') country = 'Österreich';
                else if (code === 'CH') country = 'Schweiz';
                else country = 'Deutschland';
            }

            // Known manufacturers for extraction
            const knownManufacturers = ['HP', 'HPE', 'Hewlett Packard', 'Dell', 'Cisco', 'IBM', 'Lenovo',
                'Fujitsu', 'NetApp', 'EMC', 'VMware', 'Microsoft', 'Oracle', 'Juniper', 'Arista',
                'Fortinet', 'Palo Alto', 'CheckPoint', 'F5', 'Citrix', 'Nutanix', 'Pure Storage',
                'Hitachi', 'Huawei', 'Supermicro', 'QNAP', 'Synology', 'APC', 'Eaton', 'Vertiv', 'Brocade'];

            // DIS PDF Format:
            // Zeile: "WL HW-Wartungsgebühr / gesamte Laufzeit 4,00 Stk 170,30 EUR 681,20 €"
            // Nächste Zeile: "Fujitsu Primergy RX2530 M2" (Produktname)
            // Danach: "S/N: YM6B004577, YM6B004578, YM6B004579, YM6B004580"

            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];

                // Erkenne HW-Wartungsgebühr Zeile mit Preisen
                // Format: "WL HW-Wartungsgebühr / gesamte Laufzeit 4,00 Stk 170,30 EUR 681,20 €"
                const wartungMatch = line.match(/(?:HW-Wartung|Wartung|Maintenance).*?(\d+[.,]\d{2})\s*(?:Stk?|St|PC)?\s*(\d+[.,]\d{2})\s*EUR\s+(\d+[.,]\d{2})\s*€/i);

                if (wartungMatch) {
                    const menge = Math.round(parseGermanPrice(wartungMatch[1])) || 1;
                    const einzelpreis = parseGermanPrice(wartungMatch[2]);

                    // Suche Produktname und SNs in den nächsten Zeilen
                    let productName = '';
                    let manufacturer = '';
                    let serials = [];

                    for (let j = i + 1; j < Math.min(i + 8, lines.length); j++) {
                        const nextLine = lines[j];

                        // Neue Wartungszeile? -> Stop
                        if (/HW-Wartung|Wartung|Maintenance/i.test(nextLine) && /\d+[.,]\d{2}\s*€/.test(nextLine)) break;
                        // Summe-Zeile? -> Stop
                        if (/^(Summe|Gesamt|Zwischen|Netto|Übertrag)/i.test(nextLine)) break;

                        // S/N Zeile: "S/N: YM6B004577, YM6B004578, ..."
                        const snMatch = nextLine.match(/^S\/N[:\s]+(.+)/i);
                        if (snMatch) {
                            const snText = snMatch[1].trim();
                            serials = snText.split(/[,;\s]+/).filter(s => s.length >= 4 && /^[A-Za-z0-9_-]+$/.test(s));
                            continue;
                        }

                        // Produktname-Zeile (z.B. "Fujitsu Primergy RX2530 M2")
                        if (!productName && nextLine.length > 3) {
                            // Ignoriere irrelevante Zeilen
                            if (/^(\*{3}|inkl\.|HDD|Artikel|Pos|Menge|ME|EUR|€|\d+[.,]\d{2})/i.test(nextLine)) continue;

                            // Finde Manufacturer
                            let foundMfr = false;
                            for (const mfr of knownManufacturers) {
                                if (nextLine.toLowerCase().startsWith(mfr.toLowerCase() + ' ') ||
                                    nextLine.toLowerCase() === mfr.toLowerCase()) {
                                    manufacturer = mfr;
                                    productName = nextLine.substring(mfr.length).trim() || nextLine;
                                    foundMfr = true;
                                    break;
                                }
                            }

                            if (!foundMfr) {
                                productName = nextLine;
                                // Versuche ersten Begriff als Hersteller
                                const words = nextLine.split(/\s+/);
                                if (words.length >= 2) {
                                    manufacturer = words[0];
                                    productName = words.slice(1).join(' ');
                                }
                            }
                        }
                    }

                    if (productName || serials.length > 0) {
                        items.push({
                            productName: productName || 'N/A',
                            manufacturer,
                            serials: serials.length > 0 ? serials : ['tba'],
                            sla: globalSla,
                            country,
                            duration: globalDuration,
                            purchaseCost: einzelpreis,
                            serviceStart,
                            serviceEnde,
                            menge
                        });
                    }
                    continue;
                }

                // Alternative: Zeile mit EUR und € Preis am Ende (z.B. "... 170,30 EUR 681,20 €")
                const altPriceMatch = line.match(/(\d+[.,]\d{2})\s*EUR\s+(\d+[.,]\d{2})\s*€\s*$/);
                if (altPriceMatch && !items.some(it => it.purchaseCost === parseGermanPrice(altPriceMatch[1]))) {
                    const einzelpreis = parseGermanPrice(altPriceMatch[1]);

                    // Menge aus Zeile
                    const mengeMatch = line.match(/(\d+[.,]\d{2})\s*(?:Stk?|St|PC)/i);
                    const menge = mengeMatch ? Math.round(parseGermanPrice(mengeMatch[1])) : 1;

                    // Suche Produktname und SNs
                    let productName = '';
                    let manufacturer = '';
                    let serials = [];

                    for (let j = i + 1; j < Math.min(i + 8, lines.length); j++) {
                        const nextLine = lines[j];
                        if (/^(Summe|Gesamt|Zwischen|Netto|Übertrag)/i.test(nextLine)) break;
                        if (/\d+[.,]\d{2}\s*EUR\s+\d+[.,]\d{2}\s*€/.test(nextLine)) break;

                        const snMatch = nextLine.match(/^S\/N[:\s]+(.+)/i);
                        if (snMatch) {
                            serials = snMatch[1].split(/[,;\s]+/).filter(s => s.length >= 4 && /^[A-Za-z0-9_-]+$/.test(s));
                            continue;
                        }

                        if (!productName && nextLine.length > 3 && !/^(\*{3}|inkl\.|HDD|Artikel|Pos)/i.test(nextLine)) {
                            for (const mfr of knownManufacturers) {
                                if (nextLine.toLowerCase().startsWith(mfr.toLowerCase() + ' ')) {
                                    manufacturer = mfr;
                                    productName = nextLine.substring(mfr.length).trim();
                                    break;
                                }
                            }
                            if (!productName) {
                                productName = nextLine;
                                const words = nextLine.split(/\s+/);
                                if (words.length >= 2) {
                                    manufacturer = words[0];
                                    productName = words.slice(1).join(' ');
                                }
                            }
                        }
                    }

                    if (productName || serials.length > 0) {
                        items.push({
                            productName: productName || 'N/A',
                            manufacturer,
                            serials: serials.length > 0 ? serials : ['tba'],
                            sla: globalSla,
                            country,
                            duration: globalDuration,
                            purchaseCost: einzelpreis,
                            serviceStart,
                            serviceEnde,
                            menge
                        });
                    }
                }
            }

            // Fallback: S/N-basiertes Parsing
            if (items.length === 0) {
                const snRegex = /S\/N[:\s]+([A-Za-z0-9][A-Za-z0-9_,-\s]*)/gi;
                let allSerials = [];
                let match;
                while ((match = snRegex.exec(rawText)) !== null) {
                    const sns = match[1].split(/[,;\s]+/).filter(s => s.length >= 4 && /^[A-Za-z0-9_-]+$/.test(s));
                    allSerials.push(...sns);
                }

                // Suche Produktname mit bekanntem Hersteller
                let productName = 'N/A';
                let manufacturer = '';
                for (const mfr of knownManufacturers) {
                    const regex = new RegExp(mfr + '\\s+([^\\n]+)', 'i');
                    const pMatch = rawText.match(regex);
                    if (pMatch) {
                        manufacturer = mfr;
                        productName = pMatch[1].trim();
                        break;
                    }
                }

                // Suche Einzelpreis
                let price = 0;
                const priceMatch = rawText.match(/(\d+[.,]\d{2})\s*EUR/);
                if (priceMatch) {
                    price = parseGermanPrice(priceMatch[1]);
                }

                if (allSerials.length > 0 || productName !== 'N/A') {
                    items.push({
                        productName,
                        manufacturer,
                        serials: allSerials.length > 0 ? allSerials : ['tba'],
                        sla: globalSla,
                        country,
                        duration: globalDuration,
                        purchaseCost: price,
                        serviceStart,
                        serviceEnde,
                        menge: allSerials.length || 1
                    });
                }
            }

            return items;
        }

        function generateDisTable(data) {
            const tbody = document.querySelector('#dispdf-table tbody');
            const multiplier = parseFloat(document.getElementById('dispdf-multiplier').value) || 1.84;
            const countryDefault = document.getElementById('dispdf-country').value || 'Deutschland';
            const manufacturerDefault = document.getElementById('dispdf-manufacturer').value || '';
            tbody.innerHTML = '';

            data.forEach((item, index) => {
                // Purchase Cost ist der Einzelpreis
                const purchaseCost = item.purchaseCost.toFixed(2);
                const unitPrice = (item.purchaseCost * multiplier).toFixed(2);
                // Description mit S/N, Service Start und Service Ende
                const serviceStart = item.serviceStart || 'tba';
                const serviceEnde = item.serviceEnde || 'tba';
                const description = `S/N: ${item.serials.join(', ') || 'tba'}\nService Start: ${serviceStart}\nService Ende: ${serviceEnde}`;
                // Verwende extrahierten Manufacturer wenn vorhanden, sonst Default
                const itemManufacturer = item.manufacturer || manufacturerDefault;
                const itemCountry = getCountryForLanguage(item.country || countryDefault, 'de');

                const row = document.createElement('tr');
                row.innerHTML = `
                    <td contenteditable="true">${item.productName}</td>
                    <td>1</td>
                    <td>
                        <input type="text" value="${itemManufacturer}" class="dispdf-manufacturer-input" style="width:calc(100% - 25px);">
                        <button class="dispdf-search-btn" title="Manufacturer suchen" style="width:22px;padding:2px;">🔍</button>
                    </td>
                    <td>Wartung</td>
                    <td>DIS Daten-IT-Service GmbH</td>
                    <td contenteditable="true">${unitPrice}</td>
                    <td>999</td>
                    <td>Team Wartung</td>
                    <td contenteditable="true" style="white-space:pre-wrap;">${description}</td>
                    <td contenteditable="true">${purchaseCost}</td>
                    <td><input type="text" value="${item.sla}" class="dispdf-sla-input" style="width:100%;"></td>
                    <td><input type="text" value="${itemCountry}" class="dispdf-country-input" style="width:100%;"></td>
                    <td contenteditable="true">${item.duration}</td>
                    <td><button onclick="this.closest('tr').remove();" class="imp-btn-danger">X</button></td>
                `;
                tbody.appendChild(row);
            });

            // Event-Handler fuer Such-Buttons
            tbody.querySelectorAll('.dispdf-search-btn').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    const row = e.target.closest('tr');
                    const productName = row.cells[0].textContent.trim();
                    if (productName && productName !== 'N/A') {
                        const searchUrl = `https://www.google.com/search?q=${encodeURIComponent(productName + ' manufacturer')}`;
                        window.open(searchUrl, '_blank');
                    }
                });
            });
        }

        document.getElementById('dispdf-apply-manufacturer').addEventListener('click', () => {
            const val = document.getElementById('dispdf-manufacturer').value;
            document.querySelectorAll('.dispdf-manufacturer-input').forEach(i => i.value = val);
        });
        document.getElementById('dispdf-apply-country').addEventListener('click', () => {
            const val = document.getElementById('dispdf-country').value;
            document.querySelectorAll('.dispdf-country-input').forEach(i => i.value = val);
        });
        document.getElementById('dispdf-apply-sla').addEventListener('click', () => {
            const val = document.getElementById('dispdf-sla').value;
            document.querySelectorAll('.dispdf-sla-input').forEach(i => i.value = val);
        });
        document.getElementById('dispdf-update-price').addEventListener('click', () => {
            const multiplier = parseFloat(document.getElementById('dispdf-multiplier').value) || 1.84;
            document.querySelectorAll('#dispdf-table tbody tr').forEach(row => {
                const purchaseCost = parseFloat(row.cells[9].textContent.replace(',', '.')) || 0;
                row.cells[5].textContent = (purchaseCost * multiplier).toFixed(2);
            });
        });

        document.getElementById('dispdf-download').addEventListener('click', () => {
            const headers = ["Product Name", "Product Active", "Manufacturer", "Product Category", "Vendor Name", "Unit Price", "Qty. in Stock", "Handler", "Description", "Purchase Cost", "SLA", "Country", "Duration in months"];
            const csvRows = [headers.join(';')];

            document.querySelectorAll('#dispdf-table tbody tr').forEach(row => {
                const cells = row.cells;
                csvRows.push([
                    cells[0].textContent, cells[1].textContent,
                    cells[2].querySelector('input').value,
                    cells[3].textContent, cells[4].textContent,
                    cells[5].textContent, cells[6].textContent, cells[7].textContent,
                    `"${cells[8].textContent}"`,
                    cells[9].textContent,
                    cells[10].querySelector('input').value,
                    cells[11].querySelector('input').value,
                    cells[12].textContent
                ].join(';'));
            });
            downloadCSV(csvRows, 'vtiger_import_dis_pdf.csv');
        });

        document.getElementById('dispdf-lang-toggle').addEventListener('click', () => {
            dispdfCurrentLang = toggleLanguage('dispdf-table', 'dispdf-country-input', dispdfCurrentLang);
            document.getElementById('dispdf-lang-toggle').textContent =
                dispdfCurrentLang === 'de' ? 'Sprache: DE → EN' : 'Sprache: EN → DE';
        });
    }

    // ============================================
    // IDS PDF IMPORTER
    // ============================================
    function initIdsPDF() {
        const panel = document.getElementById('panel-ids-pdf');
        panel.innerHTML = `
            <h3>IDS PDF Importer</h3>
            <div class="imp-form-group">
                <input type="file" id="idspdf-file" accept="application/pdf" class="imp-hidden">
                <div class="imp-drop-zone" id="idspdf-dropzone">IDS-PDF hierher ziehen oder klicken</div>
            </div>
            <div class="imp-row-grid">
                <div class="imp-form-group">
                    <label>Multiplikator:</label>
                    <input type="number" id="idspdf-multiplier" value="1.84" step="0.01">
                    <button id="idspdf-update-price">Unit Price aktualisieren</button>
                </div>
                <div class="imp-form-group">
                    <label>Manufacturer:</label>
                    <input type="text" id="idspdf-manufacturer" value="Cisco" placeholder="Hersteller">
                    <button id="idspdf-apply-manufacturer">Anwenden</button>
                </div>
                <div class="imp-form-group">
                    <label>Land:</label>
                    <input type="text" id="idspdf-country" value="">
                    <button id="idspdf-apply-country">Anwenden</button>
                </div>
            </div>
            <h4>CSV Vorschau</h4>
            <div style="overflow-x:auto;">
                <table class="imp-table" id="idspdf-table">
                    <thead>
                        <tr>
                            <th>Product Name</th>
                            <th>Active</th>
                            <th>Manufacturer</th>
                            <th>Category</th>
                            <th>Vendor</th>
                            <th>Unit Price</th>
                            <th>Stock</th>
                            <th>Handler</th>
                            <th>Description</th>
                            <th>Purchase Cost</th>
                            <th>SLA</th>
                            <th>Country</th>
                            <th>Duration</th>
                            <th>Aktion</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
            <div class="imp-form-group" style="margin-top:10px; display:flex; gap:10px;">
                <button id="idspdf-download">CSV speichern</button>
                <button id="idspdf-lang-toggle">Sprache: DE → EN</button>
            </div>
        `;

        let idspdfCurrentLang = 'de';
        let idspdfParsedData = [];

        const fileInput = document.getElementById('idspdf-file');
        const dropZone = document.getElementById('idspdf-dropzone');
        setupDropZone(dropZone, fileInput);

        fileInput.addEventListener('change', async () => {
            const file = fileInput.files[0];
            if (!file || file.type !== 'application/pdf') {
                alert('Bitte eine PDF-Datei auswaehlen.');
                return;
            }
            dropZone.textContent = file.name;
            await processIdsPdf(file);
        });

        async function processIdsPdf(file) {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

            let fullText = '';
            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                const page = await pdf.getPage(pageNum);
                const content = await page.getTextContent();
                const lines = {};
                content.items.forEach(item => {
                    const [,,,, x, y] = item.transform;
                    const yKey = Math.round(y * 10);
                    if (!lines[yKey]) lines[yKey] = [];
                    lines[yKey].push({ x, str: item.str });
                });
                Object.keys(lines).map(k => parseInt(k)).sort((a, b) => b - a).forEach(yKey => {
                    fullText += lines[yKey].sort((a, b) => a.x - b.x).map(i => i.str).join(' ') + '\n';
                });
                fullText += '\n';
            }

            idspdfParsedData = parseIdsPdf(fullText.trim());
            generateIdsTable(idspdfParsedData);
        }

        function parseIdsPdf(rawText) {
            // Extract country from Standort/Endkundenstandort
            let country = '';
            const countryMatch = rawText.match(/(?:Endkundenstandort|Standort):\s*([A-Za-zÄÖÜäöüß ]+)/i);
            if (countryMatch) country = countryMatch[1].trim();

            // Format raw data - remove headers and footers
            let formatted = rawText
                .replace(/\r\n/g, '\n')
                .replace(/^Pos\s+Menge\s+Art[\.-]Nr\s+Text\s+Einzelpreis\s*$/gm, '')
                .replace(/^€\s*$/gm, '')
                .replace(/^Gesamtpreis\s*$/gm, '')
                .replace(/\n{2,}/g, '\n')
                .replace(/-\n/g, '');

            // Remove common footer lines
            const filterPatterns = [
                /^\s*Übertrag/i, /^(Wilhelm-Röntgen|Kunden Nr\.|Debitoren Nr\.|Bearbeiter:)/i,
                /^Bestellnr\./i, /^Lieferdatum:/i, /^Datum:/i, /^Angebot Nr\./i,
                /^Zwischensumme/i, /^Gesamt Netto/i, /^steuerfrei/i, /^Gesamtbetrag/i
            ];

            let lines = formatted.split('\n').map(l => l.trim()).filter(l => l);
            filterPatterns.forEach(pattern => {
                lines = lines.filter(line => !pattern.test(line));
            });

            const items = [];
            let lastIdxProduct = -1;

            for (let idx = 0; idx < lines.length; idx++) {
                const line = lines[idx];

                // Product line: "Pos Menge ... Stck." or "Monat"
                const prodMatch = line.match(/^(?<pos>\d+|A)\s+(?<qty>[\d,]+)\s+(?:Stck\.|Monat)\s+(?<rest>.*)/i);
                if (prodMatch) {
                    const pos = prodMatch.groups.pos;
                    const qty = parseInt(prodMatch.groups.qty.replace(',', '.'), 10);

                    items.push({
                        pos,
                        rawLines: [line],
                        seriennummern: '',
                        seriennummerCount: 0,
                        sla: 'tba',
                        serviceStart: 'tba',
                        serviceEnde: 'tba',
                        stueck: qty,
                        einzelpreis: 0,
                        gesamtpreis: 0,
                        durationInMonths: 12,
                        artikelnummer: 'tba',
                        country
                    });
                    lastIdxProduct = items.length - 1;
                    continue;
                }

                // SN block - sammle alle Seriennummern
                if (/^(?:SN:|S\/N:|Serial:|Seriennummer:)/i.test(line)) {
                    if (lastIdxProduct >= 0) {
                        const serials = [];
                        // Erste Zeile: entferne Praefix und sammle Seriennummern
                        let first = line.replace(/^(?:SN:|S\/N:|Serial:|Seriennummer:)\s*/i, '').trim();
                        // Mehrere Seriennummern koennen komma- oder leerzeichengetrennt sein
                        if (first) {
                            const snParts = first.split(/[,;\s]+/).filter(s => s.length > 3 && /^[A-Za-z0-9_-]+$/.test(s));
                            serials.push(...snParts.length > 0 ? snParts : [first]);
                        }

                        // Folgende Zeilen pruefen - erlaube Bindestriche, Unterstriche und andere SN-Zeichen
                        let k = idx + 1;
                        while (k < lines.length) {
                            const nextLine = lines[k].trim();
                            // Stopp bei neuer Produktzeile oder anderen Sektionen
                            if (/^(?:\d+|A)\s+[\d,]+\s+(?:Stck\.|Monat)/i.test(nextLine)) break;
                            if (/^(?:SN:|S\/N:|Serial:|Seriennummer:|für\s+|Reaktionszeit:|Laufzeit:)/i.test(nextLine)) break;
                            if (/^\d+[.,]\d{2}\s*€?$/.test(nextLine)) break;
                            if (nextLine.length === 0) break;

                            // Seriennummer-Pattern: alphanumerisch mit Bindestrichen/Unterstrichen
                            if (/^[A-Za-z0-9][A-Za-z0-9_-]*$/.test(nextLine) && nextLine.length >= 4) {
                                serials.push(nextLine);
                                k++;
                            } else {
                                break;
                            }
                        }

                        const item = items[lastIdxProduct];
                        // Fuege zu existierenden Seriennummern hinzu (nicht ueberschreiben!)
                        if (item.seriennummern && item.seriennummern !== '') {
                            const existing = item.seriennummern.split(', ').filter(s => s && s !== 'n.a.');
                            serials.push(...existing);
                        }
                        // Duplikate entfernen und speichern
                        const uniqueSerials = [...new Set(serials)];
                        item.seriennummerCount = uniqueSerials.length;
                        item.seriennummern = uniqueSerials.join(', ');
                    }
                    continue;
                }

                // Continuation of product description
                if (lastIdxProduct >= 0) {
                    items[lastIdxProduct].rawLines.push(line);
                }
            }

            // Finalize each item
            items.forEach(item => {
                const blockText = item.rawLines.join(' ');

                // Article number after "für"
                const artnrMatch = blockText.match(/für\s+([A-Z0-9-]+)/i);
                if (artnrMatch) item.artikelnummer = artnrMatch[1].trim();

                // SLA from "Reaktionszeit:"
                const slaMatch = blockText.match(/Reaktionszeit:\s*([\d+xhNBD]+)/i);
                if (slaMatch) {
                    const rawSla = slaMatch[1].trim().toLowerCase();
                    if (rawSla.includes('10x5xnbd') || rawSla.includes('5x9xnbd')) item.sla = '5x9xNBD';
                    else if (rawSla.includes('24x7x4h') || rawSla.includes('7x24x4')) item.sla = '7x24x4';
                    else if (rawSla.includes('24x7xnbd') || rawSla.includes('7x24xnbd')) item.sla = '7x24xNBD';
                    else if (rawSla.includes('5x9x4') || rawSla.includes('10x5x4')) item.sla = '5x9x4';
                }

                // Duration in months
                const durationMatch = blockText.match(/Laufzeit:\s*(\d+)\s*Monate/i);
                if (durationMatch) item.durationInMonths = parseInt(durationMatch[1], 10);

                // Prices - unterstuetze deutsches (1.234,56) und englisches (1,234.56) Format
                const parsePrice = (str) => {
                    if (!str) return 0;
                    let clean = str.replace(/[€\s]/g, '').trim();
                    // Deutsches Format: Punkt als Tausender, Komma als Dezimal
                    if (clean.includes(',') && (clean.indexOf(',') > clean.lastIndexOf('.') || !clean.includes('.'))) {
                        clean = clean.replace(/\./g, '').replace(',', '.');
                    } else {
                        // Englisches Format: Komma als Tausender, Punkt als Dezimal
                        clean = clean.replace(/,/g, '');
                    }
                    return parseFloat(clean) || 0;
                };

                // Suche nach Preisen im Block - deutsches oder englisches Format
                // Deutsches Format: 1.234,56 oder 1234,56
                // Englisches Format: 1,234.56 oder 1234.56
                let priceMatch = blockText.match(/(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s+(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})/);
                if (!priceMatch) {
                    // Alternativ: Preise in Klammern
                    const altMatch = blockText.match(/(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*\(\s*(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*\)/);
                    if (altMatch) priceMatch = altMatch;
                }
                if (!priceMatch) {
                    // Fallback: Einzelner Preis am Ende
                    const singlePrice = blockText.match(/(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*€?\s*$/);
                    if (singlePrice) {
                        item.einzelpreis = parsePrice(singlePrice[1]);
                    }
                }
                if (priceMatch) {
                    item.einzelpreis = parsePrice(priceMatch[1]);
                    item.gesamtpreis = parsePrice(priceMatch[2]);
                }

                if (!item.seriennummerCount) {
                    item.seriennummerCount = item.stueck;
                    item.seriennummern = Array(item.stueck).fill('tba').join(', ');
                }
            });

            return items;
        }

        function generateIdsTable(data) {
            const tbody = document.querySelector('#idspdf-table tbody');
            const multiplier = parseFloat(document.getElementById('idspdf-multiplier').value) || 1.84;
            const countryInput = document.getElementById('idspdf-country').value || '';
            const manufacturer = document.getElementById('idspdf-manufacturer').value || 'Cisco';
            tbody.innerHTML = '';

            data.forEach((item, index) => {
                // Purchase Cost ist der Einzelpreis
                const purchaseCost = item.einzelpreis.toFixed(2);
                const unitPrice = (item.einzelpreis * multiplier).toFixed(2);
                const description = `S/N: ${item.seriennummern || 'tba'}\nService Start: ${item.serviceStart}\nService Ende: ${item.serviceEnde}`;
                // Laender normalisieren
                const itemCountry = getCountryForLanguage(item.country || countryInput, 'de');

                const row = document.createElement('tr');
                row.innerHTML = `
                    <td contenteditable="true">${item.artikelnummer}</td>
                    <td>1</td>
                    <td>
                        <input type="text" value="${manufacturer}" class="idspdf-manufacturer-input" style="width:calc(100% - 25px);">
                        <button class="idspdf-search-btn" title="Manufacturer suchen" style="width:22px;padding:2px;">🔍</button>
                    </td>
                    <td>Wartung</td>
                    <td>Inter Data Systems GmbH</td>
                    <td contenteditable="true">${unitPrice}</td>
                    <td>999</td>
                    <td>Team Wartung</td>
                    <td contenteditable="true" style="white-space:pre-wrap;">${description}</td>
                    <td contenteditable="true">${purchaseCost}</td>
                    <td><input type="text" value="${item.sla}" class="idspdf-sla-input" style="width:100%;"></td>
                    <td><input type="text" value="${itemCountry}" class="idspdf-country-input" style="width:100%;"></td>
                    <td contenteditable="true">${item.durationInMonths}</td>
                    <td><button onclick="this.closest('tr').remove();" class="imp-btn-danger">X</button></td>
                `;
                tbody.appendChild(row);
            });

            // Event-Handler fuer Such-Buttons
            tbody.querySelectorAll('.idspdf-search-btn').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    const row = e.target.closest('tr');
                    const productName = row.cells[0].textContent.trim();
                    if (productName && productName !== 'tba' && productName !== 'N/A') {
                        const searchUrl = `https://www.google.com/search?q=${encodeURIComponent(productName + ' manufacturer')}`;
                        window.open(searchUrl, '_blank');
                    }
                });
            });
        }

        document.getElementById('idspdf-apply-manufacturer').addEventListener('click', () => {
            const val = document.getElementById('idspdf-manufacturer').value;
            document.querySelectorAll('.idspdf-manufacturer-input').forEach(i => i.value = val);
        });
        document.getElementById('idspdf-apply-country').addEventListener('click', () => {
            const val = document.getElementById('idspdf-country').value;
            document.querySelectorAll('.idspdf-country-input').forEach(i => i.value = val);
        });
        document.getElementById('idspdf-update-price').addEventListener('click', () => {
            const multiplier = parseFloat(document.getElementById('idspdf-multiplier').value) || 1.84;
            document.querySelectorAll('#idspdf-table tbody tr').forEach(row => {
                const purchaseCost = parseFloat(row.cells[9].textContent.replace(',', '.')) || 0;
                row.cells[5].textContent = (purchaseCost * multiplier).toFixed(2);
            });
        });

        document.getElementById('idspdf-download').addEventListener('click', () => {
            const headers = ["Product Name", "Product Active", "Manufacturer", "Product Category", "Vendor Name", "Unit Price", "Qty. in Stock", "Handler", "Description", "Purchase Cost", "SLA", "Country", "Duration in months"];
            const csvRows = [headers.join(';')];

            document.querySelectorAll('#idspdf-table tbody tr').forEach(row => {
                const cells = row.cells;
                csvRows.push([
                    cells[0].textContent, cells[1].textContent,
                    cells[2].querySelector('input').value,
                    cells[3].textContent, cells[4].textContent,
                    cells[5].textContent, cells[6].textContent, cells[7].textContent,
                    `"${cells[8].textContent}"`,
                    cells[9].textContent,
                    cells[10].querySelector('input').value,
                    cells[11].querySelector('input').value,
                    cells[12].textContent
                ].join(';'));
            });
            downloadCSV(csvRows, 'vtiger_import_ids_pdf.csv');
        });

        document.getElementById('idspdf-lang-toggle').addEventListener('click', () => {
            idspdfCurrentLang = toggleLanguage('idspdf-table', 'idspdf-country-input', idspdfCurrentLang);
            document.getElementById('idspdf-lang-toggle').textContent =
                idspdfCurrentLang === 'de' ? 'Sprache: DE → EN' : 'Sprache: EN → DE';
        });
    }

    // ============================================
    // INITIALISIERUNG
    // ============================================
    function init() {
        addFloatingButton();
        initAxians();
        initAxiansEmail();
        initItrisEmail();
        initTechnogroup();
        initTechnogroupPDF();
        initParkplace();
        initParkplacePDF();
        initDisPDF();
        initIdsPDF();
    }

    // Warten bis DOM vollstaendig geladen
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
