// ==UserScript==
// @name         VTiger Products Importer (Loader)
// @namespace    https://vtiger.hardwarewartung.com
// @version      1.0.0
// @description  Laedt den VTiger Products Importer automatisch von GitHub
// @author       Hardwarewartung
// @match        https://vtiger.hardwarewartung.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_addElement
// @connect      raw.githubusercontent.com
// @connect      cdnjs.cloudflare.com
// @updateURL    https://raw.githubusercontent.com/HWW24-Office/products-importer/main/vtiger-importer-loader.user.js
// @downloadURL  https://raw.githubusercontent.com/HWW24-Office/products-importer/main/vtiger-importer-loader.user.js
// ==/UserScript==

(function() {
    'use strict';

    const SCRIPT_URL = 'https://raw.githubusercontent.com/HWW24-Office/products-importer/main/vtiger-importer.user.js';
    const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js';
    const PDFJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.14.305/pdf.min.js';

    // Hilfsfunktion zum Laden von Scripts
    function loadScript(url) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: 'GET',
                url: url,
                onload: function(response) {
                    if (response.status === 200) {
                        resolve(response.responseText);
                    } else {
                        reject(new Error(`Failed to load ${url}: ${response.status}`));
                    }
                },
                onerror: function(error) {
                    reject(error);
                }
            });
        });
    }

    // Externes Script in die Seite einfuegen
    function injectScript(url) {
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = url;
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

    // Inline-Script einfuegen
    function injectInlineScript(code) {
        const script = document.createElement('script');
        script.textContent = code;
        document.head.appendChild(script);
    }

    // Hauptfunktion
    async function init() {
        try {
            // 1. Externe Bibliotheken laden
            await injectScript(XLSX_URL);
            await injectScript(PDFJS_URL);

            // 2. Hauptscript von GitHub laden
            const mainScript = await loadScript(SCRIPT_URL);

            // 3. Die Userscript-Header entfernen und Code extrahieren
            const codeStart = mainScript.indexOf('(function()');
            if (codeStart === -1) {
                throw new Error('Could not find script code');
            }
            const cleanCode = mainScript.substring(codeStart);

            // 4. Script ausfuehren
            injectInlineScript(cleanCode);

            console.log('[VTiger Importer] Erfolgreich geladen von GitHub');
        } catch (error) {
            console.error('[VTiger Importer] Fehler beim Laden:', error);
            alert('VTiger Importer konnte nicht geladen werden. Siehe Konsole fuer Details.');
        }
    }

    // Starten wenn DOM bereit
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
