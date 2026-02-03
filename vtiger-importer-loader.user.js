// ==UserScript==
// @name         VTiger Products Importer (Loader)
// @namespace    https://vtiger.hardwarewartung.com
// @version      1.2.2
// @description  Laedt den VTiger Products Importer automatisch von GitHub (inkl. MSG-Support)
// @author       Hardwarewartung
// @match        https://vtiger.hardwarewartung.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_addElement
// @connect      raw.githubusercontent.com
// @connect      cdnjs.cloudflare.com
// @connect      esm.sh
// @connect      cdn.jsdelivr.net
// @connect      unpkg.com
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

    // ES-Modul laden und als globale Variable exportieren
    function loadESModule(url, globalName) {
        return new Promise((resolve, reject) => {
            const eventName = globalName + 'Loaded';
            const timeoutId = setTimeout(() => {
                reject(new Error('Timeout beim Laden von ' + globalName));
            }, 10000);

            const script = document.createElement('script');
            script.type = 'module';
            script.textContent =
                'import Module from "' + url + '";' +
                'window.' + globalName + ' = Module.default || Module;' +
                'window.dispatchEvent(new CustomEvent("' + eventName + '"));';

            window.addEventListener(eventName, () => {
                clearTimeout(timeoutId);
                resolve();
            }, { once: true });

            script.onerror = (e) => {
                clearTimeout(timeoutId);
                reject(e);
            };
            document.head.appendChild(script);
        });
    }

    const LOADER_VERSION = '1.2.2';

    // Hauptfunktion
    async function init() {
        try {
            console.log('[VTiger Importer Loader] Version ' + LOADER_VERSION);
            console.log('[VTiger Importer] Lade Bibliotheken...');

            // 1. Externe Bibliotheken laden
            await injectScript(XLSX_URL);
            console.log('[VTiger Importer] XLSX geladen');

            await injectScript(PDFJS_URL);
            console.log('[VTiger Importer] PDF.js geladen');

            // 2. MsgReader laden - versuche verschiedene Quellen
            let msgReaderLoaded = false;

            // Versuch 1: jsdelivr mit UMD build
            try {
                await injectScript('https://cdn.jsdelivr.net/npm/@poplor/msgreader@3.2.0/dist/MsgReader.umd.min.js');
                if (window.MsgReader) {
                    msgReaderLoaded = true;
                    console.log('[VTiger Importer] MsgReader v3.2.0 (jsdelivr) geladen');
                }
            } catch (e) {
                console.log('[VTiger Importer] jsdelivr fehlgeschlagen, versuche Alternative...');
            }

            // Versuch 2: unpkg mit der alten Version
            if (!msgReaderLoaded) {
                try {
                    await injectScript('https://unpkg.com/msgreader@1.0.1/dist/MsgReader.js');
                    if (window.MsgReader) {
                        msgReaderLoaded = true;
                        console.log('[VTiger Importer] MsgReader v1.0.1 (unpkg) geladen');
                    }
                } catch (e) {
                    console.log('[VTiger Importer] unpkg fehlgeschlagen');
                }
            }

            // Versuch 3: esm.sh als Fallback
            if (!msgReaderLoaded) {
                await loadESModule('https://esm.sh/msgreader@1.0.1', 'MsgReader');
                console.log('[VTiger Importer] MsgReader v1.0.1 (esm.sh) geladen');
            }

            // 3. Hauptscript von GitHub laden
            const mainScript = await loadScript(SCRIPT_URL);

            // 4. Die Userscript-Header entfernen und Code extrahieren
            const codeStart = mainScript.indexOf('(function()');
            if (codeStart === -1) {
                throw new Error('Could not find script code');
            }
            const cleanCode = mainScript.substring(codeStart);

            // 5. Script ausfuehren
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
