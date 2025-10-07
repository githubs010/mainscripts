(async function() {
    // --- CONFIGURATION ---
    const SHEET_URL = "https://opensheet.elk.sh/188552daH24yAiXUux5aHvqBNWOPRZPJeve2Nd6acRBA/Sheet1";
    const FALLBACK_ADMIN = 'prasad'; // A default user if the sheet fails to load

    // --- 🔑 START: DYNAMIC ACCESS CONTROL ---
    async function getAuthorizedUsers(sheetUrl) {
        try {
            const usersSheetUrl = sheetUrl.replace('/Sheet1', '/Users');
            const response = await fetch(usersSheetUrl);
            if (!response.ok) {
                console.error("Failed to fetch Users sheet, using fallback.");
                return [FALLBACK_ADMIN];
            }
            const users = await response.json();
            return users.map(user => user.username.toLowerCase()).filter(Boolean);
        } catch (e) {
            console.error("Error fetching users, using fallback.", e);
            return [FALLBACK_ADMIN];
        }
    }

    const AUTHORIZED_USERS = await getAuthorizedUsers(SHEET_URL);
    const currentUser = window.WoflowAccessUser;

    if (!currentUser || !AUTHORIZED_USERS.includes(currentUser.toLowerCase())) {
        alert('⛔ Access Denied. Please contact the administrator.');
        return;
    }
    // --- END: ACCESS CONTROL ---


    // --- OPTIMIZATION: Constants ---
    const TYPO_CONFIG = {
        libURL: 'https://cdn.jsdelivr.net/npm/typo-js@1.2.1/typo.js',
        dictionaries: [{
            name: 'en_US',
            affURL: 'https://cdn.jsdelivr.net/npm/dictionary-en-us@2.2.0/index.aff',
            dicURL: 'https://cdn.jsdelivr.net/npm/dictionary-en-us@2.2.0/index.dic'
        }],
        ignoreLength: 3
    };
    const MONITORED_DIV_PREFIXES = [
        "Secondary UPC :", "Mx Provided Category 2 :", "Mx Provided Category 1 :", "Mx Provided Category 3 :",
        "Original Brand Name :", "Mx Provided Product Description :", "Original Item Name :", "Mx Provided Descriptor(s) :",
        "Mx Provided Size 2 :", "Original UOM :", "Original Size :", "Mx Provided CBD/THC Content :", "Photo Source :",
        "itemName :", "Mx Provided WI Flag :", "WI Type :", "L1 Name :", "Woflow Notes :", "Exclude :", "Invalid Reason :",
        "upc :", "itemMerchantSuppliedId :"
    ];
    const LEVENSHTEIN_TYPO_THRESHOLD = 3;
    const FAST_DELAY_MS = 50;
    const INTERACTION_DELAY_MS = 100;
    const SEARCH_DELAY_MS = 400;

    const SELECTORS = {
        cleanedItemName: 'textarea[name="Woflow Cleaned Item Name"]',
        brandPath: 'input[name="Woflow brand_path"]',
        searchBox: 'input[name="search-box"]',
        searchResults: 'a.search-results',
        dropdownOption: '.vs__dropdown-option, .vs__dropdown-menu li'
    };

    // --- Global State and Caching ---
    const state = {
        isUpdatingComparison: false,
        isScriptEnabled: true,
        matchedSheetRow: null,
    };
    const dictionaries = [];
    const domCache = {
        cleanedItemNameTextarea: null,
        searchBoxInput: null,
        woflowBrandPathInput: null,
        allDivs: [] // Will be populated later
    };

    // --- Utility Functions ---
    const normalizeText = (text) => text?.replace(/&amp;/g, "&").replace(/\s+/g, " ").trim().toLowerCase();
    const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));
    const regexEscape = (str) => str.replace(/[-\/\^$*+?.()|[\]{}]/g, '\$&');
    const escapeHtml = (unsafe) => unsafe.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "'");

    if (window.__autoFillObserver) {
        window.__autoFillObserver.disconnect();
        delete window.__autoFillObserver;
    }

    function findDivByTextPrefix(prefix) {
        return domCache.allDivs.find(e => e.textContent.trim().startsWith(prefix)) || null;
    }

    function updateTextarea(textarea, value) {
        if (textarea) {
            textarea.value = value;
            textarea.dispatchEvent(new Event('input', { bubbles: true, passive: true }));
            textarea.dispatchEvent(new Event('change', { bubbles: true, passive: true }));
        }
    }

    function toTitleCase(str) {
        if (!str) return '';
        return str.toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
    }

    function levenshtein(s1, s2) {
        s1 = s1.toLowerCase();
        s2 = s2.toLowerCase();
        const costs = Array(s2.length + 1).fill(0).map((_, i) => i);
        for (let i = 1; i <= s1.length; i++) {
            let lastValue = i;
            for (let j = 1; j <= s2.length; j++) {
                const newValue = costs[j - 1] + (s1.charAt(i - 1) !== s2.charAt(j - 1) ? 1 : 0);
                costs[j - 1] = lastValue;
                lastValue = Math.min(costs[j] + 1, newValue, lastValue + 1);
            }
            costs[s2.length] = lastValue;
        }
        return costs[s2.length];
    }

    async function loadScript(url) {
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = url;
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

    // --- NEW: UI Module ---
    const UI = {
        showToast: (message, type = 'info') => {
            const toast = document.createElement('div');
            toast.className = `script-toast toast-${type}`;
            toast.textContent = message;
            document.body.appendChild(toast);
            setTimeout(() => toast.classList.add('show'), 10);
            setTimeout(() => {
                toast.classList.remove('show');
                setTimeout(() => toast.remove(), 500);
            }, 3000);
        },
        updateStatus: (text) => {
            const statusDiv = document.getElementById('script-status');
            if (statusDiv) statusDiv.textContent = text;
        },
        injectControls: () => {
            const css = `
                .script-control-panel { position: fixed; bottom: 15px; right: 15px; background: #fff; border: 1px solid #ddd; border-radius: 8px; padding: 15px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); z-index: 9999; font-family: sans-serif; font-size: 14px; }
                .script-control-panel h3 { margin: 0 0 10px; font-size: 16px; color: #333; }
                .script-control-panel button { background-color: #007bff; color: white; border: none; padding: 8px 12px; border-radius: 4px; cursor: pointer; margin-top: 5px; width: 100%; transition: background-color 0.2s; }
                .script-control-panel button:hover { background-color: #0056b3; }
                .script-control-panel .toggle-switch { display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; }
                .script-control-panel #script-status { margin-top: 10px; font-size: 12px; color: #666; text-align: center; }
                .switch { position: relative; display: inline-block; width: 40px; height: 20px; }
                .switch input { opacity: 0; width: 0; height: 0; }
                .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #ccc; transition: .4s; border-radius: 20px; }
                .slider:before { position: absolute; content: ""; height: 12px; width: 12px; left: 4px; bottom: 4px; background-color: white; transition: .4s; border-radius: 50%; }
                input:checked + .slider { background-color: #28a745; }
                input:checked + .slider:before { transform: translateX(20px); }
                .script-toast { position: fixed; top: 20px; right: 20px; background-color: #333; color: white; padding: 15px 20px; border-radius: 5px; z-index: 10000; opacity: 0; transition: opacity 0.5s, top 0.5s; font-family: sans-serif; }
                .script-toast.show { opacity: 1; top: 30px; }
                .toast-error { background-color: #dc3545; }
            `;
            const style = document.createElement('style');
            style.textContent = css;
            document.head.appendChild(style);

            const panel = document.createElement('div');
            panel.className = 'script-control-panel';
            panel.innerHTML = `
                <h3>Automation Controls</h3>
                <div class="toggle-switch">
                    <label for="enable-script-toggle">Enable Automation</label>
                    <label class="switch">
                        <input type="checkbox" id="enable-script-toggle" ${state.isScriptEnabled ? 'checked' : ''}>
                        <span class="slider"></span>
                    </label>
                </div>
                <button id="clean-fill-btn">Clean & Fill Name</button>
                <button id="run-autofill-btn">Run Autofill</button>
                <div id="script-status">Initializing...</div>
            `;
            document.body.appendChild(panel);

            document.getElementById('enable-script-toggle').addEventListener('change', (e) => {
                state.isScriptEnabled = e.target.checked;
                UI.showToast(`Automation ${state.isScriptEnabled ? 'Enabled' : 'Disabled'}`);
                if (state.isScriptEnabled) runAllComparisons();
            });
            document.getElementById('run-autofill-btn').addEventListener('click', runFullAutomation);
            document.getElementById('clean-fill-btn').addEventListener('click', cleanAndFillName);
        }
    };

    async function loadTypoLibrary() {
        try {
            if (typeof Typo === 'undefined') await loadScript(TYPO_CONFIG.libURL);
            if (dictionaries.length > 0) return;
            const dictPromises = TYPO_CONFIG.dictionaries.map(async dictConfig => {
                const [affResponse, dicResponse] = await Promise.all([fetch(dictConfig.affURL), fetch(dictConfig.dicURL)]);
                return new Typo(dictConfig.name, await affResponse.text(), await dicResponse.text());
            });
            dictionaries.push(...(await Promise.all(dictPromises)));
        } catch (error) {
            console.error("Could not load Typo library.", error);
            UI.showToast('Error loading spell checker.', 'error');
        }
    }

    function getSpellingSuggestions(words) {
        if (dictionaries.length === 0) return new Map();
        const suggestions = new Map();
        const checkedWords = new Set();
        for (const word of words) {
            const cleanWord = word.replace(/['"(),.?]/g, '');
            const lowerCleanWord = cleanWord.toLowerCase();
            if (checkedWords.has(lowerCleanWord) || cleanWord.length <= TYPO_CONFIG.ignoreLength || /\d/.test(cleanWord) || cleanWord.toUpperCase() === cleanWord) continue;
            checkedWords.add(lowerCleanWord);
            if (!dictionaries.some(dict => dict.check(cleanWord))) {
                const corrections = dictionaries[0].suggest(cleanWord);
                if (corrections && corrections.length > 0) {
                    suggestions.set(word, corrections[0]);
                }
            }
        }
        return suggestions;
    }

    function runSmartComparison() {
        if (!state.isScriptEnabled || state.isUpdatingComparison) return;
        const originalItemNameDiv = findDivByTextPrefix("Original Item Name :");
        if (!originalItemNameDiv || !domCache.cleanedItemNameTextarea) return;
        const originalBTag = originalItemNameDiv.querySelector("b");
        if (!originalBTag) return;
        const originalValue = originalBTag.textContent.trim();
        const textareaValue = domCache.cleanedItemNameTextarea.value.trim();
        const getSortedNormalizedWords = (str) => normalizeText(str).split(/\s+/).filter(Boolean).sort().join(' ');

        if (getSortedNormalizedWords(originalValue) === getSortedNormalizedWords(textareaValue)) {
            domCache.cleanedItemNameTextarea.style.backgroundColor = 'rgba(212, 237, 218, 0.2)';
            originalBTag.innerHTML = escapeHtml(originalValue);
            return;
        }
        domCache.cleanedItemNameTextarea.style.backgroundColor = "rgba(252, 242, 242, 0.3)";

        const originalWords = originalValue.split(/\s+/).filter(Boolean);
        const textareaWords = textareaValue.split(/\s+/).filter(Boolean);
        const textareaWordMap = new Map(textareaWords.map(w => [w.toLowerCase(), { word: w, used: false }]));
        const missingWords = [];

        originalWords.forEach(origWord => {
            const lowerOrigWord = origWord.toLowerCase();
            if (textareaWordMap.has(lowerOrigWord) && !textareaWordMap.get(lowerOrigWord).used) {
                textareaWordMap.get(lowerOrigWord).used = true;
                return;
            }
            let bestMatch = null;
            let minDistance = LEVENSHTEIN_TYPO_THRESHOLD;
            for (const [lowerTextWord, data] of textareaWordMap.entries()) {
                if (!data.used) {
                    const distance = levenshtein(lowerOrigWord, lowerTextWord);
                    if (distance < minDistance) {
                        minDistance = distance;
                        bestMatch = data;
                    }
                }
            }
            if (bestMatch) {
                bestMatch.used = true;
            } else {
                missingWords.push(origWord);
            }
        });

        if (missingWords.length > 0) {
            const highlightRegex = new RegExp(`\\b(${missingWords.map(regexEscape).join('|')})\\b`, 'gi');
            originalBTag.innerHTML = escapeHtml(originalValue).replace(highlightRegex,
                (match) => `<span style="background-color: #FFF3A3; border-radius: 2px;">${match}</span>`
            );
        } else {
            originalBTag.innerHTML = escapeHtml(originalValue);
        }
    }

    let isTextareaListenerAttached = false;
    function runAllComparisons() {
        if (state.isUpdatingComparison) return;
        runSmartComparison();
        if (!isTextareaListenerAttached && domCache.cleanedItemNameTextarea) {
            let debounceTimer;
            domCache.cleanedItemNameTextarea.addEventListener('input', () => {
                if (!state.isUpdatingComparison) {
                    clearTimeout(debounceTimer);
                    debounceTimer = setTimeout(runSmartComparison, 300);
                }
            }, { passive: true });
            isTextareaListenerAttached = true;
        }
    }

    let observerTimer;
    const mutationObserver = new MutationObserver(() => {
        if (!state.isUpdatingComparison && state.isScriptEnabled) {
            clearTimeout(observerTimer);
            observerTimer = setTimeout(() => {
                domCache.allDivs = [...document.querySelectorAll("div")]; // Re-cache divs
                runAllComparisons();
            }, 300);
        }
    });

    async function processGoogleSheetData() {
        try {
            const divContentMap = new Map();
            for (const prefix of MONITORED_DIV_PREFIXES) {
                const targetDiv = findDivByTextPrefix(prefix);
                if (targetDiv) {
                    const text = normalizeText(targetDiv.textContent.replace(prefix, ""));
                    divContentMap.set(prefix, text);
                }
            }
            const sheetResponse = await fetch(SHEET_URL);
            if (!sheetResponse.ok) throw new Error(`HTTP error! status: ${sheetResponse.status}`);
            const sheetData = await sheetResponse.json();
            for (const row of sheetData) {
                const keywords = row.Keyword?.split(",").map(kw => normalizeText(kw.trim())).filter(Boolean);
                if (!keywords || keywords.length === 0) continue;
                for (const text of divContentMap.values()) {
                    for (const keyword of keywords) {
                        if (text.includes(keyword)) {
                            return row;
                        }
                    }
                }
            }
            return null;
        } catch (error) {
            UI.showToast('Could not connect to Google Sheet.', 'error');
            console.error('Google Sheet fetch error:', error);
            return null;
        }
    }

    async function fillDropdown(comboboxId, valueToSelect) {
        if (!valueToSelect) return;
        const inputElement = document.querySelector(`input[aria-labelledby="${comboboxId}"]`);
        if (!inputElement) return;
        inputElement.focus();
        inputElement.click();
        inputElement.value = valueToSelect;
        inputElement.dispatchEvent(new Event("input", { bubbles: true, passive: true }));
        await delay(FAST_DELAY_MS);
        const targetOption = [...document.querySelectorAll(SELECTORS.dropdownOption)]
            .find(option => normalizeText(option.textContent) === normalizeText(valueToSelect));
        if (targetOption) {
            targetOption.click();
        } else {
            const clearButton = document.querySelector(`#${comboboxId} + .vs__actions .vs__clear`);
            if (clearButton) clearButton.click();
        }
        await delay(INTERACTION_DELAY_MS);
    }

    async function runSearchAutomation(cleanedItemName) {
        if (!domCache.searchBoxInput || !cleanedItemName) return;
        const words = cleanedItemName.split(/\s+/).filter(Boolean);
        if (words.length === 0) return;
        domCache.searchBoxInput.focus();
        domCache.searchBoxInput.click();
        await delay(FAST_DELAY_MS);
        let potentialSearchTerms = [];
        if (words.length >= 2) potentialSearchTerms.push(words.slice(0, 2).join(' '));
        potentialSearchTerms.push(words[0]);
        for (const searchTerm of potentialSearchTerms) {
            updateTextarea(domCache.searchBoxInput, searchTerm);
            domCache.searchBoxInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', bubbles: true, passive: true }));
            await delay(SEARCH_DELAY_MS);
            const currentResults = document.querySelectorAll(SELECTORS.searchResults);
            if (currentResults.length > 0) {
                let bestMatchElement = null;
                let minLevDistance = Infinity;
                const targetTextNormalized = normalizeText(cleanedItemName);
                for (const result of currentResults) {
                    const resultTextNormalized = normalizeText(result.textContent);
                    if (resultTextNormalized === targetTextNormalized) {
                        result.click();
                        return;
                    }
                    const distance = levenshtein(targetTextNormalized, resultTextNormalized);
                    if (distance < minLevDistance) {
                        minLevDistance = distance;
                        bestMatchElement = result;
                    }
                }
                if (bestMatchElement && (minLevDistance / Math.max(targetTextNormalized.length, bestMatchElement.textContent.length) < 0.3)) {
                    bestMatchElement.click();
                }
                return;
            }
        }
    }
    
    // --- NEW: Enhanced Automation Flows ---

    async function cleanAndFillName() {
        if (!state.isScriptEnabled) {
            UI.showToast('Automation is disabled.', 'error');
            return;
        }
        const originalItemNameDiv = findDivByTextPrefix("Original Item Name :");
        if (!originalItemNameDiv || !domCache.cleanedItemNameTextarea) {
            UI.showToast('Could not find item name fields.', 'error');
            return;
        }
        let cleanedName = originalItemNameDiv.querySelector("b").textContent.trim();
        const words = cleanedName.split(/\s+/).filter(Boolean);

        // 1. Apply spelling corrections
        const spellingSuggestions = getSpellingSuggestions(words);
        for (const [original, suggestion] of spellingSuggestions.entries()) {
            cleanedName = cleanedName.replace(new RegExp(`\\b${regexEscape(original)}\\b`, 'g'), suggestion);
        }

        // 2. Apply rules from sheet (remove/add keywords)
        if (state.matchedSheetRow) {
            const toRemove = state.matchedSheetRow["Remove Keywords"]?.split(',').map(k => k.trim()).filter(Boolean) || [];
            const toAdd = state.matchedSheetRow["Add Keywords"]?.split(',').map(k => k.trim()).filter(Boolean) || [];
            
            if (toRemove.length > 0) {
                const removeRegex = new RegExp(`\\b(${toRemove.map(regexEscape).join('|')})\\b`, 'gi');
                cleanedName = cleanedName.replace(removeRegex, '').replace(/\s+/g, ' ').trim();
            }
            if (toAdd.length > 0) {
                cleanedName = `${cleanedName} ${toAdd.join(' ')}`.trim();
            }
        }
        
        // 3. Apply Title Case and update
        cleanedName = toTitleCase(cleanedName);
        updateTextarea(domCache.cleanedItemNameTextarea, cleanedName);
        UI.showToast('Item Name Cleaned & Filled!');
    }


    async function runFullAutomation() {
        if (!state.isScriptEnabled) {
            UI.showToast('Automation is disabled.', 'error');
            return;
        }
        UI.updateStatus("Processing sheet...");
        state.matchedSheetRow = await processGoogleSheetData();
        if (!state.matchedSheetRow) {
            UI.updateStatus("No match found in sheet.");
            UI.showToast("No matching rule found in Google Sheet.");
            return;
        }

        UI.showToast("Sheet data loaded, filling form...");
        
        const dropdownConfigs = [
            { id: "vs1__combobox",  sheetColumn: "Vertical Name" },
            { id: "vs2__combobox",  sheetColumn: "Category" }, // More descriptive name
            { id: "vs3__combobox",  sheetColumn: "Sub-Category" },
            { id: "vs4__combobox",  sheetColumn: "Invalid Reason", defaultValue: "No Error" },
            { id: "vs5__combobox",  sheetColumn: "Contains Alcohol?" },
            { id: "vs6__combobox",  sheetColumn: "Contains Tobacco?" },
            { id: "vs7__combobox",  sheetColumn: "Is a Kit?", defaultValue: "Yes" },
            { id: "vs8__combobox",  sheetColumn: "Is a Parfait?" },
            { id: "vs17__combobox", sheetColumn: "Is a Sample?", defaultValue: "Yes" }
        ];

        for (const config of dropdownConfigs) {
            const value = state.matchedSheetRow[config.sheetColumn]?.trim() || config.defaultValue;
            await fillDropdown(config.id, value);
        }

        if (domCache.woflowBrandPathInput && domCache.woflowBrandPathInput.value.trim() === "") {
            updateTextarea(domCache.woflowBrandPathInput, "Brand Not Available");
        }

        if (domCache.cleanedItemNameTextarea && domCache.cleanedItemNameTextarea.value.trim() !== "") {
            await runSearchAutomation(domCache.cleanedItemNameTextarea.value.trim());
        }
        UI.updateStatus("Automation complete!");
    }


    // --- Main Execution Flow ---
    async function main() {
        UI.injectControls();
        
        // Cache DOM elements
        domCache.cleanedItemNameTextarea = document.querySelector(SELECTORS.cleanedItemName);
        domCache.searchBoxInput = document.querySelector(SELECTORS.searchBox);
        domCache.woflowBrandPathInput = document.querySelector(SELECTORS.brandPath);
        domCache.allDivs = [...document.querySelectorAll("div")];
        
        window.__autoFillObserver = mutationObserver;
        mutationObserver.observe(document.body, { childList: true, subtree: true });

        UI.updateStatus("Loading spell checker...");
        await loadTypoLibrary();
        UI.updateStatus("Ready.");
        
        runAllComparisons();
        await runFullAutomation(); // Run autofill on initial load
    }

    main();
})();
