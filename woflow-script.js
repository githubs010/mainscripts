(async function() {
    // --- CONFIGURATION ---
    const SHEET_URL = "https://opensheet.elk.sh/188552daH24yAiXUux5aHvqBNWOPRZPJeve2Nd6acRBA/Sheet1";
    const FALLBACK_ADMIN = 'prasad';

    // --- 🔑 DYNAMIC ACCESS CONTROL ---
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

    // --- NEW: Flexible UOM Definitions ---
    // This new list helps the script understand all variations of a unit.
    const UOM_DEFINITIONS = [
        { key: 'oz', textAliases: ['oz'], dropdownAliases: ['Ounce', 'oz'] },
        { key: 'fl oz', textAliases: ['fl oz', 'floz'], dropdownAliases: ['Fluid Ounce', 'fl oz', 'floz'] },
        { key: 'g', textAliases: ['g', 'gr'], dropdownAliases: ['Gram', 'g', 'gr'] },
        { key: 'kg', textAliases: ['kg'], dropdownAliases: ['Kilogram', 'kg'] },
        { key: 'ml', textAliases: ['ml'], dropdownAliases: ['Milliliter', 'ml'] },
        { key: 'l', textAliases: ['l'], dropdownAliases: ['Liter', 'l'] },
        { key: 'lb', textAliases: ['lb', 'lbs'], dropdownAliases: ['Pound', 'lb', 'lbs', 'by pound'] },
        { key: 'ct', textAliases: ['ct', 'count'], dropdownAliases: ['Count', 'ct', 'each'] },
        { key: 'pk', textAliases: ['pk', 'pack'], dropdownAliases: ['Pack', 'pk', 'pack'] },
        { key: 'gal', textAliases: ['gal', 'gallon'], dropdownAliases: ['Gallon', 'gal'] },
        { key: 'qt', textAliases: ['qt', 'quart'], dropdownAliases: ['Quart', 'qt'] },
        { key: 'pt', textAliases: ['pt', 'pint'], dropdownAliases: ['Pint', 'pt'] }
    ];

    // --- OPTIMIZATION: Constants ---
    const TYPO_CONFIG = {
        libURL: 'https://cdn.jsdelivr.net/npm/typo-js@1.2.1/typo.js',
        dictionaries: [{ name: 'en_US', affURL: 'https://cdn.jsdelivr.net/npm/dictionary-en-us@2.2.0/index.aff', dicURL: 'https://cdn.jsdelivr.net/npm/dictionary-en-us@2.2.0/index.dic' }],
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
        dropdownOption: '.vs__dropdown-option, .vs__dropdown-menu li',
        cleanedSize: 'input[name="Woflow Cleaned Size"]',
        cleanedUom: 'input[aria-labelledby="vs9__combobox"]',
    };

    // --- Global State and Caching ---
    let isUpdatingComparison = false;
    const dictionaries = [];
    const domCache = {
        cleanedItemNameTextarea: null, searchBoxInput: null, woflowBrandPathInput: null,
        allDivs: [], cleanedSizeInput: null, cleanedUomDropdown: null,
    };

    // --- Utility Functions ---
    const normalizeText = (text) => text?.replace(/&amp;/g, "&").replace(/\s+/g, " ").trim().toLowerCase();
    const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));
    const regexEscape = (str) => str.replace(/[-\/\^$*+?.()|[\]{}]/g, '\$&');
    const escapeHtml = (unsafe) => unsafe.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "'");

    if (window.__autoFillObserver) { window.__autoFillObserver.disconnect(); delete window.__autoFillObserver; }

    function findDivByTextPrefix(prefix) { return domCache.allDivs.find(e => e.textContent.trim().startsWith(prefix)) || null; }
    function updateTextarea(textarea, value) {
        if (textarea) {
            textarea.value = value;
            textarea.dispatchEvent(new Event('input', { bubbles: true, passive: true }));
            textarea.dispatchEvent(new Event('change', { bubbles: true, passive: true }));
        }
    }

    async function loadTypoLibrary() {
        try {
            if (typeof Typo === 'undefined') await loadScript(TYPO_CONFIG.libURL);
            if (dictionaries.length > 0) return;
            const dictPromises = TYPO_CONFIG.dictionaries.map(async dictConfig => {
                const [affResponse, dicResponse] = await Promise.all([fetch(dictConfig.affURL), fetch(dictConfig.dicURL)]);
                return new Typo(dictConfig.name, await affResponse.text(), await dicResponse.text());
            });
            dictionaries.push(...(await Promise.all(dictPromises)));
        } catch (error) { console.error("Could not load Typo library.", error); }
    }

    function runSmartComparison() {
        if (isUpdatingComparison) return;
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
                textareaWordMap.get(lowerOrigWord).used = true; return;
            }
            let bestMatch = null; let minDistance = LEVENSHTEIN_TYPO_THRESHOLD;
            for (const [lowerTextWord, data] of textareaWordMap.entries()) {
                if (!data.used) {
                    const distance = levenshtein(lowerOrigWord, lowerTextWord);
                    if (distance < minDistance) { minDistance = distance; bestMatch = data; }
                }
            }
            if (bestMatch) { bestMatch.used = true; } else { missingWords.push(origWord); }
        });
        if (missingWords.length > 0) {
            const highlightRegex = new RegExp(`\\b(${missingWords.map(regexEscape).join('|')})\\b`, 'gi');
            originalBTag.innerHTML = escapeHtml(originalValue).replace(highlightRegex, (match) => `<span style="background-color: #FFF3A3; border-radius: 2px;">${match}</span>`);
        } else { originalBTag.innerHTML = escapeHtml(originalValue); }
    }

    let isTextareaListenerAttached = false;
    function runAllComparisons() {
        if (isUpdatingComparison) return; runSmartComparison();
        if (!isTextareaListenerAttached && domCache.cleanedItemNameTextarea) {
            let debounceTimer;
            domCache.cleanedItemNameTextarea.addEventListener('input', () => {
                if (!isUpdatingComparison) { clearTimeout(debounceTimer); debounceTimer = setTimeout(runSmartComparison, 300); }
            }, { passive: true });
            isTextareaListenerAttached = true;
        }
    }

    let observerTimer;
    const mutationObserver = new MutationObserver(() => {
        if (!isUpdatingComparison) {
            clearTimeout(observerTimer);
            observerTimer = setTimeout(() => { domCache.allDivs = [...document.querySelectorAll("div")]; runAllComparisons(); }, 300);
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
                    for (const keyword of keywords) { if (text.includes(keyword)) { return row; } }
                }
            }
            return null;
        } catch (error) {
            alert('❌ Could not connect to the Google Sheet.'); console.error('Google Sheet fetch error:', error);
            return null;
        }
    }

    async function fillDropdown(comboboxSelector, valueToSelect) {
        if (!valueToSelect) return;
        const inputElement = document.querySelector(comboboxSelector);
        if (!inputElement) return;
        inputElement.focus(); inputElement.click(); await delay(FAST_DELAY_MS);
        inputElement.value = Array.isArray(valueToSelect) ? valueToSelect[0] : valueToSelect; // Use first alias for typing
        inputElement.dispatchEvent(new Event("input", { bubbles: true, passive: true }));
        await delay(FAST_DELAY_MS);
        
        // Use an array for matching
        const valuesToFind = (Array.isArray(valueToSelect) ? valueToSelect : [valueToSelect]).map(v => normalizeText(v));

        const targetOption = [...document.querySelectorAll(SELECTORS.dropdownOption)]
            .find(option => valuesToFind.includes(normalizeText(option.textContent)));
        
        if (targetOption) {
            targetOption.click();
        } else {
            const baseId = comboboxSelector.match(/vs\d+__combobox/)?.[0];
            if(baseId) {
                const clearButton = document.querySelector(`#${baseId} + .vs__actions .vs__clear`);
                if (clearButton) clearButton.click();
            }
        }
        await delay(INTERACTION_DELAY_MS);
    }
    
    // --- UPDATED: Size and UOM Extraction Function ---
    async function extractAndFillSize() {
        if (domCache.cleanedSizeInput?.value || domCache.cleanedUomDropdown?.value) {
            return;
        }

        // Dynamically build the regex from all text aliases in UOM_DEFINITIONS
        const allUomTextAliases = UOM_DEFINITIONS.flatMap(u => u.textAliases).sort((a, b) => b.length - a.length);
        const sizeRegex = new RegExp(`(\\d*\\.?\\d+)\\s*[-]?\\s*(${allUomTextAliases.map(regexEscape).join('|')})\\b`, 'i');

        const textSources = [
            findDivByTextPrefix("Original Item Name :"),
            findDivByTextPrefix("Original Size :"),
            findDivByTextPrefix("Mx Provided Product Description :")
        ];

        for (const sourceDiv of textSources) {
            if (!sourceDiv) continue;
            
            const text = sourceDiv.textContent;
            const match = text.match(sizeRegex);

            if (match) {
                const sizeValue = match[1];
                const uomFound = normalizeText(match[2]);
                
                // Find the UOM definition that matches the found text
                const uomDefinition = UOM_DEFINITIONS.find(u => u.textAliases.includes(uomFound));

                if (uomDefinition) {
                    updateTextarea(domCache.cleanedSizeInput, sizeValue);
                    // Pass all possible dropdown aliases to the fill function
                    await fillDropdown(SELECTORS.cleanedUom, uomDefinition.dropdownAliases);
                    console.log(`Filled Size: ${sizeValue}, Attempted UOM with: ${uomDefinition.dropdownAliases.join(', ')}`);
                    return; // Stop after the first successful find
                }
            }
        }
    }

    async function runAutoFill() {
        const matchedSheetRow = await processGoogleSheetData();
        if (!matchedSheetRow) {
            console.log("No matching rule found in sheet.");
            return;
        }
        console.log("Sheet data loaded, filling form...");
        const dropdownConfigs = [
            { id: "vs1__combobox",  sheetColumn: "Vertical Name" }, { id: "vs2__combobox",  sheetColumn: "vs2" },
            { id: "vs3__combobox",  sheetColumn: "vs3" }, { id: "vs4__combobox",  sheetColumn: "vs4", defaultValue: "No Error" },
            { id: "vs5__combobox",  sheetColumn: "vs5" }, { id: "vs6__combobox",  sheetColumn: "vs6" },
            { id: "vs7__combobox",  sheetColumn: "vs7", defaultValue: "Yes" }, { id: "vs8__combobox",  sheetColumn: "vs8" },
            { id: "vs17__combobox", sheetColumn: "vs17", defaultValue: "Yes" }
        ];
        for (const config of dropdownConfigs) {
            const value = matchedSheetRow[config.sheetColumn]?.trim() || config.defaultValue;
            const selector = `input[aria-labelledby="${config.id}"]`;
            await fillDropdown(selector, value);
        }
        if (domCache.woflowBrandPathInput && domCache.woflowBrandPathInput.value.trim() === "") {
            updateTextarea(domCache.woflowBrandPathInput, "Brand Not Available");
        }
        console.log("Auto-fill complete!");
    }
    
    // Levenshtein function, needed by runSmartComparison
    function levenshtein(s1, s2) {
        s1 = s1.toLowerCase(); s2 = s2.toLowerCase();
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

    // --- Main Execution Flow ---
    async function main() {
        Object.keys(SELECTORS).forEach(key => {
            const cacheKey = key === 'cleanedUom' ? 'cleanedUomDropdown' : `${key}Input`;
            if (key.includes('ItemName')) {
                 domCache[`${key}Textarea`] = document.querySelector(SELECTORS[key]);
            } else {
                 domCache[cacheKey] = document.querySelector(SELECTORS[key]);
            }
        });
        domCache.allDivs = [...document.querySelectorAll("div")];
        
        window.__autoFillObserver = mutationObserver;
        mutationObserver.observe(document.body, { childList: true, subtree: true });

        await loadTypoLibrary();
        runAllComparisons();
        await extractAndFillSize();
        await runAutoFill();
    }

    main();
})();
