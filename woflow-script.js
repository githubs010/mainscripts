(async function() {
    // --- CONFIGURATION ---
    const SHEET_URL = "https://opensheet.elk.sh/188552daH24yAiXUux5aHvqBNWOPRZPJeve2Nd6acRBA/Sheet1";
    const FALLBACK_ADMIN = 'prasad';

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
    const LEVENSHTEIN_TYPO_THRESHOLD = 5;
    const FAST_DELAY_MS = 50;
    const INTERACTION_DELAY_MS = 100;
    const SEARCH_DELAY_MS = 400;

    const SELECTORS = {
        cleanedItemName: 'textarea[name="Woflow Cleaned Item Name"]',
        brandPath: 'input[name="Woflow brand_path"]',
        searchBox: 'input[name="search-box"]',
        searchResults: 'a.search-results',
        dropdownOption: '.vs__dropdown-option, .vs__dropdown-menu li',
        woflowCleanedSize: 'input[name="Woflow Cleaned Size"]', // Selector for Cleaned Size
        woflowCleanedUOM: 'input[aria-labelledby="vs7__combobox"]' // Specific selector for Cleaned UOM dropdown
    };

    // --- Global State and Caching ---
    let isUpdatingComparison = false;
    let isHighlightingEnabled = true;
    const dictionaries = [];
    const domCache = {
        cleanedItemNameTextarea: null,
        searchBoxInput: null,
        woflowBrandPathInput: null,
        woflowCleanedSizeInput: null,
        woflowCleanedUOMInput: null,
        allDivs: [...document.querySelectorAll("div")]
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

    async function loadTypoLibrary() {
        try {
            if (typeof Typo === 'undefined') await loadScript(TYPO_CONFIG.libURL);
            const dictPromises = TYPO_CONFIG.dictionaries.map(async dictConfig => {
                const [affResponse, dicResponse] = await Promise.all([fetch(dictConfig.affURL), fetch(dictConfig.dicURL)]);
                return new Typo(dictConfig.name, await affResponse.text(), await dicResponse.text());
            });
            if (dictionaries.length === 0) dictionaries.push(...(await Promise.all(dictPromises)));
        } catch (error) {
            console.error("Could not load Typo library.", error);
        }
    }

    function getSpellingSuggestions(words) {
        if (dictionaries.length === 0) return [];
        const suggestions = [];
        const checkedWords = new Set();
        for (const word of words) {
            const cleanWord = word.replace(/['"(),.?]/g, '');
            const lowerCleanWord = cleanWord.toLowerCase();
            if (checkedWords.has(lowerCleanWord) || cleanWord.length <= TYPO_CONFIG.ignoreLength || /\d/.test(cleanWord) || cleanWord.toUpperCase() === cleanWord) continue;
            checkedWords.add(lowerCleanWord);
            if (!dictionaries.some(dict => dict.check(cleanWord))) {
                const corrections = dictionaries[0].suggest(cleanWord);
                if (corrections && corrections.length > 0) {
                    suggestions.push({ type: 'spell', from: word, to: corrections[0] });
                }
            }
        }
        return suggestions;
    }
    
    // --- FINAL: Full comparison logic with differentiated highlighting ---
    function runSmartComparison() {
        if (!isHighlightingEnabled || isUpdatingComparison) return;
        const originalItemNameDiv = findDivByTextPrefix("Original Item Name :");
        if (!originalItemNameDiv || !domCache.cleanedItemNameTextarea || !domCache.woflowBrandPathInput) return;
        const originalBTag = originalItemNameDiv.querySelector("b");
        if (!originalBTag) return;

        // --- Get values ---
        const brandPathValue = domCache.woflowBrandPathInput.value.trim();
        const originalValue = originalBTag.textContent.trim();
        const textareaValue = domCache.cleanedItemNameTextarea.value.trim();

        // --- Word Matching Logic (Levenshtein + Exact) ---
        const getWords = (str) => str.split(/\s+/).filter(Boolean);
        const combinedOriginalWords = getWords((brandPathValue + " " + originalValue).trim());
        const textareaWords = getWords(textareaValue);

        const textareaWordMap = new Map();
        textareaWords.forEach(word => {
            const lower = word.toLowerCase();
            if (!textareaWordMap.has(lower)) textareaWordMap.set(lower, []);
            textareaWordMap.get(lower).push({ word: word, used: false });
        });

        const matchedOriginalIndices = new Set();
        // Pass 1: Exact matches
        combinedOriginalWords.forEach((origWord, index) => {
            const lowerOrigWord = origWord.toLowerCase();
            const occurrences = textareaWordMap.get(lowerOrigWord);
            if (occurrences) {
                const unused = occurrences.find(o => !o.used);
                if (unused) { unused.used = true; matchedOriginalIndices.add(index); }
            }
        });
        // Pass 2: Levenshtein typo matches
        combinedOriginalWords.forEach((origWord, index) => {
            if (matchedOriginalIndices.has(index)) return;
            const lowerOrigWord = origWord.toLowerCase();
            let bestMatch = null, minDistance = LEVENSHTEIN_TYPO_THRESHOLD;
            for (const [, occurrences] of textareaWordMap.entries()) {
                for (const occ of occurrences) {
                    if (!occ.used) {
                        const dist = levenshtein(lowerOrigWord, occ.word.toLowerCase());
                        if (dist < minDistance && (dist / Math.min(lowerOrigWord.length, occ.word.length) < 0.6)) {
                            minDistance = dist;
                            bestMatch = occ;
                        }
                    }
                }
            }
            if (bestMatch) { bestMatch.used = true; matchedOriginalIndices.add(index); }
        });

        const missingWords = combinedOriginalWords.filter((_, index) => !matchedOriginalIndices.has(index));
        const excessWords = [];
        textareaWordMap.forEach(occurrences => occurrences.forEach(occ => { if (!occ.used) excessWords.push(occ.word); }));

        // --- UI Update Logic ---
        const brandWords = getWords(brandPathValue.toLowerCase());
        const originalDisplayWords = getWords(originalValue);
        
        // --- NEW HIGHLIGHTING LOGIC ---
        const newHtml = originalDisplayWords.map(word => {
            const lowerWord = word.toLowerCase();
            const isMissing = missingWords.map(w => w.toLowerCase()).includes(lowerWord);
            const isBrandWord = brandWords.includes(lowerWord);

            if (isMissing) {
                if (isBrandWord) {
                    // It's a missing word that is part of the brand. Highlight light blue.
                    return `<span style="background-color: #d0ebff;">${escapeHtml(word)}</span>`;
                } else {
                    // It's a missing word from the original name (but not brand). Highlight yellow.
                    return `<span style="background-color: #FFF3A3;">${escapeHtml(word)}</span>`;
                }
            }
            return escapeHtml(word); // Not missing, no highlight.
        }).join(' ');

        originalBTag.innerHTML = newHtml;
        domCache.woflowBrandPathInput.style.backgroundColor = ''; // Ensure brand input is not highlighted
        
        // --- End of New Highlighting Logic ---
        
        // 3. Display Excess Words
        let excessWordsDiv = document.getElementById('excess-words-display');
        if (!excessWordsDiv) {
            excessWordsDiv = document.createElement('div');
            excessWordsDiv.id = 'excess-words-display';
            excessWordsDiv.style.cssText = 'padding: 5px; margin-top: 5px; border: 1px solid #f5c6cb; border-radius: 4px; background-color: #f8d7da; color: #721c24; font-size: 12px;';
            domCache.cleanedItemNameTextarea.parentNode.insertBefore(excessWordsDiv, domCache.cleanedItemNameTextarea.nextSibling);
        }
        excessWordsDiv.style.display = excessWords.length > 0 ? 'block' : 'none';
        if (excessWords.length > 0) {
            excessWordsDiv.innerHTML = `<strong>Excess Words:</strong> ${excessWords.map(escapeHtml).join(' ')}`;
        }

        // 4. Highlight Cleaned Textarea
        domCache.cleanedItemNameTextarea.style.backgroundColor = (missingWords.length === 0 && excessWords.length === 0) ? 'rgba(212, 237, 218, 0.2)' : 'rgba(252, 242, 242, 0.3)';
    }

    let isTextareaListenerAttached = false;
    function runAllComparisons() {
        if (isUpdatingComparison) return;
        runSmartComparison();
        if (!isTextareaListenerAttached && domCache.cleanedItemNameTextarea) {
            let debounceTimer;
            const listener = () => { if (!isUpdatingComparison) { clearTimeout(debounceTimer); debounceTimer = setTimeout(runSmartComparison, 300); } };
            domCache.cleanedItemNameTextarea.addEventListener('input', listener, { passive: true });
            domCache.woflowBrandPathInput?.addEventListener('input', listener, { passive: true });
            isTextareaListenerAttached = true;
        }
    }

    let observerTimer;
    const mutationObserver = new MutationObserver(() => {
        if (!isUpdatingComparison && isHighlightingEnabled) { clearTimeout(observerTimer); observerTimer = setTimeout(runAllComparisons, 300); }
    });

    async function processGoogleSheetData() {
        try {
            const divContentMap = new Map();
            for (const prefix of MONITORED_DIV_PREFIXES) {
                const targetDiv = findDivByTextPrefix(prefix);
                if (targetDiv) {
                    divContentMap.set(prefix, normalizeText(targetDiv.textContent.replace(prefix, "")));
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
                        if (text.includes(keyword)) return row;
                    }
                }
            }
            return null;
        } catch (error) {
            alert('❌ Could not connect to the Google Sheet. Please check your connection and try again.');
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
        const targetOption = [...document.querySelectorAll(SELECTORS.dropdownOption)].find(option => normalizeText(option.textContent) === normalizeText(valueToSelect));
        if (targetOption) {
            targetOption.click();
        } else {
            const clearButton = document.querySelector(`#${comboboxId} + .vs__actions .vs__clear`);
            if (clearButton) clearButton.click();
        }
        await delay(INTERACTION_DELAY_MS);
    }
    
    // --- AUTO-SEARCH FUNCTION (RESTORED) ---
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

    // --- Helper for UOM normalization ---
    function normalizeUOM(uom) {
        const lowerUom = uom.toLowerCase();
        if (lowerUom === 'fl oz' || lowerUom === 'floz') return 'fl oz';
        if (lowerUom === 'l') return 'L';
        if (lowerUom === 'g') return 'G';
        if (lowerUom === 'ml') return 'ml';
        if (lowerUom === 'oz') return 'oz';
        if (lowerUom === 'kg') return 'kg';
        if (lowerUom === 'lb') return 'lb';
        if (lowerUom === 'pack' || lowerUom === 'pk') return 'pack';
        if (lowerUom === 'each' || lowerUom === 'ea') return 'each';
        if (lowerUom === 'ct' || lowerUom === 'count') return 'ct';
        return lowerUom; // Return as is if not a special case
    }

    // --- NEW FEATURE: Auto-fill Size and UOM from Original Item Name ---
    async function autoFillSizeAndUOM() {
        const originalItemNameDiv = findDivByTextPrefix("Original Item Name :");
        if (!originalItemNameDiv) return;

        const originalItemNameText = originalItemNameDiv.querySelector("b")?.textContent.trim();
        if (!originalItemNameText) return;

        const sizeInput = domCache.woflowCleanedSizeInput;
        const uomInput = domCache.woflowCleanedUOMInput;

        if (!sizeInput || !uomInput) return;

        // Expanded regex to capture multiple size/unit patterns globally
        const extendedRegex = /(\d+\.?\d*)\s*(fl\s*oz|oz|ml|l|gal|pt|qt|kg|g|lb|pack|pk|case|ct|count|doz|ea|each|sq\s*ft|btl|box|can|roll|pr|pair|ctn|bag|servings|bunch|by\s*pound)\b/ig;
        
        let matches = [];
        let match;
        while ((match = extendedRegex.exec(originalItemNameText)) !== null) {
            matches.push({
                size: match[1],
                uom: normalizeUOM(match[2])
            });
        }

        if (matches.length > 0) {
            // Only fill if the fields are empty
            if (sizeInput.value.trim() === '') {
                if (matches.length > 1) {
                    // Scenario: Multiple pairs (e.g., "16OZ 4PK")
                    const extendedSizeValue = matches.map(m => `${m.size} ${m.uom}`).join(' x ');
                    updateTextarea(sizeInput, extendedSizeValue);
                } else {
                    // Scenario: Single pair (e.g., "750ML")
                    // Only put the size number in the size input
                    updateTextarea(sizeInput, matches[0].size);
                }
                await delay(INTERACTION_DELAY_MS);
            }

            // Always attempt to fill the UOM dropdown with the UOM from the first match, if it's empty
            const currentUOMValue = domCache.woflowCleanedUOMInput.value.trim();
            if (currentUOMValue === '' || currentUOMValue === 'Select an option' || currentUOMValue === 'UOM') { // Added 'UOM' as a potential default empty state
                await fillDropdown("vs7__combobox", matches[0].uom);
            }
        }
    }

    // --- Main Execution Flow ---
    domCache.cleanedItemNameTextarea = document.querySelector(SELECTORS.cleanedItemName);
    domCache.searchBoxInput = document.querySelector(SELECTORS.searchBox);
    domCache.woflowBrandPathInput = document.querySelector(SELECTORS.brandPath);
    domCache.woflowCleanedSizeInput = document.querySelector(SELECTORS.woflowCleanedSize);
    domCache.woflowCleanedUOMInput = document.querySelector(SELECTORS.woflowCleanedUOM);

    window.__autoFillObserver = mutationObserver;
    mutationObserver.observe(document.body, { childList: true, subtree: true });

    await loadTypoLibrary();
    runAllComparisons();
    const matchedSheetRow = await processGoogleSheetData();

    // Call the new auto-fill function BEFORE other dropdowns, as size/UOM might influence later steps
    await autoFillSizeAndUOM();

    if (!matchedSheetRow) return;

    const dropdownConfigurations = [
        { id: "vs1__combobox", value: matchedSheetRow?.["Vertical Name"]?.trim() }, { id: "vs2__combobox", value: matchedSheetRow?.vs2?.trim() },
        { id: "vs3__combobox", value: matchedSheetRow?.vs3?.trim() }, { id: "vs4__combobox", value: matchedSheetRow?.vs4?.trim() || "No Error" },
        { id: "vs5__combobox", value: matchedSheetRow?.vs5?.trim() }, { id: "vs6__combobox", value: matchedSheetRow?.vs6?.trim() },
        // This `vs7__combobox` will be handled by autoFillSizeAndUOM primarily.
        { id: "vs7__combobox", value: matchedSheetRow?.vs7?.trim() || "Yes" }, 
        { id: "vs8__combobox", value: matchedSheetRow?.vs8?.trim() },
        { id: "vs17__combobox", value: matchedSheetRow?.vs17?.trim() || "Yes" }
    ];

    for (const { id, value } of dropdownConfigurations) {
        await fillDropdown(id, value);
    }
    if (domCache.woflowBrandPathInput && domCache.woflowBrandPathInput.value.trim() === "") {
        updateTextarea(domCache.woflowBrandPathInput, "Brand Not Available");
    }
    // --- RUN AUTO-SEARCH (RESTORED) ---
    if (domCache.cleanedItemNameTextarea) {
        await runSearchAutomation(domCache.cleanedItemNameTextarea.value.trim());
    }
})();
