(async function() {
    // --- CONFIGURATION ---
    const SHEET_URL = "https://opensheet.elk.sh/188552daH24yAiXUux5aHvqBNWOPRZPJeve2Nd6acRBA/Sheet1";
    const FALLBACK_ADMIN = 'prasad'; // A default user if the sheet fails to load

    // --- 🔑 START: DYNAMIC ACCESS CONTROL ---
    async function getAuthorizedUsers(sheetUrl) {
        try {
            // Assumes your Google Sheet has a second tab (sheet) named "Users"
            const usersSheetUrl = sheetUrl.replace('/Sheet1', '/Users');
            const response = await fetch(usersSheetUrl);
            if (!response.ok) {
                console.error("Failed to fetch Users sheet, using fallback.");
                return [FALLBACK_ADMIN];
            }
            const users = await response.json();
            // Assumes the "Users" sheet has a column header named "username"
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
    let isUpdatingComparison = false;
    let isHighlightingEnabled = true;
    const dictionaries = [];
    const domCache = {
        cleanedItemNameTextarea: null,
        searchBoxInput: null,
        woflowBrandPathInput: null,
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

    function runSmartComparison() {
        if (!isHighlightingEnabled || isUpdatingComparison) return;
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
        const diffSuggestions = [];
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
                    const shorterLength = Math.min(lowerOrigWord.length, lowerTextWord.length);
                    const relativeDistance = shorterLength > 0 ? distance / shorterLength : (distance === 0 ? 0 : Infinity);
                    if (distance <= minDistance && relativeDistance < 0.5) {
                        minDistance = distance;
                        bestMatch = data;
                    }
                }
            }
            if (bestMatch) {
                bestMatch.used = true;
                diffSuggestions.push({ type: 'fix', from: bestMatch.word, to: origWord });
            } else {
                diffSuggestions.push({ type: 'add', word: origWord });
            }
        });
        textareaWordMap.forEach((data) => {
            if (!data.used) diffSuggestions.push({ type: 'remove', word: data.word });
        });
        const missingWords = diffSuggestions.filter(s => s.type === 'add').map(s => s.word);
        const uniqueMissingWords = [...new Set(missingWords)];
        if (uniqueMissingWords.length > 0) {
            const highlightRegex = new RegExp(`\\b(${uniqueMissingWords.map(regexEscape).join('|')})\\b`, 'gi');
            originalBTag.innerHTML = escapeHtml(originalValue).replace(highlightRegex,
                (match) => `<span style="background-color: #FFF3A3; border-radius: 2px;">${match}</span>`
            );
        } else {
            originalBTag.innerHTML = escapeHtml(originalValue);
        }
    }

    let isTextareaListenerAttached = false;
    function runAllComparisons() {
        if (isUpdatingComparison) return;
        runSmartComparison();
        if (!isTextareaListenerAttached && domCache.cleanedItemNameTextarea) {
            let debounceTimer;
            domCache.cleanedItemNameTextarea.addEventListener('input', () => {
                if (!isUpdatingComparison) {
                    clearTimeout(debounceTimer);
                    debounceTimer = setTimeout(runSmartComparison, 300);
                }
            }, { passive: true });
            isTextareaListenerAttached = true;
        }
    }

    let observerTimer;
    const mutationObserver = new MutationObserver(() => {
        if (!isUpdatingComparison && isHighlightingEnabled) {
            clearTimeout(observerTimer);
            observerTimer = setTimeout(runAllComparisons, 300);
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
            // Provide better feedback to the user
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

    // --- Main Execution Flow ---
    domCache.cleanedItemNameTextarea = document.querySelector(SELECTORS.cleanedItemName);
    domCache.searchBoxInput = document.querySelector(SELECTORS.searchBox);
    domCache.woflowBrandPathInput = document.querySelector(SELECTORS.brandPath);
    window.__autoFillObserver = mutationObserver;
    mutationObserver.observe(document.body, { childList: true, subtree: true });
    await loadTypoLibrary();
    runAllComparisons();
    const matchedSheetRow = await processGoogleSheetData();
    if (!matchedSheetRow) {
        return;
    }
    const dropdownConfigurations = [
        { id: "vs1__combobox", value: matchedSheetRow?.["Vertical Name"]?.trim() },
        { id: "vs2__combobox", value: matchedSheetRow?.vs2?.trim() },
        { id: "vs3__combobox", value: matchedSheetRow?.vs3?.trim() },
        { id: "vs4__combobox", value: matchedSheetRow?.vs4?.trim() || "No Error" },
        { id: "vs5__combobox", value: matchedSheetRow?.vs5?.trim() },
        { id: "vs6__combobox", value: matchedSheetRow?.vs6?.trim() },
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
    if (domCache.cleanedItemNameTextarea) {
        await runSearchAutomation(domCache.cleanedItemNameTextarea.value.trim());
    }
})();
