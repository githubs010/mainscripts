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
                console.warn("Could not fetch Users sheet, using fallback admin.");
                return [FALLBACK_ADMIN];
            }
            const users = await response.json();
            return users.map(user => user.username.toLowerCase()).filter(Boolean);
        } catch (e) {
            console.error("Error fetching users, using fallback admin.", e);
            return [FALLBACK_ADMIN];
        }
    }

    const AUTHORIZED_USERS = await getAuthorizedUsers(SHEET_URL);
    const currentUser = window.WoflowAccessUser;

    if (!currentUser || !AUTHORIZED_USERS.includes(currentUser.toLowerCase())) {
        alert('⛔ Access Denied. You are not authorized to use this script.');
        return;
    }

    // --- CONSTANTS ---
    const TYPO_CONFIG = {
        libURL: 'https://cdn.jsdelivr.net/npm/typo-js@1.2.1/typo.js',
        dictionaries: [
            { name: 'en_US', affURL: 'https://cdn.jsdelivr.net/npm/dictionary-en-us@2.2.0/index.aff', dicURL: 'https://cdn.jsdelivr.net/npm/dictionary-en-us@2.2.0/index.dic' },
            { name: 'en_CA', affURL: 'https://cdn.jsdelivr.net/npm/dictionary-en-ca@2.0.0/index.aff', dicURL: 'https://cdn.jsdelivr.net/npm/dictionary-en-ca@2.0.0/index.dic' },
            { name: 'en_AU', affURL: 'https://cdn.jsdelivr.net/npm/dictionary-en-au/index.aff', dicURL: 'https://cdn.jsdelivr.net/npm/dictionary-en-au/index.dic' }
        ],
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

    // --- Global State ---
    let isUpdatingComparison = false;
    let isHighlightingEnabled = true;
    const dictionaries = [];
    const domCache = { allDivs: [] };

    // --- UTILITY FUNCTIONS ---
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
                try {
                    const [affResponse, dicResponse] = await Promise.all([fetch(dictConfig.affURL), fetch(dictConfig.dicURL)]);
                    if (!affResponse.ok || !dicResponse.ok) throw new Error(`Failed to fetch ${dictConfig.name}`);
                    return new Typo(dictConfig.name, await affResponse.text(), await dicResponse.text());
                } catch (e) {
                    console.error(`Could not load dictionary: ${dictConfig.name}`, e);
                    return null;
                }
            });
            const loadedDicts = (await Promise.all(dictPromises)).filter(Boolean);
            if (dictionaries.length === 0) dictionaries.push(...loadedDicts);
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
        const cleanedItemNameTextarea = document.querySelector(SELECTORS.cleanedItemName);
        if (!isHighlightingEnabled || isUpdatingComparison || !cleanedItemNameTextarea) return;

        const originalItemNameDiv = findDivByTextPrefix("Original Item Name :");
        if (!originalItemNameDiv) return;

        const originalBTag = originalItemNameDiv.querySelector("b");
        if (!originalBTag) return;

        const originalValue = originalBTag.textContent.trim();
        const textareaValue = cleanedItemNameTextarea.value.trim();

        let suggestionContainer = document.getElementById('woflow-suggestion-container');
        if (!suggestionContainer) {
            suggestionContainer = document.createElement('div');
            suggestionContainer.id = 'woflow-suggestion-container';
            Object.assign(suggestionContainer.style, { marginTop: '10px', padding: '5px', border: '1px dashed #ccc', borderRadius: '5px', minHeight: '40px', backgroundColor: 'rgba(240, 240, 240, 0.5)' });
            cleanedItemNameTextarea.insertAdjacentElement('afterend', suggestionContainer);
        }

        const getSortedNormalizedWords = (str) => normalizeText(str).split(/\s+/).filter(Boolean).sort().join(' ');
        if (getSortedNormalizedWords(originalValue) === getSortedNormalizedWords(textareaValue)) {
            cleanedItemNameTextarea.style.backgroundColor = 'rgba(212, 237, 218, 0.2)';
            originalBTag.innerHTML = escapeHtml(originalValue);
            suggestionContainer.innerHTML = '';
            return;
        }

        cleanedItemNameTextarea.style.backgroundColor = "rgba(252, 242, 242, 0.3)";

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
                    const relativeDistance = shorterLength > 0 ? distance / shorterLength : Infinity;
                    if (distance < minDistance && relativeDistance < 0.5) {
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
        textareaWordMap.forEach((data) => { if (!data.used) diffSuggestions.push({ type: 'remove', word: data.word }); });

        const missingWords = diffSuggestions.filter(s => s.type === 'add').map(s => s.word);
        if(missingWords.length > 0) {
            originalBTag.innerHTML = escapeHtml(originalValue).replace(new RegExp(`\\b(${missingWords.map(regexEscape).join('|')})\\b`, 'gi'), (match) => `<span style="background-color: #FFF3A3; border-radius: 2px;">${match}</span>`);
        } else {
            originalBTag.innerHTML = escapeHtml(originalValue);
        }

        const allSuggestions = [...diffSuggestions, ...getSpellingSuggestions(textareaWords)];
        const newSuggestionKeys = new Set(allSuggestions.map(sugg => `${sugg.type}-${sugg.from || sugg.word}-${sugg.to || ''}`));
        
        suggestionContainer.querySelectorAll('button[data-sugg-key]').forEach(btn => { if (!newSuggestionKeys.has(btn.dataset.suggKey)) btn.remove(); });
        
        allSuggestions.forEach(sugg => {
            const suggKey = `${sugg.type}-${sugg.from || sugg.word}-${sugg.to || ''}`;
            if (suggestionContainer.querySelector(`[data-sugg-key="${suggKey}"]`)) return;

            const btn = document.createElement('button');
            btn.type = 'button';
            btn.dataset.suggKey = suggKey;
            Object.assign(btn.style, { display: 'inline-block', marginTop: '5px', marginRight: '5px', padding: '3px 8px', fontSize: '13px', border: '1px solid', borderRadius: '8px', cursor: 'pointer', transition: 'transform 0.1s' });
            btn.onmouseover = () => btn.style.transform = 'translateY(-1px)';
            btn.onmouseout = () => btn.style.transform = 'translateY(0)';
            
            let newValueOnClick = cleanedItemNameTextarea.value.trim();

            if (sugg.type === 'add') {
                btn.textContent = `+ ${toTitleCase(sugg.word)}`;
                Object.assign(btn.style, { backgroundColor: '#dee2e6', color: '#343a40', borderColor: '#adb5bd' });
                newValueOnClick = (newValueOnClick + ' ' + toTitleCase(sugg.word)).trim();
            } else if (sugg.type === 'fix') {
                btn.textContent = `Fix: ${sugg.from} → ${toTitleCase(sugg.to)}`;
                Object.assign(btn.style, { backgroundColor: '#ced4da', color: '#343a40', borderColor: '#adb5bd' });
                newValueOnClick = newValueOnClick.replace(new RegExp(`\\b${regexEscape(sugg.from)}\\b`, 'gi'), toTitleCase(sugg.to));
            } else if (sugg.type === 'remove') {
                btn.textContent = `– ${sugg.word}`;
                Object.assign(btn.style, { backgroundColor: '#adb5bd', color: '#f8f9fa', borderColor: '#343a40' });
                newValueOnClick = newValueOnClick.replace(new RegExp(`\\s*\\b${regexEscape(sugg.word)}\\b`, 'gi'), '').replace(/\s+/g, ' ').trim();
            } else if (sugg.type === 'spell') {
                btn.textContent = `Spell: ${sugg.from} → ${toTitleCase(sugg.to)}`;
                Object.assign(btn.style, { backgroundColor: '#ced4da', color: '#343a40', borderColor: '#adb5bd' });
                newValueOnClick = newValueOnClick.replace(new RegExp(`\\b${regexEscape(sugg.from)}\\b`, 'gi'), toTitleCase(sugg.to));
            }

            btn.onclick = async () => {
                isUpdatingComparison = true;
                updateTextarea(document.querySelector(SELECTORS.cleanedItemName), newValueOnClick);
                await delay(INTERACTION_DELAY_MS);
                isUpdatingComparison = false;
                runSmartComparison();
            };
            suggestionContainer.appendChild(btn);
        });
    }

    // --- ⭐ START: FINAL, ROBUST LISTENER ATTACHMENT LOGIC ---
    function runAllComparisonsAndAttachListener() {
        if (isUpdatingComparison) return;

        domCache.allDivs = [...document.querySelectorAll("div")];
        const cleanedItemNameTextarea = document.querySelector(SELECTORS.cleanedItemName);
        
        if (cleanedItemNameTextarea) {
            runSmartComparison();

            // This block is the definitive fix. It guarantees the listener is always active.
            if (!cleanedItemNameTextarea.__listenerAttached) {
                let debounceTimer;
                cleanedItemNameTextarea.addEventListener('input', () => {
                    if (!isUpdatingComparison) {
                        clearTimeout(debounceTimer);
                        debounceTimer = setTimeout(runSmartComparison, 300);
                    }
                }, { passive: true });
                cleanedItemNameTextarea.__listenerAttached = true;
            }
        }
    }
    // --- ⭐ END: FINAL, ROBUST LISTENER ATTACHMENT LOGIC ---

    function closeAndResetUI() {
        const middleBottomBtn = document.getElementById('middle-bottom-close-button');
        const suggestionContainer = document.getElementById('woflow-suggestion-container');
        const cleanedItemNameTextarea = document.querySelector(SELECTORS.cleanedItemName);
        
        if (cleanedItemNameTextarea) cleanedItemNameTextarea.style.removeProperty('background-color');
        if (suggestionContainer) suggestionContainer.remove();
        if (middleBottomBtn) middleBottomBtn.remove();
        isHighlightingEnabled = false;

        const originalDiv = findDivByTextPrefix("Original Item Name :");
        if (originalDiv) {
            const originalBTag = originalDiv.querySelector("b");
            if (originalBTag) originalBTag.innerHTML = escapeHtml(originalBTag.textContent);
        }
    }

    function createMiddleBottomButton() {
        if (document.getElementById('middle-bottom-close-button')) return;
        const btn = document.createElement('button');
        btn.id = 'middle-bottom-close-button';
        btn.textContent = '✕';
        Object.assign(btn.style, { position: 'fixed', bottom: '20px', left: '50%', transform: 'translateX(-50%)', zIndex: '9999', width: '40px', height: '40px', padding: '0', background: 'rgba(40, 40, 40, 0.5)', color: '#fff', border: '1px solid rgba(0, 0, 0, 0.2)', borderRadius: '50%', cursor: 'pointer', fontSize: '18px', textAlign: 'center', boxShadow: '0 2px 5px rgba(0,0,0,0.2)', backdropFilter: 'blur(3px)' });
        btn.onclick = closeAndResetUI;
        document.body.appendChild(btn);
    }

    let observerTimer;
    const mutationObserver = new MutationObserver(() => {
        if (!isUpdatingComparison && isHighlightingEnabled) {
            clearTimeout(observerTimer);
            observerTimer = setTimeout(runAllComparisonsAndAttachListener, 300);
        }
    });

    async function processGoogleSheetData() {
        try {
            const divContentMap = new Map();
            for (const prefix of MONITORED_DIV_PREFIXES) {
                const targetDiv = findDivByTextPrefix(prefix);
                if (targetDiv) divContentMap.set(prefix, normalizeText(targetDiv.textContent.replace(prefix, "")));
            }
            const sheetResponse = await fetch(SHEET_URL);
            if (!sheetResponse.ok) throw new Error(`HTTP error! status: ${sheetResponse.status}`);
            const sheetData = await sheetResponse.json();
            for (const row of sheetData) {
                const keywords = row.Keyword?.split(",").map(kw => normalizeText(kw.trim())).filter(Boolean);
                if (keywords?.length > 0) {
                    for (const text of divContentMap.values()) {
                        if (keywords.some(keyword => text.includes(keyword))) return row;
                    }
                }
            }
            return null;
        } catch (error) {
            alert('❌ Could not connect to the Google Sheet. Please check your connection.');
            console.error('Google Sheet fetch error:', error);
            return null;
        }
    }

    async function fillDropdown(comboboxId, valueToSelect) {
        if (!valueToSelect) return;
        const inputElement = document.querySelector(`input[aria-labelledby="${comboboxId}"]`);
        if (!inputElement) return;
        inputElement.focus(); inputElement.click();
        inputElement.value = valueToSelect;
        inputElement.dispatchEvent(new Event("input", { bubbles: true, passive: true }));
        await delay(FAST_DELAY_MS);
        const targetOption = [...document.querySelectorAll(SELECTORS.dropdownOption)].find(option => normalizeText(option.textContent) === normalizeText(valueToSelect));
        if (targetOption) targetOption.click();
        else {
            const clearButton = document.querySelector(`#${comboboxId} + .vs__actions .vs__clear`);
            if (clearButton) clearButton.click();
        }
        await delay(INTERACTION_DELAY_MS);
    }

    async function runSearchAutomation() {
        const cleanedItemNameTextarea = document.querySelector(SELECTORS.cleanedItemName);
        const searchBoxInput = document.querySelector(SELECTORS.searchBox);
        if (!searchBoxInput || !cleanedItemNameTextarea) return;

        const cleanedItemName = cleanedItemNameTextarea.value.trim();
        const words = cleanedItemName.split(/\s+/).filter(Boolean);
        if (words.length === 0) return;

        searchBoxInput.focus(); searchBoxInput.click();
        await delay(FAST_DELAY_MS);

        const potentialSearchTerms = [];
        if (words.length >= 2) potentialSearchTerms.push(words.slice(0, 2).join(' '));
        potentialSearchTerms.push(words[0]);

        for (const searchTerm of potentialSearchTerms) {
            updateTextarea(searchBoxInput, searchTerm);
            searchBoxInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', bubbles: true, passive: true }));
            await delay(SEARCH_DELAY_MS);
            const currentResults = document.querySelectorAll(SELECTORS.searchResults);
            if (currentResults.length > 0) {
                let bestMatchElement = null;
                let minLevDistance = Infinity;
                const targetTextNormalized = normalizeText(cleanedItemName);
                for (const result of currentResults) {
                    const resultTextNormalized = normalizeText(result.textContent);
                    if (resultTextNormalized === targetTextNormalized) {
                        result.click(); return;
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

    // --- MAIN EXECUTION FLOW ---
    window.__autoFillObserver = mutationObserver;
    mutationObserver.observe(document.body, { childList: true, subtree: true, attributes: false });

    await loadTypoLibrary();
    createMiddleBottomButton();
    runAllComparisonsAndAttachListener(); // Initial run

    const matchedSheetRow = await processGoogleSheetData();
    if (!matchedSheetRow) return;

    const dropdownConfigurations = [
        { id: "vs1__combobox", value: matchedSheetRow?.["Vertical Name"]?.trim() }, { id: "vs2__combobox", value: matchedSheetRow?.vs2?.trim() },
        { id: "vs3__combobox", value: matchedSheetRow?.vs3?.trim() }, { id: "vs4__combobox", value: matchedSheetRow?.vs4?.trim() || "No Error" },
        { id: "vs5__combobox", value: matchedSheetRow?.vs5?.trim() }, { id: "vs6__combobox", value: matchedSheetRow?.vs6?.trim() },
        { id: "vs7__combobox", value: matchedSheetRow?.vs7?.trim() || "Yes" }, { id: "vs8__combobox", value: matchedSheetRow?.vs8?.trim() },
        { id: "vs17__combobox", value: matchedSheetRow?.vs17?.trim() || "Yes" }
    ];
    for (const { id, value } of dropdownConfigurations) await fillDropdown(id, value);

    const woflowBrandPathInput = document.querySelector(SELECTORS.brandPath);
    if (woflowBrandPathInput && woflowBrandPathInput.value.trim() === "") {
        updateTextarea(woflowBrandPathInput, "Brand Not Available");
    }
    
    await runSearchAutomation();
})();
