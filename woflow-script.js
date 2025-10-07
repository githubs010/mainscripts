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

    // Disconnect any previous observer to prevent duplicates
    if (window.__ultraRefinedObserver) {
        window.__ultraRefinedObserver.disconnect();
        document.getElementById('ultra-refined-hub')?.remove();
    }

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
    const SEARCH_DELAY_MS = 500; // Increased for more reliable results
    const CACHE_DURATION_MIN = 5;

    const SELECTORS = {
        cleanedItemName: 'textarea[name="Woflow Cleaned Item Name"]',
        brandPath: 'input[name="Woflow brand_path"]',
        searchBox: 'input[name="search-box"]',
        searchResults: 'a.search-results',
        dropdownOption: '.vs__dropdown-option, .vs__dropdown-menu li'
    };

    // --- Global State and Caching ---
    let isUpdatingComparison = false;
    const dictionaries = [];
    const domCache = {};
    const scriptState = {
        isHighlightingEnabled: true,
        isAutoFillEnabled: true,
        isSearchEnabled: true,
    };

    // --- NEW: Control Hub UI ---
    function buildControlHub() {
        const hubHTML = `
            <div id="ultra-refined-hub">
                <div id="hub-header">✨ Control Hub <span id="hub-status"></span></div>
                <div id="hub-content">
                    <label class="hub-switch">
                        <input type="checkbox" id="toggle-highlight" checked> Smart Highlight
                    </label>
                    <label class="hub-switch">
                        <input type="checkbox" id="toggle-autofill" checked> Auto Fill
                    </label>
                    <label class="hub-switch">
                        <input type="checkbox" id="toggle-search" checked> Auto Search
                    </label>
                    <button id="hub-manual-trigger">🚀 Run Manually</button>
                </div>
            </div>
        `;
        const hubCSS = `
            #ultra-refined-hub { position: fixed; top: 10px; right: 10px; width: 200px; background: #fff; border: 1px solid #ddd; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); font-family: sans-serif; font-size: 13px; z-index: 99999; color: #333; }
            #hub-header { padding: 8px 12px; background: #f7f7f7; border-bottom: 1px solid #ddd; font-weight: 600; cursor: move; border-radius: 8px 8px 0 0; }
            #hub-status { font-weight: normal; margin-left: 5px; color: #555; }
            #hub-content { padding: 12px; display: flex; flex-direction: column; gap: 10px; }
            .hub-switch { display: flex; align-items: center; gap: 6px; }
            #hub-manual-trigger { background-color: #007bff; color: white; border: none; padding: 8px; border-radius: 4px; cursor: pointer; font-weight: 600; transition: background-color 0.2s; }
            #hub-manual-trigger:hover { background-color: #0056b3; }
        `;
        document.head.insertAdjacentHTML('beforeend', `<style>${hubCSS}</style>`);
        document.body.insertAdjacentHTML('beforeend', hubHTML);

        const hub = document.getElementById('ultra-refined-hub');
        const header = document.getElementById('hub-header');

        // Make Hub draggable
        let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
        header.onmousedown = (e) => {
            e.preventDefault();
            pos3 = e.clientX;
            pos4 = e.clientY;
            document.onmouseup = () => { document.onmouseup = null; document.onmousemove = null; };
            document.onmousemove = (e) => {
                e.preventDefault();
                pos1 = pos3 - e.clientX;
                pos2 = pos4 - e.clientY;
                pos3 = e.clientX;
                pos4 = e.clientY;
                hub.style.top = (hub.offsetTop - pos2) + "px";
                hub.style.left = (hub.offsetLeft - pos1) + "px";
            };
        };

        // Event Listeners
        document.getElementById('toggle-highlight').addEventListener('change', (e) => {
             scriptState.isHighlightingEnabled = e.target.checked;
             runAllComparisons(); // Re-run to apply/remove styles
        });
        document.getElementById('toggle-autofill').addEventListener('change', (e) => scriptState.isAutoFillEnabled = e.target.checked);
        document.getElementById('toggle-search').addEventListener('change', (e) => scriptState.isSearchEnabled = e.target.checked);
        document.getElementById('hub-manual-trigger').addEventListener('click', mainExecutionFlow);
    }
    
    function setHubStatus(text, color = '#555') {
        const statusEl = document.getElementById('hub-status');
        if (statusEl) {
            statusEl.textContent = text;
            statusEl.style.color = color;
        }
    }

    // --- Utility Functions ---
    const normalizeText = (text) => text?.replace(/&amp;/g, "&").replace(/\s+/g, " ").trim().toLowerCase();
    const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));
    const regexEscape = (str) => str.replace(/[-\/\^$*+?.()|[\]{}]/g, '\$&');
    const escapeHtml = (unsafe) => unsafe.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "'");

    function findDivByTextPrefix(prefix) {
        return [...document.querySelectorAll("div")].find(e => e.textContent.trim().startsWith(prefix)) || null;
    }

    function updateTextarea(textarea, value) {
        if (textarea) {
            textarea.value = value;
            textarea.dispatchEvent(new Event('input', { bubbles: true }));
            textarea.dispatchEvent(new Event('change', { bubbles: true }));
        }
    }

    function levenshtein(s1, s2) { // (No changes to this function)
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
    
    async function loadScript(url) { // (No changes to this function)
        return new Promise((resolve, reject) => {
            const script = document.createElement('script'); script.src = url;
            script.onload = resolve; script.onerror = reject; document.head.appendChild(script);
        });
    }

    // --- Typo.js Spell Checker ---
    async function loadTypoLibrary() {
        if (dictionaries.length > 0) return;
        try {
            if (typeof Typo === 'undefined') await loadScript(TYPO_CONFIG.libURL);
            const dictPromises = TYPO_CONFIG.dictionaries.map(async dictConfig => {
                const [affResponse, dicResponse] = await Promise.all([fetch(dictConfig.affURL), fetch(dictConfig.dicURL)]);
                return new Typo(dictConfig.name, await affResponse.text(), await dicResponse.text());
            });
            dictionaries.push(...(await Promise.all(dictPromises)));
        } catch (error) { console.error("Could not load Typo library.", error); }
    }
    
    // --- MODIFIED: Interactive Smart Comparison ---
    function runSmartComparison() {
        if (isUpdatingComparison) return;
        const originalItemNameDiv = findDivByTextPrefix("Original Item Name :");
        if (!originalItemNameDiv || !domCache.cleanedItemNameTextarea) return;

        const originalBTag = originalItemNameDiv.querySelector("b");
        if (!originalBTag) return;
        
        // Always reset innerHTML to original text content to avoid nested spans
        if (originalBTag.originalContent === undefined) {
             originalBTag.originalContent = originalBTag.textContent.trim();
        }
        const originalValue = originalBTag.originalContent;
        originalBTag.innerHTML = escapeHtml(originalValue);
        
        if (!scriptState.isHighlightingEnabled) {
            domCache.cleanedItemNameTextarea.style.backgroundColor = '';
            return;
        }

        const textareaValue = domCache.cleanedItemNameTextarea.value.trim();
        const getSortedNormalizedWords = (str) => normalizeText(str).split(/\s+/).filter(Boolean).sort().join(' ');

        if (getSortedNormalizedWords(originalValue) === getSortedNormalizedWords(textareaValue)) {
            domCache.cleanedItemNameTextarea.style.backgroundColor = 'rgba(212, 237, 218, 0.4)';
            return;
        }

        domCache.cleanedItemNameTextarea.style.backgroundColor = "rgba(252, 242, 242, 0.5)";
        
        const originalWords = originalValue.split(/\s+/).filter(Boolean);
        const textareaWords = textareaValue.split(/\s+/).filter(Boolean);
        const textareaWordSet = new Set(textareaWords.map(w => w.toLowerCase()));
        
        const missingWords = originalWords.filter(origWord => !textareaWordSet.has(origWord.toLowerCase()));
        
        if (missingWords.length > 0) {
            const highlightRegex = new RegExp(`\\b(${missingWords.map(regexEscape).join('|')})\\b`, 'gi');
            originalBTag.innerHTML = escapeHtml(originalValue).replace(highlightRegex,
                (match) => `<span class="smart-suggestion" data-action="add" data-word="${escapeHtml(match)}" title="Click to add '${escapeHtml(match)}'" style="background-color: #FFF3A3; border-radius: 3px; cursor: pointer; padding: 1px 2px;">${match}</span>`
            );
        }
    }

    let isTextareaListenerAttached = false;
    function runAllComparisons() {
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
    
    // --- MODIFIED: Google Sheet Processing with Caching ---
    async function processGoogleSheetData() {
        const cacheKey = 'ultraRefinedSheetData';
        const cached = sessionStorage.getItem(cacheKey);
        if (cached) {
            const { timestamp, data } = JSON.parse(cached);
            if (Date.now() - timestamp < CACHE_DURATION_MIN * 60 * 1000) {
                console.log("Using cached sheet data.");
                return data;
            }
        }

        try {
            setHubStatus('Fetching...');
            const sheetResponse = await fetch(SHEET_URL);
            if (!sheetResponse.ok) throw new Error(`HTTP error! status: ${sheetResponse.status}`);
            const sheetData = await sheetResponse.json();
            const cachePayload = { timestamp: Date.now(), data: sheetData };
            sessionStorage.setItem(cacheKey, JSON.stringify(cachePayload));
            setHubStatus('Ready', '#28a745');
            return sheetData;
        } catch (error) {
            setHubStatus('Sheet Error', '#dc3545');
            alert('❌ Could not connect to the Google Sheet. Please check your connection and the sheet URL.');
            console.error('Google Sheet fetch error:', error);
            return null;
        }
    }

    async function findMatchingRow(sheetData) {
        if (!sheetData) return null;
        const divContentMap = new Map();
        for (const prefix of MONITORED_DIV_PREFIXES) {
            const targetDiv = findDivByTextPrefix(prefix);
            if (targetDiv) {
                const text = normalizeText(targetDiv.textContent.replace(prefix, ""));
                divContentMap.set(prefix, text);
            }
        }
        
        for (const row of sheetData) {
            const keywords = row.Keyword?.split(",").map(kw => normalizeText(kw.trim())).filter(Boolean);
            if (!keywords || keywords.length === 0) continue;
            for (const text of divContentMap.values()) {
                if (keywords.some(keyword => text.includes(keyword))) {
                    return row; // Found a match
                }
            }
        }
        return null;
    }

    // --- Automation Functions ---
    async function fillDropdown(comboboxId, valueToSelect) {
        if (!valueToSelect) return;
        const inputElement = document.querySelector(`input[aria-labelledby="${comboboxId}"]`);
        if (!inputElement) return;

        inputElement.focus();
        inputElement.click();
        await delay(FAST_DELAY_MS);
        
        const clearButton = document.querySelector(`#${comboboxId} ~ .vs__actions .vs__clear`);
        if (clearButton) clearButton.click();
        await delay(INTERACTION_DELAY_MS);
        
        inputElement.value = valueToSelect;
        inputElement.dispatchEvent(new Event("input", { bubbles: true }));
        await delay(INTERACTION_DELAY_MS);
        
        const targetOption = [...document.querySelectorAll(SELECTORS.dropdownOption)].find(option => normalizeText(option.textContent) === normalizeText(valueToSelect));
        if (targetOption) targetOption.click();
        
        await delay(INTERACTION_DELAY_MS);
    }
    
    // --- MODIFIED: More Robust Search Automation ---
    async function runSearchAutomation(cleanedItemName) {
        if (!scriptState.isSearchEnabled || !domCache.searchBoxInput || !cleanedItemName) return;
        
        domCache.searchBoxInput.focus();
        domCache.searchBoxInput.click();
        await delay(FAST_DELAY_MS);

        const words = cleanedItemName.split(/\s+/).filter(Boolean);
        const potentialSearchTerms = [];
        for (let i = words.length; i >= 2; i--) {
            potentialSearchTerms.push(words.slice(0, i).join(' '));
        }
        
        for (const searchTerm of potentialSearchTerms) {
            updateTextarea(domCache.searchBoxInput, searchTerm);
            domCache.searchBoxInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', bubbles: true }));
            await delay(SEARCH_DELAY_MS);
            
            const currentResults = document.querySelectorAll(SELECTORS.searchResults);
            if (currentResults.length > 0) {
                let bestMatchElement = null;
                let minLevDistance = Infinity;
                const targetTextNormalized = normalizeText(cleanedItemName);

                for (const result of currentResults) {
                    const resultTextNormalized = normalizeText(result.textContent);
                    const distance = levenshtein(targetTextNormalized, resultTextNormalized);
                    if (distance < minLevDistance) {
                        minLevDistance = distance;
                        bestMatchElement = result;
                    }
                }
                
                const similarity = 1 - (minLevDistance / Math.max(targetTextNormalized.length, bestMatchElement.textContent.length));
                if (bestMatchElement && similarity > 0.7) { // 70% similarity threshold
                    setHubStatus('Match Found!', '#28a745');
                    bestMatchElement.click();
                    return; // Stop searching
                }
            }
        }
        setHubStatus('No Good Match', '#ffc107');
    }

    // --- Main Execution Flow ---
    async function mainExecutionFlow() {
        setHubStatus('Running...');
        const sheetData = await processGoogleSheetData();
        const matchedSheetRow = await findMatchingRow(sheetData);
        
        if (!matchedSheetRow) {
            setHubStatus('No Rule Match', '#ffc107');
            return;
        }
        
        if (!scriptState.isAutoFillEnabled) {
            setHubStatus('AutoFill Off', '#6c757d');
            return;
        }

        setHubStatus('Filling...', '#17a2b8');
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

        await delay(500); // Wait for potential state updates from dropdowns
        
        if (domCache.cleanedItemNameTextarea) {
            setHubStatus('Searching...', '#17a2b8');
            await runSearchAutomation(domCache.cleanedItemNameTextarea.value.trim());
        }
    }
    
    // --- Initializer ---
    function initialize() {
        buildControlHub();
        
        // Cache DOM elements
        domCache.cleanedItemNameTextarea = document.querySelector(SELECTORS.cleanedItemName);
        domCache.searchBoxInput = document.querySelector(SELECTORS.searchBox);
        domCache.woflowBrandPathInput = document.querySelector(SELECTORS.brandPath);
        
        // Setup interactive suggestion listener
        document.body.addEventListener('click', (e) => {
            if (e.target.classList.contains('smart-suggestion') && e.target.dataset.action === 'add') {
                const word = e.target.dataset.word;
                if (word && domCache.cleanedItemNameTextarea) {
                    const currentVal = domCache.cleanedItemNameTextarea.value.trim();
                    updateTextarea(domCache.cleanedItemNameTextarea, `${currentVal} ${word}`);
                }
            }
        });

        // Setup Mutation Observer
        const mutationObserver = new MutationObserver(() => {
            if (!isUpdatingComparison) {
                // Re-cache key elements if they disappear/reappear
                domCache.cleanedItemNameTextarea = document.querySelector(SELECTORS.cleanedItemName);
                runAllComparisons();
            }
        });
        
        mutationObserver.observe(document.body, { childList: true, subtree: true });
        window.__ultraRefinedObserver = mutationObserver;
        
        loadTypoLibrary();
        runAllComparisons();
        mainExecutionFlow(); // Initial run
    }
    
    // Let the page settle before initializing
    setTimeout(initialize, 1000);

})();
