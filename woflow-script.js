(async function() {
    console.clear();
    console.log("--- AUTO-FILL SCRIPT INJECTED. Waiting for page to load... ---");

    // --- Configuration ---
    const SHEET_URL = "https://opensheet.elk.sh/188552daH24yAiXUux5aHvqBNWOPRZPJeve2Nd6acRBA/Sheet1";
    const MONITORED_DIV_PREFIXES = [
        "Secondary UPC :", "Mx Provided Category 2 :", "Mx Provided Category 1 :", "Mx Provided Category 3 :",
        "Original Brand Name :", "Mx Provided Product Description :", "Original Item Name :", "Mx Provided Descriptor(s) :",
        "itemName :", "upc :"
    ];
    const REQUIRED_INPUT_ID = "vs4__combobox"; // An element that must exist before the script runs

    // --- Utility Functions ---
    const normalizeText = (text) => text?.replace(/&amp;/g, "&").replace(/\s+/g, " ").trim().toLowerCase();
    const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    function waitForElement(selector, timeout = 7000) {
        return new Promise((resolve, reject) => {
            const interval = setInterval(() => {
                const element = document.querySelector(selector);
                if (element) {
                    clearInterval(interval);
                    resolve(element);
                }
            }, 100);
            setTimeout(() => {
                clearInterval(interval);
                reject(new Error(`Required element "${selector}" did not appear in time.`));
            }, timeout);
        });
    }

    function getValueFromRow(row, columnName) {
        if (!row || !columnName) return undefined;
        // This finds the correct column even if it has hidden spaces or tabs
        const correctKey = Object.keys(row).find(key => key.trim().toLowerCase() === columnName.toLowerCase());
        return correctKey ? row[correctKey] : undefined;
    }

    async function fillDropdown(comboboxId, valueToSelect) {
        if (!valueToSelect) return;
        const inputElement = document.querySelector(`input[aria-labelledby="${comboboxId}"]`);
        if (!inputElement) {
            console.warn(`[FILL] ❌ Input for '${comboboxId}' not found.`);
            return;
        }
        inputElement.focus();
        inputElement.click();
        inputElement.value = valueToSelect;
        inputElement.dispatchEvent(new Event("input", { bubbles: true, passive: true }));
        await delay(100); // Wait for dropdown to react
        const targetOption = [...document.querySelectorAll('.vs__dropdown-option, .vs__dropdown-menu li')]
            .find(option => normalizeText(option.textContent) === normalizeText(valueToSelect));
        if (targetOption) {
            console.log(`[FILL] ✅ Selecting "${valueToSelect}" for '${comboboxId}'.`);
            targetOption.click();
        } else {
            console.warn(`[FILL] ❌ Option "${valueToSelect}" not found for '${comboboxId}'. Clearing.`);
            const clearButton = inputElement.parentElement.querySelector('.vs__clear');
            if (clearButton) clearButton.click();
        }
        await delay(150); // Wait for selection to process
    }

    // --- Main Execution Logic ---
    try {
        await waitForElement(`input[aria-labelledby="${REQUIRED_INPUT_ID}"]`);
        console.log("--- PAGE IS READY. RUNNING SCRIPT. ---");

        console.log("1. Fetching Google Sheet data...");
        const sheetResponse = await fetch(SHEET_URL);
        if (!sheetResponse.ok) throw new Error(`HTTP error! Status: ${sheetResponse.status}`);
        const sheetData = await sheetResponse.json();
        console.log("2. Sheet data loaded.");

        const pageKeywords = MONITORED_DIV_PREFIXES.map(prefix => {
            const el = [...document.querySelectorAll("div")].find(e => e.textContent.trim().startsWith(prefix));
            return el ? normalizeText(el.textContent.replace(prefix, "")) : null;
        }).filter(Boolean);
        console.log("3. Found page content to check against sheet.");

        let matchedSheetRow = null;
        for (const row of sheetData) {
            const sheetKeywords = getValueFromRow(row, "S")?.split(",").map(kw => normalizeText(kw.trim())).filter(Boolean);
            if (!sheetKeywords) continue;
            if (pageKeywords.some(pk => sheetKeywords.some(sk => pk.includes(sk)))) {
                console.log(`%c4. MATCH FOUND!`, "color: green; font-weight: bold;");
                matchedSheetRow = row;
                break;
            }
        }

        if (!matchedSheetRow) {
            throw new Error("NO MATCH FOUND between the page content and any row in your Google Sheet.");
        }

        console.log("5. Filling form fields...");
        const configurations = [
            { id: "vs1__combobox",  col: "Vertical Name", def: "" },
            { id: "vs2__combobox",  col: "vs2",             def: "" },
            { id: "vs3__combobox",  col: "vs3",             def: "" },
            { id: "vs4__combobox",  col: "vs4",             def: "No Error" },
            { id: "vs5__combobox",  col: "vs5",             def: "No Change" },
            { id: "vs6__combobox",  col: "vs6",             def: "No Change" },
            { id: "vs7__combobox",  col: "vs7",             def: "No Change" },
            { id: "vs8__combobox",  col: "vs8",             def: "Yes" },
            { id: "vs10__combobox", col: "vs10",            def: "Yes" }
        ];

        for (const config of configurations) {
            const valueFromSheet = getValueFromRow(matchedSheetRow, config.col)?.trim();
            const valueToFill = valueFromSheet || config.def;
            if (valueToFill) {
                await fillDropdown(config.id, valueToFill);
            }
        }

    } catch (error) {
        console.error("--- SCRIPT FAILED ---");
        console.error("ERROR:", error.message);
    } finally {
        console.log("--- SCRIPT FINISHED ---");
    }
})();
