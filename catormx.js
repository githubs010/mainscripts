javascript:(async function() {
    // --- CONFIGURATION ---
    const SHEET_URL = "https://opensheet.elk.sh/188552daH24yAiXUux5aHvqBNWOPRZPJeve2Nd6acRBA/Sheet1";
    const FALLBACK_ADMIN = 'prasad'; // A default user if the sheet fails to load

    // --- 🔑 DYNAMIC ACCESS CONTROL ---
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

    // --- SCRIPT EXECUTION ---

    // 1. Authorize the current user
    const authorizedUsers = await getAuthorizedUsers(SHEET_URL);
    const currentUser = prompt("Please enter your username to continue:");

    // Stop if the user cancels the prompt or is not in the authorized list
    if (!currentUser || !authorizedUsers.includes(currentUser.trim().toLowerCase())) {
        alert("❌ Access Denied. You are not an authorized user.");
        return; // Halt the script
    }

    // 2. Find the first valid search term (if authorized)
    let termToSearch = null;
    const bodyText = document.body.innerText || "";

    // Attempt 1: Look for "Original Item Name"
    const originalMatch = bodyText.match(/Original Item Name\s*:\s*(.+)/i);
    if (originalMatch && originalMatch[1]) {
        termToSearch = originalMatch[1].trim();
    }
    // Attempt 2: If not found, look for a specific textarea
    else {
        const cleanedEl = document.querySelector('textarea[name="Woflow Cleaned Item Name"]');
        if (cleanedEl && cleanedEl.value) {
            termToSearch = cleanedEl.value.trim();
        }
    }
    // Attempt 3: If not found, look for a specific category field
    else {
        const categoryEl = document.querySelector('[name="Woflow product_category_path"]');
        if (categoryEl) {
            const categoryText = categoryEl.value || categoryEl.innerText;
            if (categoryText && categoryText.trim()) {
                termToSearch = categoryText.trim();
            }
        }
    }
    // Attempt 4: Fallback for category search
    else {
        const categoryMatch = bodyText.match(/Category\s*:\s*(.+)/i);
        if (categoryMatch && categoryMatch[1]) {
            termToSearch = categoryMatch[1].trim();
        }
    }
    // Attempt 5: If not found, look for an aria-controls element
    else {
        const ariaEl = document.querySelector('[aria-controls="vs9__listbox"]');
        if (ariaEl && ariaEl.value) {
            termToSearch = ariaEl.value.trim();
        }
    }

    // 3. Perform the final action
    if (termToSearch) {
        const searchTerm = encodeURIComponent(termToSearch);
        window.open(`https://www.google.com/search?q=${searchTerm}`, "_blank");
    } else {
        alert("✅ Authorization successful, but no item name or category was found on the page.");
    }
})();
