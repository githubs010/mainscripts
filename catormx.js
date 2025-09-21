javascript: (function() {
  const openSearch = (text) => {
    if (!text || text.trim() === "") {
      return;
    }
    const searchTerm = encodeURIComponent(text.trim());
    window.open(`https://www.google.com/search?q=${searchTerm}`, "_blank");
  };

  let foundSomething = false;
  const bodyText = document.body.innerText || "";

  // Attempt 1: Search for "Original Item Name" in the body text.
  const originalMatch = bodyText.match(/Original Item Name\s*:\s*(.+)/i);
  if (originalMatch) {
    openSearch(originalMatch[1]);
    foundSomething = true;
  }

  // Attempt 2: Search for a specific textarea element.
  const cleanedEl = document.querySelector('textarea[name="Woflow Cleaned Item Name"]');
  if (cleanedEl && cleanedEl.value) {
    openSearch(cleanedEl.value);
    foundSomething = true;
  }

  // Attempt 3: Search for a specific category field.
  let categoryFound = false;
  const categoryEl = document.querySelector('[name="Woflow product_category_path"]');
  if (categoryEl) {
    const categoryText = categoryEl.value || categoryEl.innerText;
    if (categoryText && categoryText.trim()) {
      openSearch(categoryText);
      foundSomething = true;
      categoryFound = true;
    }
  }

  // Fallback for category search if the specific field wasn't found.
  if (!categoryFound) {
    const categoryMatch = bodyText.match(/Category\s*:\s*(.+)/i);
    if (categoryMatch && categoryMatch[1]) {
      openSearch(categoryMatch[1].trim());
      foundSomething = true;
    }
  }

  // Attempt 4: Search for an element with a specific aria-controls attribute.
  const ariaEl = document.querySelector('[aria-controls="vs9__listbox"]');
  if (ariaEl && ariaEl.value) {
    openSearch(ariaEl.value);
    foundSomething = true;
  }

  // Final check: If no relevant information was found, alert the user.
  if (!foundSomething) {
    alert("❌ No item name, category, or specified input field was found on the page.");
  }
})();
