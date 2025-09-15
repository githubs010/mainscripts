(function() {
  // Use a regular expression to find a UPC code within the page's text content.
  // This looks for "upc:", followed by 8-20 digits, optionally surrounded by quotes.
  // The `?.[2]` safely accesses the captured UPC number, or returns undefined if no match is found.
  const rawUpc = document.body.innerText.match(
    /upc\s*:\s*(['"]?)([\d\s%27"“‘’”]{8,20})\1/i
  )?.[2];

  // If a potential UPC string was found, clean it up by removing all non-digit characters
  // and then removing any leading zeros. If no raw UPC was found, this results in an empty string.
  const upc = rawUpc ? rawUpc.replace(/[^\d]/g, "").replace(/^0+/, "") : "";

  // Proceed only if a valid, cleaned UPC exists.
  if (upc) {
    // Attempt to copy the cleaned UPC to the user's clipboard.
    // A try...catch block handles cases where this permission might be denied.
    try {
      navigator.clipboard.writeText(upc);
    } catch (e) {
      console.error("Failed to copy UPC to clipboard:", e);
    }

    // An array of UPC database URLs, with the found UPC interpolated into each.
    const urls = [
      `https://www.barcodelookup.com/${upc}`,
      `https://www.upcitemdb.com/upc/${upc}`,
      `https://go-upc.com/search?q=${upc}`,
      `https://upcdatabase.org/code/${upc}`
    ];

    // Loop through the URLs and open each one in a new tab.
    // A timeout staggers the opening of tabs, which can help prevent browser popup blockers.
    urls.forEach((url, i) => {
      setTimeout(() => {
        window.open(url, "_blank");
      }, i * 150);
    });

  } else {
    // If no UPC was found on the page, alert the user.
    alert("❌ Valid UPC not found on page.");
  }
})();
