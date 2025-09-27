javascript:(function(){
    try {
        function extractValue(label) {
            const regex = new RegExp(label + '\\s*:\\s*(.+)', 'im');
            const match = document.body.innerText.match(regex);
            return match ? match[1].trim() : '';
        }

        const brand = extractValue('Original Brand Name');
        const item = extractValue('Original Item Name');
        const desc = extractValue('Mx Provided Descriptor(s)');

        const execpatCase = (str) => {
            const preserve = ['LLC', 'LTD'];
            return str.split(/\s+/).map(w => {
                const wClean = w.replace(/[.,]/g,'');
                if(preserve.includes(wClean.toUpperCase())) return wClean.toUpperCase();
                return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
            }).join(' ');
        };

        const merged = execpatCase([brand, item, desc].filter(Boolean).join(' '));
        if(!merged) {
            alert('⚠️ No values found for the specified fields.');
            return;
        }

        navigator.clipboard.writeText(merged)
            .then(() => alert('✅ Copied (Execpat Title Case): ' + merged))
            .catch(e => alert('❌ Copy failed: ' + e.message));
    } catch(e) {
        alert('❌ Error: ' + e.message);
        console.error(e);
    }
})();
