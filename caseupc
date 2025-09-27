javascript:(function(){
    const body=document.body.innerText||document.body.textContent||"";
    const brandMatch=body.match(/Original Brand Name\s*:\s*(.+)/i);
    const itemMatch=body.match(/Original Item Name\s*:\s*(.+)/i);
    const descMatch=body.match(/Mx Provided Descriptor\(s\)\s*:\s*(.+)/i);
    if(brandMatch&&itemMatch&&descMatch){
        let brand=brandMatch[1].trim();
        let item=itemMatch[1].trim();
        let desc=descMatch[1].trim();
        function execpatCase(str){
            const preserve=["LLC","LTD"];
            return str.split(/\s+/).map(w=>{
                let wClean=w.replace(/[\.,]/g,"");
                if(preserve.includes(wClean.toUpperCase())){
                    return wClean.toUpperCase();
                } else {
                    return w.charAt(0).toUpperCase()+w.slice(1).toLowerCase();
                }
            }).join(" ");
        }
        let merged=execpatCase(`${brand} ${item} ${desc}`);
        navigator.clipboard.writeText(merged)
            .then(()=>alert("✅ Copied (Execpat Title Case): "+merged))
            .catch(e=>alert("❌ Copy failed: "+e));
    } else {
        alert("❌ One or more fields not found on page.");
    }
})();
