function handleFileSelect(evt) {
    var files = evt.target.files;
    var reader = new FileReader();
    docx(files[0]).then(function (r) {
        r.mainDocument.childNodes.forEach(function (q) { document.getElementById("content").value += q.outerHTML; }); 
        r.footnotes.childNodes.forEach(function (q) { document.getElementById("content").value += q.outerHTML; });
    });
}
document.getElementById("import_docx_file").addEventListener("change", handleFileSelect, false);
