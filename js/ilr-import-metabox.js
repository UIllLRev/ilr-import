function handleFileSelect(evt) {
    var files = evt.target.files;
    var reader = new FileReader();
    docx(files[0]).then(function (r) {
        r.mainDocument.childNodes.forEach(function (q) {
            if (q.className == 'pt-Head1-Articles') {
                var titleNode = document.getElementById("title");
                title.focus();
                title.value = q.innerHTML;
                title.blur();
            } else if (q.className == 'pt-AuthorName1-Articles') {
                document.getElementById("acf-field-ilr_author").value = q.textContent;
            } else if (q.className == 'pt-Abstract') {
                document.getElementById("excerpt").value += q.outerHTML;
                document.getElementById("content").value += q.outerHTML; 
            } else {
                document.getElementById("content").value += q.outerHTML; 
            }
        }); 
        if (r.footnotes) {
            r.footnotes.childNodes.forEach(function (q) { document.getElementById("content").value += q.outerHTML; });
        }
    });
}
document.getElementById("import_docx_file").addEventListener("change", handleFileSelect, false);
