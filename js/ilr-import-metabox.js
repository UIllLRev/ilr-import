function handleFileSelect(evt) {
  var files = evt.target.files;
  document.getElementById('import-docx').querySelector('.spinner').classList.add('is-active');
  docx(files[0]).then(function (r) {
    document.getElementById("content").value = '';
    r.mainDocument.childNodes.forEach(function (q) {
      if (q.className == 'pt-Head1-Articles') {
        var titleNode = document.getElementById("title");
        title.focus();
        title.value = q.innerHTML;
        title.blur();
      } else if (q.className == 'pt-AuthorName1-Articles') {
        try {
          document.querySelectorAll("[data-name='ilr_author'] input")[0].value = q.textContent;
        } catch (e) {
          // Oh well
        }
      } else if (q.className == 'pt-Abstract') {
        document.getElementById("excerpt").value += q.outerHTML;
        document.getElementById("content").value += q.outerHTML;
      } else {
        document.getElementById("content").value += q.outerHTML;
      }
    });
    if (r.footnotes) {
      r.footnotes.childNodes.forEach(q => document.getElementById("content").value += q.outerHTML);
    }
  }).then(() => {
    tinymce?.editors?.content?.load();
    document.getElementById('import-docx').querySelector('.spinner').classList.remove('is-active');
  });
}
document.getElementById("import_docx_file").addEventListener("change", handleFileSelect, false);
