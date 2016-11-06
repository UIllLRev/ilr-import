<?php
/*
Plugin Name: ILR Import
Plugin URI:  https://illlinoislawreview.org/plugins/ilr-import
Description: Import from Word files
Version:     20161104
Author:      Matt Loar <matt@loar.name>
Author URI:  https://github.com/mloar
License:     BSD
License URI: https://directory.fsf.org/wiki/License:BSD_3Clause
Text Domain: ilr-import
Domain Path: /languages
*/

function ilr_import_posts_page() {
    //must check that the user has the required capability 
    if (!current_user_can('edit_posts')) {
        wp_die( __('You do not have sufficient permissions to access this page.') );
    }

    wp_enqueue_script("jszip", "/wp-content/plugins/ilr-import/js/jszip.min.js");
    wp_enqueue_script("docx.js", "/wp-content/plugins/ilr-import/js/docx.js");

    echo '<div class="wrap">';
    echo '<h1>Upload Word File</h1>';
    echo '<form enctype="multipart/form-data" method="POST">';
    echo '<input type="file" id="import_word_file">';
    echo '<script>';
    echo 'function handleFileSelect(evt) {';
    echo '  var files = evt.target.files;';
    echo '  var reader = new FileReader();';
    echo '  docx(files[0]).then(function (r) {';
    //echo '      document.write(r.styles.outerHTML);';
    echo '      r.mainDocument.forEach(function (q) { document.getElementById("content").value += q.outerHTML; }); });';
    echo '}';
    echo 'document.getElementById("import_word_file").addEventListener("change", handleFileSelect, false);';
    echo '</script>';
    echo '</form>';
    echo '</div>';
}

function ilr_import_admin_metabox() {
    add_meta_box('import-docx', 'Import DOCX', 'ilr_import_posts_page', 'post');
}

add_action('add_meta_boxes', 'ilr_import_admin_metabox');
?>
