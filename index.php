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

function ilr_import_docx_metabox() {
    //must check that the user has the required capability 
    if (!current_user_can('edit_posts')) {
        wp_die( __('You do not have sufficient permissions to access this page.') );
    }

    wp_enqueue_script("jszip", "/wp-content/plugins/ilr-import/js/jszip.min.js");
    wp_enqueue_script("docx.js", "/wp-content/plugins/ilr-import/js/docx.js");
    wp_enqueue_script("ilr-import-metabox.js", "/wp-content/plugins/ilr-import/js/ilr-import-metabox.js");

    echo '<label class="screen-reader-text" for="import_docx_file">Import DOCX File</label>';
    echo '<input type="file" id="import_docx_file">';
    echo '<p>Selecting a DOCX file will replace the content of this post with the contents of the DOCX file converted to HTML.</p>';
}

function ilr_import_register_metabox() {
    add_meta_box('import-docx', 'Import DOCX', 'ilr_import_docx_metabox', 'post');
}

function ilr_import_new_print_issue() {
    if (isset($_POST['year']) && isset($_POST['issue'])) {
        wp_insert_category(array(
            'cat_name' => "Vol. {$_POST['year']} No. {$_POST['issue']}",
            'category_nicename' => "volume-{$_POST['year']}-issue-{$_POST['issue']}",
            'category_parent' => 36
        ));
    } else {
        echo '<div class="wrap">';
        echo '<h1>New Print Issue</h1>';
        echo '<form enctype="multipart/form-data" method="POST">';
        echo '<label for="year">Year</label>';
        echo '<input type="text" name="year" id="year"/>';
        echo '<label for="issue">Issue</label>';
        echo '<input type="text" name="issue" id="issue"/>';
        echo '<input type="submit" value="Add"/>';
        echo '</form>';
        echo '</div>';
    }
}

function ilr_import_add_submenu() {
    add_posts_page('Create a new print issue', 'New Print Issue', 'manage_categories', 'ilr-import-new-print-issue', 
        'ilr_import_new_print_issue');
}

add_action('add_meta_boxes', 'ilr_import_register_metabox');
add_action('admin_menu', 'ilr_import_add_submenu');
?>
