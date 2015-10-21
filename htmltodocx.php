<?php

class HTMLTODOCX {
  private $section;
  private $phpword_object;
  function load($phpword_path = '') {
    if ($phpword_path === '') {
      require_once __DIR__ . '/PHPWord/src/PhpWord/Autoloader.php';
    }
    else {
      require_once $phpword_path . '/src/PhpWord/Autoloader.php';
    }
    require_once __DIR__ . '/simplehtmldom/simple_html_dom.php';
    require_once __DIR__ . '/htmltodocx_converter/h2d_htmlconverter.php';
    require_once __DIR__ . '/documentation/support_functions.php';
    \PhpOffice\PhpWord\Autoloader::register();
  }

  function prepare($html, $section = null, $phpword_object = null, $style_sheet = array()) {
    $phpword_object = new \PhpOffice\PhpWord\PhpWord();
    $section = $phpword_object->addSection();

  // HTML Dom object.
    $html_dom = new simple_html_dom();
    $html_dom->load('<html><body>' . $html . '</body></html>');
    $html_dom_array = $html_dom->find('html',0)->children();

    $paths = htmltodocx_paths();

    // Provide some initial settings.
    $initial_state = array(
      'phpword_object' => &$phpword_object,
      'base_root' => $paths['base_root'],
      'base_path' => $paths['base_path'],
      'current_style' => array('size' => '8'),
      'parents' => array(0 => 'body'),
      'list_depth' => 0,
      'context' => 'section',
      'pseudo_list' => TRUE,
      'pseudo_list_indicator_font_name' => 'Wingdings',
      'pseudo_list_indicator_font_size' => '7',
      'pseudo_list_indicator_character' => "\tl ",
      'table_allowed' => TRUE,
      'treat_div_as_paragraph' => FALSE,
      'style_sheet' => $style_sheet,
    );

    // Convert the HTML and put it into the PHPWord object.
    htmltodocx_insert_html($section, $html_dom_array[0]->nodes, $initial_state);

    // Clear the HTML dom object.
    $html_dom->clear();
    unset($html_dom);
    $this->section = &$section;
    $this->phpword_object = &$phpword_object;
  }
  function write($path) {
    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($this->phpword_object, 'Word2007');
    $objWriter->save($path);
  }
  function download($path, $filename) {
    // Create the file
    $this->write($path);

    // Create a header for download an archive.
    $headers = array(
      'Content-Type'              => 'force-download',
      'Content-Disposition'       => 'attachment; filename="' . $filename . '.docx"',
      'Content-Transfer-Encoding' => 'binary',
      'Pragma'                    => 'no-cache',
      'Cache-Control'             => 'must-revalidate, post-check=0, pre-check=0',
      'Expires'                   => '0',
      'Accept-Ranges'             => 'bytes'
    );

    // Download the archive.
    file_transfer("public://{$filename}.docx", $headers);
  }
}



