<?php

function htmltodocx_load($phpword_path = '') {
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
