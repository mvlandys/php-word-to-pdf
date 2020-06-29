<?php

require_once(__DIR__ . "/vendor/autoload.php");

use \PhpOffice\PhpWord\Settings;
use \PhpOffice\PhpWord\Reader;
use \PhpOffice\PhpWord\TemplateProcessor;
use \PhpOffice\PhpWord\IOFactory;

// Setup PHPWord
Settings::setPdfRendererPath('vendor/dompdf/dompdf');
Settings::setPdfRendererName('DomPDF');
$reader   = new Reader\Word2007();
$template = new TemplateProcessor('template.docx');

// Setup View Data
$title = "John Smith";
$rows  = [["key" => "101",
           "val" => "ABC"],
          ["key" => "202",
           "val" => "DEF"]];

// Apply View Data
$template->setValues(["world" => $title]);
$template->cloneRow('rowCol1', count($rows));
$x = 1;
foreach ($rows as $row) {
    $template->setValue('rowCol1#' . $x, $row["key"]);
    $template->setValue('rowCol2#' . $x, $row["val"]);
    $x++;
}

// Save the outputted file
$tmpDocFilename = md5(rand(0, 9999) . date("U"));
$template->saveAs($tmpDocFilename . ".docx");
$phpWord   = $reader->load($tmpDocFilename . ".docx");
$objWriter = IOFactory::createWriter($phpWord, 'PDF');
$objWriter->save($tmpDocFilename . ".pdf");

// Cleanup
unlink($tmpDocFilename . ".docx");

// Render
echo "Open PDF File: " . $tmpDocFilename . ".pdf";