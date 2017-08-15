<?php

date_default_timezone_set('Europe/Moscow');
require_once __DIR__ . '/../../vendor/autoload.php';

$objReader = \PhpOffice\PhpWord\IOFactory::createReader('Word2007');


/** @var \PhpOffice\PhpWord\PhpWord $phpWord */
$phpWord = $objReader->load(__DIR__ . '/../../resources/template.docx');
echo "<pre>"; print_r($phpWord->getDocInfo()); echo "</pre>";