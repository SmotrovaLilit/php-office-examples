<?php
date_default_timezone_set('Europe/Moscow');
require_once __DIR__ . '/../../vendor/autoload.php';

$templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor(__DIR__ . '/resources/template.docx');

$templateProcessor->setValue('totalSum', 56);
$templateProcessor->setValue('name', 'Анна Алексеевна');

$templateProcessor->cloneRow('itemSum', 2);
$templateProcessor->setValue('itemName#1', 'Сникерс');
$templateProcessor->setValue('itemName#2', 'Хлеб');
$templateProcessor->setValue('itemPrice#1', '35');
$templateProcessor->setValue('itemPrice#2', '20');
$templateProcessor->setValue('itemCount#1', '2');
$templateProcessor->setValue('itemCount#2', '1');
$templateProcessor->setValue('itemSum#1', '70');
$templateProcessor->setValue('itemSum#2', '20');

$templateProcessor->saveAs(__DIR__ . '/../../results/resultTemplate.docx');
