<?php
date_default_timezone_set('Europe/Moscow');
require_once __DIR__ . '/../../vendor/autoload.php';
define('DOMPDF_FONT_DIR', __DIR__ . '/fonts');
use Dompdf\Dompdf;

$dompdf = new Dompdf();
$content = file_get_contents(__DIR__ . '/../../resources/test.html');
$dompdf->loadHtml($content);
$dompdf->setPaper('A4', 'landscape');
$dompdf->render();
file_put_contents(__DIR__ . '/../../results/test.pdf', $dompdf->output());



