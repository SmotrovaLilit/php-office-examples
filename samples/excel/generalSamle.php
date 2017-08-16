<?php
date_default_timezone_set('Europe/Moscow');
require_once __DIR__ . '/../../vendor/autoload.php';

//создание нового файла
$objPHPExcel = new PHPExcel();

$objPHPExcel->getProperties()->setCreator("Smotrova Lilit")
    ->setLastModifiedBy("Smotrova Lilit")
    ->setTitle("Office 2007 XLSX Test Document");

$objPHPExcel->getDefaultStyle()->getFont()->setName('Arial');
$objPHPExcel->getDefaultStyle()->getFont()->setSize(12);
$sheet = $objPHPExcel->getActiveSheet();
$sheet->getColumnDimension('A')->setWidth(20);
$sheet->getRowDimension('1')->setRowHeight(20);


$sheet = $objPHPExcel->getActiveSheet();
$sheet->setTitle('Чек');

$sheet->mergeCells('A2:D2');
$sheet->setCellValue('A2', 'Чек');
$styles = [
    'alignment' => [
        'horizontal' => PHPExcel_STYLE_ALIGNMENT::HORIZONTAL_CENTER,
    ],
    'fill' => [
        'type' => PHPExcel_STYLE_FILL::FILL_SOLID,
        'color'=> [
            'rgb' => 'CFCFCF'
        ]
    ],
    'font'=> [
        'bold' => true,
        'italic' => true,
        'name' => 'Times New Roman',
        'size' => 10
    ],
];
$sheet->getStyle('A2:D2')->applyFromArray($styles);

$sheet->setCellValue('A3', 'Название товара');
$sheet->setCellValue('B3', 'Цена');
$sheet->setCellValue('C3', 'Количество');
$sheet->setCellValue('D3', 'Сумма');

$sheet->setCellValue('A4', 'Сникерс');
$sheet->setCellValue('B4', 35);
$sheet->setCellValue('C4', 2);
$sheet->setCellValue('D4', '=B4*C4');

$sheet->setCellValue('A5', 'Хлеб');
$sheet->setCellValue('B5', '20');
$sheet->setCellValue('C5', '1');
$sheet->setCellValue('D5', '=B5*C5');

$sheet->setCellValue('A6', 'Итого');
$sheet->setCellValue('D6', '=sum(D4:D5)');


$sheet->getStyle('A3:D6')
    ->getBorders()
    ->getAllBorders()
    ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

$sheet->mergeCells('A6:C6');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save(__DIR__ . '/../../results/excel.xlsx');

//изменение файла
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$pExcel = $objReader->load(__DIR__ . '/../../resources/template.xlsx');

$sheet = $pExcel->getActiveSheet();
$data = [
    [
        'title'		=> 'Сникерс',
        'price'		=> 35,
        'quantity'	=> 2
    ],
    [
        'title'		=> 'Хлеб',
        'price'		=> 15,99,
        'quantity'	=> 1
    ],
    [
        'title'		=> 'Макарончик',
        'price'		=> 12,95,
        'quantity'	=> 5
    ]
];

$baseRow = 4;
foreach($data as $r => $dataRow) {
    $row = $baseRow + $r;
    $sheet->insertNewRowBefore($row,1);
    $sheet
        ->setCellValue('A'.$row, $dataRow['title'])
        ->setCellValue('B'.$row, $dataRow['price'])
        ->setCellValue( 'C'.$row, $dataRow['quantity'])
        ->setCellValue('D'.$row, '=B'.$row.'*C'.$row);
}

$sheet->setCellValue('D' . ($row + 2), '=sum(D' . $baseRow . ':D' . $row . ')');

//пример чтения
$iterator = $pExcel->getActiveSheet()->getRowIterator(4);

$readData = [];
/** @var PHPExcel_Worksheet_Row $item */
foreach ($iterator as $key1 => $item) {
    $cellIterator = $item->getCellIterator('A', 'D');
    /** @var PHPExcel_Cell $itemCell */
    foreach ($cellIterator as $key2 => $itemCell) {
        $readData[$key1][$key2] = $itemCell->getValue();
    }
}

echo "<pre>"; print_r($readData); echo "</pre>";

$objWriter = PHPExcel_IOFactory::createWriter($pExcel, 'Excel2007');
$objWriter->save(__DIR__ . '/../../results/excelTempalate.xlsx');