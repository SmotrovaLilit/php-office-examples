<?php
use SmotrovaLilit\ConfigWordWriter;
use SmotrovaLilit\WordWriter;

date_default_timezone_set('Europe/Moscow');
require_once __DIR__ . '/../../vendor/autoload.php';

$accessToken = '5c44381c5c44381c5c44381c645c19e83855c445c44381c05d5b458d1ac0da036bbaed0';

$ownerId = '258671015';
$client = new \GuzzleHttp\Client();

$result = $client->request('GET',
    "https://api.vk.com/method/wall.get?owner_id=$ownerId&access_token=$accessToken&count=10&extended=1");
$data = json_decode($result->getBody(), true);

$config = new ConfigWordWriter();
$wordWriter = new WordWriter($config);
$wordWriter->addH1('Содержание');
$wordWriter->addTableContents();
$wordWriter->addH1("Пример работы с PhpWord");
$wordWriter->addH2("Посты со стены вк");
foreach ($data['response']['wall'] as $item) {
    if (isset($item['text']) && $item['text']) {
        $wordWriter->addH3("Пост " . $item['id']);
        $wordWriter->addText($item['text'] . "\n");
    }
}

$wordWriter->addNewSection();

$wordWriter->addH2('Пример работы со списками');

$wordWriter->addH3('Нумированный список');
$wordWriter->addNumberList([
    'el1',
    'el2',
    'el3',
    'el4',
    'el5',
]);

$wordWriter->addH3('Список с закрашенными квадратиками');
$wordWriter->addSquareFilledList([
    'el1',
    'el2',
    'el3',
    'el4',
    'el5',
]);

$wordWriter->addH3('Многоуровневый список');
$wordWriter->add2levelList([
    [
        'value' => 'el1',
        'items' => [
            'sub1',
            'sub2',
            'sub3',
        ]
    ],
    [
        'value' => 'el2',
        'items' => [
            'sub1',
            'sub2',
            'sub3',
        ]
    ],
    [
        'value' => 'el3',
        'items' => [
            'sub1',
            'sub2',
            'sub3',
        ]
    ]
]);

$wordWriter->addNewSection();
$wordWriter->addH2('Пример работы с таблицами');
$wordWriter->addTable([
    'Название товара',
    'цена',
    'колво',
    'сумма'
], [
    [
        'Сникерс',
        35,
        2,
        70
    ],
    [
        'Хлеб',
        20,
        3,
        60
    ],
    [
        'Макароны',
        150,
        2,
        150
    ],
]);

$wordWriter->addNewSection();
$wordWriter->addH2('Изображение');

$wordWriter->addImage(__DIR__ . '/../../resources/azcot_2.jpg');

$wordWriter->addNewSection();
$wordWriter->addH2('Ссылки');
$wordWriter->addLink('https://github.com/SmotrovaLilit', 'Ссылка');

$wordWriter->saveAs(__DIR__ . '/../../results/test.docx');
?>