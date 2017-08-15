<?php

namespace SmotrovaLilit;

use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Style\ListItem;

class WordWriter
{
    /**
     * @var \PhpOffice\PhpWord\PhpWord
     */
    private $wordObject;

    /**
     * @var \PhpOffice\PhpWord\Element\Section
     */
    private $section;

    /**
     * WordWriter constructor.
     */
    public function __construct(ConfigWordWriter $config)
    {
        $this->wordObject = new \PhpOffice\PhpWord\PhpWord();
        $this->init($config);
    }

    /**
     * Установить дефолтные стили
     */
    private function init(ConfigWordWriter $config)
    {
        $this->section = $this->wordObject->addSection();
        $this->wordObject->setDefaultFontName($config->font);
        $this->wordObject->setDefaultFontSize($config->size);

        $this->wordObject->addTitleStyle(1, $config->h1Style);
        $this->wordObject->addTitleStyle(2, $config->h2Style);
        $this->wordObject->addTitleStyle(3, $config->h3Style);
    }

    /**
     * Добавить текст
     * @param $text
     */
    public function addText($text)
    {
        $text = strip_tags($text);

        $this->section->addText($text);
    }

    /**
     * Сохранить
     * @param $fileName
     */
    public function saveAs($fileName)
    {
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($this->wordObject, 'Word2007');

        $objWriter->save($fileName);
    }

    /**
     * Заголовок 1-ого уровня
     * @param $text
     */
    public function addH1($text)
    {
        $text = strip_tags($text);

        $this->section->addTextBreak();
        $this->section->addTitle($text, 1);
        $this->section->addTextBreak();

    }

    /**
     * Заголовок второго уровня
     * @param $text
     */
    public function addH2($text)
    {
        $text = strip_tags($text);

        $this->section->addTextBreak();
        $this->section->addTitle($text, 2);
        $this->section->addTextBreak();

    }

    /**
     * Заголовок третьего уровня
     * @param $text
     */
    public function addH3($text)
    {
        $text = strip_tags($text);

        $this->section->addTextBreak();
        $this->section->addTitle($text, 3);
        $this->section->addTextBreak();
    }

    /**
     * Добавить нумированный список
     * @param $data
     */
    public function addNumberList($data)
    {
        foreach ($data as $item) {
            $this->section->addListItem($item, 0, null, [
                'listType' => ListItem::TYPE_NUMBER_NESTED
            ]);
        }
    }


    /**
     * Добавить нумированный список
     * @param $data
     */
    public function addSquareFilledList($data)
    {
        foreach ($data as $item) {
            $this->section->addListItem($item, 0, null, [
                'listType' => ListItem::TYPE_SQUARE_FILLED
            ]);
        }
    }

    /**
     * Добавить многоуровневый список
     */
    public function add2levelList($data)
    {
        $multilevelNumberingStyleName = 'multilevel';
        $this->wordObject->addNumberingStyle(
            $multilevelNumberingStyleName,
            array(
                'type' => 'multilevel',
                'levels' => array(
                    array('format' => 'decimal', 'text' => '%1.', 'left' => 360, 'hanging' => 360, 'tabPos' => 360),
                    array('format' => 'upperLetter', 'text' => '%2.', 'left' => 720, 'hanging' => 360, 'tabPos' => 720),
                ),
            )
        );

        foreach ($data as $itemFirst) {
            $this->section->addListItem($itemFirst['value'], 0, null, $multilevelNumberingStyleName);
            foreach ($itemFirst['items'] as $itemSecond) {
                $this->section->addListItem($itemSecond, 1, null, $multilevelNumberingStyleName);
            }
        }
    }

    /**
     * Создает новую секцию
     */
    public function addNewSection()
    {
        $this->section = $this->wordObject->addSection();
    }

    /**
     * Добавить ссылку
     * @param $link
     * @param $text
     */
    public function addLink($link, $text)
    {
        $this->section->addLink(
            $link,
            $text,
            ['color' => '0000FF', 'underline' => \PhpOffice\PhpWord\Style\Font::UNDERLINE_SINGLE]
        );
    }

    /**
     * Добавить изображение
     * @param $imageName
     */
    public function addImage($imageName)
    {
        $this->section->addImage($imageName);
    }

    /**
     * Добавить таблицу
     */
    public function addTable($headers, $data)
    {
        $table = $this->section->addTable([
            'borderSize' => 6,
            'borderColor' => '006699',
            'cellMargin' => 80,
            'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::CENTER
        ]);

        $table->addRow();

        foreach ($headers as $cell) {
            $table->addCell(1750)->addText($cell, [
                'bold' => true,
            ]);
        }

        foreach ($data as $line) {
            $table->addRow();
            foreach ($line as $cell) {
                $table->addCell(1750, ['valign' => 'center'])->addText($cell);
            }
        }

    }

    /**
     * Добавить содержание
     */
    public function addTableContents()
    {
        $this->section->addTOC(new \PhpOffice\PhpWord\Style\Font('text', ['alignment' => Jc::CENTER]));
    }

}