<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\Border;
use Spire\Cloud\Excel\Sdk\Model\Color;
use Spire\Cloud\Excel\Sdk\Model\Font;
use Spire\Cloud\Excel\Sdk\Model\Style;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "SetColumnStyle.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$columnIndex = 2;
$style = new Style();
$font = new Font();
$font->setIsBold(true);
$font->setIsItalic(true);
$font->setUnderline("Single");
$font->setName("Algerian");
$font->setSize(8);

$style->setFont($font);
$style->setIsTextWrapped(true);
$style->setPattern("Solid");
$style->setPatternColor("red");
$style->setHorizontalAlignment("Center");
$style->setBackgroundColor(new Color(array('a' => 255, 'r' => 0, 'g' => 255, 'b' => 0)));
$borderColor1 = new Color(array('a' => 255, 'r' => 255, 'g' => 0, 'b' => 0));
$border1 = new Border(array('line_style' => "Medium", 'color' => $borderColor1, 'border_type' => "EdgeTop"));
$borderColor2 = new Color(array('a' => 255, 'r' => 0, 'g' => 255, 'b' => 0));
$border2 = new Border(array('line_style' => "DashDot", 'color' => $borderColor2, 'border_type' => "EdgeRight"));
$style->setBorderCollection(array($border1, $border2));
$apiInstance->setColumnStyle($name, $sheetName, $columnIndex, $style, $folder, $storage);
