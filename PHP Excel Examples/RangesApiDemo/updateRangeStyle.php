<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\Border;
use Spire\Cloud\Excel\Sdk\Model\Color;
use Spire\Cloud\Excel\Sdk\Model\Font;
use Spire\Cloud\Excel\Sdk\Model\Style;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new RangesApi($configuration);

$name = "UpdateRangeStyle.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$range = "C2:D6";
$style = new Style();
$font = new Font();
$font->setColor(new Color(array('a' => 0, 'r' => 160, 'g' => 32, 'b' => 240)));
$font->setUnderline("Single");
$font->setSize(13);
$font->setName("Times New Roman");
$font->setDoubleSize(2);
$font->setIsItalic(true);
$font->setIsStrikeout(true);
$font->setIsBold(true);
$style->setFont($font);
$borderColor1 = new Color(array('a' => 255, 'r' => 255, 'g' => 0, 'b' => 0));
$border1 = new Border(array('line_style' => "Medium", 'color' => $borderColor1, 'border_type' => "EdgeTop"));
$borderColor2 = new Color(array('a' => 255, 'r' => 0, 'g' => 255, 'b' => 0));
$border2 = new Border(array('line_style' => "DashDot", 'color' => $borderColor2, 'border_type' => "EdgeRight"));
$style->setBorderCollection(array($border1, $border2));
$style->setIndentLevel(2);
#supports horizontal alignments: left/center
$style->setHorizontalAlignment("left");
$result = $apiInstance->updateRangeStyle($name, $sheetName, $range, $style, $folder, $storage);