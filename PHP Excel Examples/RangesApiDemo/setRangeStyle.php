<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\Border;
use Spire\Cloud\Excel\Sdk\Model\Color;
use Spire\Cloud\Excel\Sdk\Model\Font;
use Spire\Cloud\Excel\Sdk\Model\Range;
use Spire\Cloud\Excel\Sdk\Model\RangeSetStyleRequest;
use Spire\Cloud\Excel\Sdk\Model\Style;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new RangesApi($configuration);

$name = "SetRangeStyle.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$rangeOperate=new RangeSetStyleRequest();
$rangeOperate->setStyle(new Style());
$rangeOperate->getStyle()->setName("TestStyle");
$font = new Font();
$font->setColor(new Color(array('a' => 255, 'r' => 255, 'g' => 0, 'b' => 0)));
#supports underlines:single/double
$font->setUnderline("single");
$font->setSize(12);
$font->setName("Times New Roman");
$font->setIsItalic(true);
$font->setIsStrikeout(true);
$font->setIsBold(true);
$rangeOperate->getStyle()->setFont($font);
$borderCollection = array();
$borderColor1 = new Color(array('a' => 255, 'r' => 255, 'g' => 0, 'b' => 0));
$border1 = new Border(array('line_style'=>"Medium", 'color'=>$borderColor1, 'border_type'=>"EdgeTop"));
$borderColor2 = new Color(array('a' => 255, 'r' => 0, 'g' => 255, 'b' => 0));
$border2 = new Border(array('line_style'=>"DashDot", 'color'=>$borderColor2, 'border_type'=>"EdgeRight"));
$borderColor3 = new Color(array('a' => 0, 'r' => 160, 'g' => 32, 'b' => 240));
$border3 = new Border(array('line_style'=>"Thin", 'color'=>$borderColor3, 'border_type'=>"EdgeLeft"));
$borderCollection[0] = $border1;
$borderCollection[1] = $border2;
$borderCollection[2] = $border3;
$rangeOperate->getStyle()->setBorderCollection($borderCollection);
#supports horizontal alignments: left/center
$rangeOperate->getStyle()->setHorizontalAlignment("left");
$rangeOperate->getStyle()->setIndentLevel(1);
$rangeOperate->getStyle()->setTextDirection("RightToLeft");
$rangeOperate->getStyle()->setBackgroundColor(new Color(array('a' => 0, 'r' => 255, 'g' => 255, 'b' => 0)));
$rangeOperate->getStyle()->setIsTextWrapped(true);
$range = new Range();
$range->setFirstRow(1);
$range->setFirstColumn(2);
$range->setColumnCount(1);
$range->setRowCount(10);
$rangeOperate->setRange($range);
$result = $apiInstance->setRangeStyle($name, $sheetName, $rangeOperate, $folder, $storage);