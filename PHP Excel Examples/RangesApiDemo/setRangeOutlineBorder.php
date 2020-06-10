<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\Color;
use Spire\Cloud\Excel\Sdk\Model\Range;
use Spire\Cloud\Excel\Sdk\Model\RangeSetOutlineBorderRequest;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new RangesApi($configuration);

$name = "SetRangeOutlineBorder.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$rangeOperate=new RangeSetOutlineBorderRequest();
#supports border edges:vertical/horizontal/outline/edgeleft
$rangeOperate->setBorderEdge('outline');
#supports styles:Medium DashDot Thin Dotted Hair MediumDashDot MediumDashDotDot...etc
$rangeOperate->setBorderStyle('Medium');
$rangeOperate->setBorderColor(new Color(array('a' => 255, 'r' => 255, 'g' => 0, 'b' => 0)));
$range = new Range();
$range->setFirstRow(1);
$range->setFirstColumn(1);
$range->setColumnCount(10);
$range->setRowCount(5);
$rangeOperate->setRange($range);
$result = $apiInstance->setRangeOutlineBorder($name, $sheetName, $rangeOperate, $folder, $storage);