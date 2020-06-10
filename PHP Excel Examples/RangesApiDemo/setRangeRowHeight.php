<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\Range;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new RangesApi($configuration);

$name = "SetRangeRowHeight.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$range = new Range();
$range->setFirstRow(1);
$range->setFirstColumn(1);
$range->setColumnCount(5);
$range->setRowCount(10);
$value= "20";
$result = $apiInstance->setRangeRowHeight($name, $sheetName, $range, $value, $folder, $storage);