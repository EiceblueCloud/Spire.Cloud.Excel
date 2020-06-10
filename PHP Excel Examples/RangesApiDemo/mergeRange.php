<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\Range;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new RangesApi($configuration);

$name = "MergeRange.xlsx";
$folder = "input";
$storage =  null;
$sheetName = "Sheet1";
$range = new Range();
$range->setFirstRow(3);
$range->setFirstColumn(3);
$range->setColumnCount(8);
$range->setRowCount(5);
$result = $apiInstance->mergeRange($name, $sheetName, $range, $folder, $storage );