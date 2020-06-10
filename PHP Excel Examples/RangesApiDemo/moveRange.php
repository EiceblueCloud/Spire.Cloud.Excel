<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\Range;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new RangesApi($configuration);

$name = "MoveRange.xlsx";
$folder = "input";
$storage =  null;
$sheetName = "Sheet1";
$range = new Range();
$range->setFirstRow(1);
$range->setFirstColumn(1);
$range->setColumnCount(1);
$range->setRowCount(1);
$destRow =25;
$destColumn =3;
$result = $apiInstance->moveRange($name, $sheetName, $range, $destRow, $destColumn, $folder, $storage);
