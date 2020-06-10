<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new RangesApi($configuration);

$name = "GetRangeValue.xlsx";
$sheetName = 'Sheet1';
$folder = "input";
$storage =  null;
#setting format:"C1:D3"
$namerange= null;
$firstRow = 1;
$firstColumn = 6;
$rowCount = 2;
$columnCount = 2;
$result = $apiInstance->getRangeValue($name, $sheetName,$namerange, $firstRow ,$firstColumn, $rowCount, $columnCount, $folder , $storage);
