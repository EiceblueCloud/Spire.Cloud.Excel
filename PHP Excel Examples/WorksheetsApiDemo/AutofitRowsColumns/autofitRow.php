<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "AutofitRow.xlsx";
$folder = "input";
$storage = "";
$sheetName = "Sheet4";
#data range:1~1048576
$rowIndex = 17;
$firstColumn = 1;
$lastColumn = 1;
$result = $apiInstance->autofitRow($name, $sheetName, $rowIndex, $firstColumn, $lastColumn, $folder, $storage);