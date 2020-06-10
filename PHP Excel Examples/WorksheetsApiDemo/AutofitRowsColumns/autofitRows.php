<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "AutofitRows_true.xlsx";
$folder = "input";
$storage = "";
$sheetName = "Sheet4";
$startRow = 17;
$endRow = 23;
#default value is false
$onlyAuto = true;
$result = $apiInstance->autofitRows($name, $sheetName, $startRow, $endRow, $onlyAuto, $folder, $storage);