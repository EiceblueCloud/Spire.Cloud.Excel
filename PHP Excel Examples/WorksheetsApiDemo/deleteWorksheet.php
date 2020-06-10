<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "DeleteWorksheet.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet2";
$result = $apiInstance->deleteWorksheet($name, $sheetName, $folder, $storage);