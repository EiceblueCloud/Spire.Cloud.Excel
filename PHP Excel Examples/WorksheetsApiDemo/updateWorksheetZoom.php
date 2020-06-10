<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "UpdateWorksheetZoom.xlsx";
$sheetName = "Sheet2";
$value = 150;
$folder = "input";
$storage = null;
$result = $apiInstance->updateWorksheetZoom($name, $sheetName, $value, $folder, $storage);