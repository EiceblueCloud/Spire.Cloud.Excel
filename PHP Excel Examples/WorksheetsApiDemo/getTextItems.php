<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "GetTextItems.xlsx";
$sheetName = "Sheet2";
$folder = "input";
$storage = null;
$result = $apiInstance->getTextItems($name,$sheetName, $folder, $storage);