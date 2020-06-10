<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "DeleteBackground.xlsx";
$sheetName = "Sheet2";
$storage = null;
$folder = "input";
$result = $apiInstance->deleteBackground($name, $sheetName,  $folder, $storage);