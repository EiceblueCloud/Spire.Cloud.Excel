<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "DeleteFreezePanes.xlsx";
$storage = null;
$folder = "input";
$sheetName = "Sheet5";
$result = $apiInstance->deleteFreezePanes($name, $sheetName,$folder, $storage);