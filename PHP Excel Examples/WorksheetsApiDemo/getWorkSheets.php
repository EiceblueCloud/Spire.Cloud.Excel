<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "GetWorkSheets_Normal.xlsx";
$folder= "input";
$storage = null;
$result = $apiInstance->getWorkSheets($name, $folder, $storage);