<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "SearchText.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$text = "America";
#default value is false
$matchCase = false;
$result = $apiInstance->searchText($name, $sheetName, $text, $matchCase,$folder, $storage);