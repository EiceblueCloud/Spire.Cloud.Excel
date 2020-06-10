<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "ReplaceText.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$oldValue = "North America";
$newValue = "USA";
#default value is false
$matchCase = true;
$result = $apiInstance->replaceText($name,  $sheetName, $oldValue, $newValue,$matchCase,$folder, $storage);