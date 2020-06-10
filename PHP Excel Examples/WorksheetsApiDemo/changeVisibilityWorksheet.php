<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "ChangeVisibilityWorksheet_true.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$isVisible = true;
$result = $apiInstance->changeVisibilityWorksheet($name, $sheetName,$isVisible,$folder,$storage);