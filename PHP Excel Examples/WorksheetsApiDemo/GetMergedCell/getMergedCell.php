<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "GetMergedCell.xlsx";
$sheetName = "Sheet5";
$mergedCellIndex = 0;
$folder= "input";
$storage = null;
$result = $apiInstance->getMergedCell($name, $sheetName,$mergedCellIndex,$folder, $storage);