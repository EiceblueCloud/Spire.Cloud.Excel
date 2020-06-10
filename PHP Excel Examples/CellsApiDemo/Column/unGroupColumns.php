<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "UngroupColumns.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$firstIndex = 4;
$lastIndex = 9;
$apiInstance->ungroupColumns($name, $sheetName, $firstIndex, $lastIndex, $folder, $storage);
