<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "UngroupRows.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$firstIndex = 5;
$lastIndex = 7;
$isAll = false;
$apiInstance->ungroupRows($name, $sheetName, $firstIndex, $lastIndex, $isAll, $folder, $storage);
