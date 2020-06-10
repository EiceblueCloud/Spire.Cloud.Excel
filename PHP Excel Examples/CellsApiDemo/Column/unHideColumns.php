<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "UnhideColumns.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet5";
$startColumn = 2;
$totalColumns = 1;
$width = null;
$apiInstance->unhideColumns($name, $sheetName, $startColumn, $totalColumns, $width, $folder, $storage);
