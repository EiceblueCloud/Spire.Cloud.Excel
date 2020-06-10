<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "ClearFormats.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$range = "A1:G19";
$startRow = null;
$startColumn = null;
$endRow = null;
$endColumn = null;
$apiInstance->clearFormats($name, $sheetName, $range, $startRow, $startColumn, $endRow, $endColumn, $folder, $storage);
