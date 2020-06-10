<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "UnMergeCells.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet2";
$startRow = 3;
$startColumn = 1;
$totalRows = 1;
$totalColumns = 2;
$apiInstance->unmergeCells($name, $sheetName, $startRow, $startColumn, $totalRows, $totalColumns, $folder, $storage);
