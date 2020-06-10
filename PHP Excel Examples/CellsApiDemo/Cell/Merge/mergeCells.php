<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "MergeCells.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$startRow = 2;
$startColumn = 3;
$totalRows = 4;
$totalColumns = 4;
$apiInstance->mergeCells($name, $sheetName, $startRow, $startColumn, $totalRows, $totalColumns, $folder, $storage);
