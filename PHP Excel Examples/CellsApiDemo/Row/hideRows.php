<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "HideRows.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet3";
$startRow = 1;
$totalRows = 1;
$apiInstance->hideRows($name, $sheetName, $startRow, $totalRows, $folder, $storage);
