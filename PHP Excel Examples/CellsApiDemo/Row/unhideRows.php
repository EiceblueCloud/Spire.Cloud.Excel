<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "UnhideRows.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet5";
$startRow = 15;
$totalRows = 4;
$height = null;
$apiInstance->unhideRows($name, $sheetName, $startRow, $totalRows, $height, $folder, $storage);
