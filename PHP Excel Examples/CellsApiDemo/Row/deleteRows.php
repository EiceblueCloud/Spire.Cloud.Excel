<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "DeleteRows.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$rowIndex = 2;
$totalRows = 2; //Default value : 1
$apiInstance->deleteRows($name, $sheetName, $rowIndex, $totalRows, $folder, $storage);
