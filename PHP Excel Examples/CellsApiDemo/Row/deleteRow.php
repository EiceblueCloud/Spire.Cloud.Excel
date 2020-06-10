<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "DeleteRow.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet2";
$rowIndex = 1;
$apiInstance->deleteRow($name, $sheetName, $rowIndex, $folder, $storage);
