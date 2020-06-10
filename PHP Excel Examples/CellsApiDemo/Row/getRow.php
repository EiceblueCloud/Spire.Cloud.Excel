<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "GetRow.xlsx";
$sheetName = "Sheet4";
$rowIndex = 2;
$folder = "input";
$storage = null;
$result = $apiInstance->getRow($name, $sheetName, $rowIndex, $folder, $storage);
