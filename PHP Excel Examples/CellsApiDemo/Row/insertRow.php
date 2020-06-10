<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "InsertRow.xlsx";
$folder = "Cells/Row";
$storage = null;
$sheetName = "Sheet1";
$rowIndex = 3;
$apiInstance->insertRow($name, $sheetName, $rowIndex, $folder, $storage);
