<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "DeleteColumns.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$columnIndex = 1;
$columns = 2;
$updateReference = true;
$apiInstance->deleteColumns($name, $sheetName, $columnIndex, $columns,/* updateReference,*/ $folder, $storage);
