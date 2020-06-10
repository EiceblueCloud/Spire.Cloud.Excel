<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "GetCells.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$offest = 0;
$count = 0;
$result = $apiInstance->getCells($name, $sheetName, $offest, $count, $folder, $storage);
