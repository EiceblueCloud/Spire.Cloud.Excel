<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "GetRows.xlsx";
$sheetName = "Sheet1";
$folder = "input";
$storage = null;
$result = $apiInstance->getRows($name, $sheetName, $folder, $storage);
