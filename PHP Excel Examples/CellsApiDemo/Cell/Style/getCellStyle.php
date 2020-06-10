<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "GetCellStyle.xlsx";
$sheetName = "Sheet4";
$cellName = "A4";
$folder = "input";;
$storage = null;
$result = $apiInstance->getCellStyle($name, $sheetName, $cellName, $folder, $storage);
