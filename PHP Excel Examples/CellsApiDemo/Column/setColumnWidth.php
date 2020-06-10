<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "SetColumnWidth.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$columnIndex = 2;
$width = 40;
$apiInstance->setColumnWidth($name, $sheetName, $columnIndex, $width, $folder, $storage);
