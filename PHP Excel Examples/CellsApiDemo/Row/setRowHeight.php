<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "SetRowHeight.xlsx";
$folder = "input";
$storage = null;
$rowIndex = 3;
$height = 26;
$sheetName = "Sheet1";
$apiInstance->setRowHeight($name, $sheetName, $rowIndex, $height, $folder, $storage);
