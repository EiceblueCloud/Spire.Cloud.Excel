<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "SetCellValue.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$cellName = "C20";
$value = "1";
$type = "String";
$formula = null;
$apiInstance->setCellValue($name, $sheetName, $cellName, $value, $type, $formula, $folder, $storage);
