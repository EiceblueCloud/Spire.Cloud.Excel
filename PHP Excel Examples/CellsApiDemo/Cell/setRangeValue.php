<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "SetRangeValue.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$cellarea = "A1:C3";
$value = "250.5";
$type = "double";
$apiInstance->setRangeValue($name, $sheetName, $cellarea, $value, $type, $folder, $storage);

