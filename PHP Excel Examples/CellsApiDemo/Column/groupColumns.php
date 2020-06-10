<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "GroupColumns.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$firstIndex = 1;
$lastIndex = 4;
$hide = true;
$apiInstance->groupColumns($name, $sheetName, $firstIndex, $lastIndex, $hide, $folder, $storage);
