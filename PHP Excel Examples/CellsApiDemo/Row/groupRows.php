<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "GroupRows.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$firstIndex = 2;
$lastIndex = 6;
$hide = true;
$apiInstance->groupRows($name, $sheetName, $firstIndex, $lastIndex, $hide, $folder, $storage);
