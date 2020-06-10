<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "ClearContents_1.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$range = "A1:G9";
$apiInstance->clearContents($name, $sheetName, $range, null, null, null, null, $folder, $storage);
