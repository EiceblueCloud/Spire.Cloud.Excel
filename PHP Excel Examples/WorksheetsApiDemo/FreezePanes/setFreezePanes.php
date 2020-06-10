<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "SetFreezePanes.xlsx";
$storage = null;
$folder = "input";
$sheetName = "Sheet2";
$freezedRows = 3;
$freezedColumns = 3;
$result = $apiInstance->setFreezePanes($name, $sheetName,$freezedRows,$freezedColumns,$folder, $storage);