<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "DeleteComments.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet2";
$result = $apiInstance->deleteComments($name, $sheetName,$folder, $storage);