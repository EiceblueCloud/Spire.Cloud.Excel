<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "GetComment.xlsx";
$storage = null;
$folder = "input";
$sheetName = "Sheet3";
$cellName = "E17";
$result = $apiInstance->getComment($name, $sheetName, $cellName,$folder, $storage);