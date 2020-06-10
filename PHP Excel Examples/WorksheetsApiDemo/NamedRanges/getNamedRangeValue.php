<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "GetNamedRangeValue_1.xlsx";
$namerange = "Name1";
$folder = "input";
$storage = null;
$result = $apiInstance->getNamedRangeValue($name,$namerange, $folder, $storage);