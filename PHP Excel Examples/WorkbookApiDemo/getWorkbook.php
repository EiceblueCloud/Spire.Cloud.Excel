<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorkbookApi($configuration);

#supports formats: Xlsx, Xlsb, Xls,xlsm,xltm,xltx, Ods
$name = "GetWorkbook.xlsx";
$password = null;
$storage = null;
$folder = "input";
$result = $apiInstance->getWorkbook($name, $password, $storage, $folder);