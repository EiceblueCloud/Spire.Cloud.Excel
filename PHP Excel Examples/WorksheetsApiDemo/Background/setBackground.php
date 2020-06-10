<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "SetBackground.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$picPath = "input/SpireXls.png";
$result = $apiInstance->setBackground($name, $sheetName, $picPath, $folder,null, $storage);