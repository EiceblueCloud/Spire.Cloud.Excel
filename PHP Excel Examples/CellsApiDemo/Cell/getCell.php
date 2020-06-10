<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "GetCell.xlsx";
$sheetName = "Sheet1";
$cellOrMethodName = "A2";
$folder = "input";
$storage = null;
$result = $this->apiInstance->getCell($name, $sheetName, $cellOrMethodName, $folder, $storage);