<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\ProtectSheetParameter;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "UnprotectWorksheet.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$protectParameter = new ProtectSheetParameter();
$protectParameter->setPassword("123");
$protectParameter->setAllowInsertingRow("false");
$result = $apiInstance->unprotectWorksheet($name, $sheetName,$protectParameter,$folder,$storage);