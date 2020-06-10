<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "CopyWorksheet_destination.xlsx";
$sourceSheet = "Sheet2";
$sourceWorkbook = "CopyWorksheet_original.xlsx";
$sourceFolder = "input";
$folder = "output";
$storage = null;
$destsheet = "Sheet4";
$result = $apiInstance->copyWorksheet($name, $destsheet, $sourceSheet, $sourceWorkbook, $sourceFolder, $folder, $storage);