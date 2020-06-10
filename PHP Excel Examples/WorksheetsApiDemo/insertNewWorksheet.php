<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "InsertNewWorksheet_NormalSheet.xlsx";
$folder = "input";
$storage = null;
$sheettype = "NormalWorksheet";
$index = 2;
$newsheetname = "SheetInsert";
$result = $apiInstance->insertNewWorksheet($name, $newsheetname,$index, $sheettype,$folder,$storage);