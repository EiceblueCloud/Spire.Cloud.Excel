<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "RenameWorksheet_Text.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet2";
$newname = "newSheetName";
$result = $apiInstance->renameWorksheet($name, $sheetName,$newname,$folder,$storage);