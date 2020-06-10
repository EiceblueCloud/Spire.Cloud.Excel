<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "CopyColumns.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$sourceColumnIndex = 2;
$destinationColumnIndex = 3;
$columnNumber = 2;
$dest_sheet = "Sheet2";//optional
$apiInstance->copyColumns($name, $sheetName, $sourceColumnIndex, $destinationColumnIndex, $columnNumber, $dest_sheet, $folder, $storage);

