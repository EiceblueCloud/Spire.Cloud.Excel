<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "CopyRows.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$sourceRowIndex = 1;
$destinationRowIndex = 5;
$rowNumber = 3;
$dest_sheet = "Sheet2";//optional
$apiInstance->copyRows($name, $sheetName, $sourceRowIndex, $destinationRowIndex, $rowNumber, $dest_sheet, $folder, $storage);
