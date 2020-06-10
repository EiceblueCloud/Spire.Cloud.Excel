<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "InsertColumns.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$columnIndex = 2;
$columns = 2;
$formatAs = "FormatAsBefore";
$apiInstance->insertColumns($name, $sheetName, $columnIndex, $columns, $formatAs, $folder, $storage);
