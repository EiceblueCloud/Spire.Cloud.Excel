<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "InsertRows.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$startrow = 5;
$totalRows = 2;
$formatAs = "FormatAsBefore";
$apiInstance->insertRows($name, $sheetName, $startrow, $totalRows, $formatAs, $folder, $storage);
