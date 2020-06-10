<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "CopyCell.xlsx";
$storage = null;
$folder = "input";
$worksheet = "Sheet1";//Source
$sheetName = "Sheet2";//Destination
$cellName = "A13";//Source
$destCellName = "F20";
$row = 13;
$column = 2;
$apiInstance->copyCell($name, $destCellName, $sheetName, $worksheet, $cellName, $row, $column, $folder, $storage);
