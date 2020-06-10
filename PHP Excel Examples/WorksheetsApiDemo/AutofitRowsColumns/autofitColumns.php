<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\AutoFitterOptions;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "AutofitColumns.xlsx";
$folder = "input";
$storage = "";
$sheetName = "sheet5";
$firstColumn = 1;
$lastColumn = 5;
$autoFitterOptions = new AutoFitterOptions(null);
$autoFitterOptions->setAutoFitMergedCells(true);
$autoFitterOptions->setIgnoreHidden(true);
$autoFitterOptions->setOnlyAuto(true);
$firstRow = 3;
$lastRow = 10;
$result = $apiInstance->autofitColumns($name, $sheetName, $firstColumn, $lastColumn, $autoFitterOptions, $firstRow, $lastRow, $folder, $storage);


