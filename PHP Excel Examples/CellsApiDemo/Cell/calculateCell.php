<?php

use Spire\Cloud\Excel\Sdk\Api\CellsApi;
use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\CalculationOptions;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new CellsApi($configuration);

$name = "CalculateCell_1.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet5";
$cellName = "G6";
$options = new  CalculationOptions();
$options->setCalcStackSize(0);
$options->setPrecisionStrategy("");
$options->setIgnoreError(true);
$options->setRecursive(true);
$apiInstance->calculateCell($name, $sheetName, $cellName, $folder, $storage);