<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\ProtectSheetParameter;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "ProtectWorksheet.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet3";
$protectParameter = new ProtectSheetParameter();
$protectParameter->setPassword("123");
$protectParameter->setProtectionType("scenarios");
$protectParameter->setAllowSelectingLockedCell("true");
$protectParameter->setAllowDeletingColumn("true");
$protectParameter->setAllowFormattingRow("true");
$result = $apiInstance->protectWorksheet($name, $sheetName,$protectParameter,$folder,$storage);