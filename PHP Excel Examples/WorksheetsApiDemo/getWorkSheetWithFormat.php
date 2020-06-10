<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "GetWorkSheetWithFormat.xlsx";
$sheetName = "Sheet2";
#supports formats:html/pdf/png/jpeg/bmp/emf/tiff
$format = "html";
$verticalResolution = null;
$horizontalResolution = null;
$folder = "input";
$storage= null;
$result = $apiInstance->getWorkSheetWithFormat($name, $sheetName,$format, $verticalResolution,$horizontalResolution,$folder,$storage);