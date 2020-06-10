<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\ExportFormat;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorkbookApi($configuration);

$name = "ConvertWorkbook.xlsx";
#Supports formats : Xlsx, Xlsb, Xls, Ods, Pdf, Xps, Ps, Pcl
$format = ExportFormat::XLSX;
$options = null;
$password = null;
$storage = null;
$folder = "input";
$result = $apiInstance->convertWorkbook($name, $format, $options, $password, $storage, $folder);