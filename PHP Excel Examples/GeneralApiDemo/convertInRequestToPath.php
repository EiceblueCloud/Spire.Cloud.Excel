<?php

use Spire\Cloud\Excel\Sdk\Api\GeneralApi;
use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\ExportFormat;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new GeneralApi($configuration);

$document = "D:/input/charts.xlsx";
$password = null;
$outPath ="output/ConvertInRequestToPath.xlsx";
$format = ExportFormat::XLSX; //Supported formats : Xlsx, Xlsb, Xls, Ods, Pdf, Xps, Ps, Pcl
$apiInstance->convertInRequestToPath($format, $outPath, $document, $password);
