<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\Color;
use Spire\Cloud\Excel\Sdk\Model\Worksheet;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "UpdateWorksheetProperty.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet2";
$sheet = new Worksheet();
$sheet->setZoom(50);
$sheet->setIndex(3);
$sheet->setIsGridlinesVisible(false);
$sheet->setVisibilityType("hidden");
$sheet->setTabColor(new Color(array('a' => 255, 'r' => 255, 'g' => 0, 'b' => 0)));
$sheet->setIsProtected(true);
$sheet->setIsSelected(true);
$result = $apiInstance->updateWorksheetProperty($name, $sheetName, $sheet, $folder, $storage);