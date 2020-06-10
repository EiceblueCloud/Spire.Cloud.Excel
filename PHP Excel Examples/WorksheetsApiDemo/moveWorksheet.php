<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\WorksheetMovingRequest;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "MoveWorksheet_After_1.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet1";
$moving = new WorksheetMovingRequest();
$moving->setDestinationWorksheet("sheet4");
#supports positions:Before/After
$moving->setPosition("after");
$result = $apiInstance->moveWorksheet($name, $sheetName,$moving,$folder,$storage);