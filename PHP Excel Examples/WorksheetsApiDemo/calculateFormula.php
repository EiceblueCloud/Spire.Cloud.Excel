<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "GetCalculateFormula_SUM.xlsx";
$folder = "input";
$storage = "";
$sheetName = 'Sheet4';
#supports formulas:SUM/DATE/TODAY/COUNT/IF/ROUND/SUBTOTAL/COS/
$formula = "=SUM(D8:D9)";
$result = $apiInstance->calculateFormula($name, $sheetName, $formula, $folder, $storage);