<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\WorkbookProtectionRequest;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorkbookApi($configuration);

$name = "UnProtectDocument.xlsx";
$storage = null;
$folder = "input";
$encryption = new WorkbookProtectionRequest();
$encryption->setPassword("123");
#supports types: ALL, READONLY, STRUCTURE, WINDOWS
$encryption->setProtectionType("READONLY");
$result = $apiInstance->UnProtectDocument($name, $encryption, $folder, $storage);

