<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\WorkbookProtectionRequest;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorkbookApi($configuration);

$name = "ProtectDocument_ALL.xlsx";
$folder = "input";
$storage = null;
$encryption = new WorkbookProtectionRequest();
$encryption->setPassword("123");
#supports types: ALL, READONLY, STRUCTURE, WINDOWS
$encryption->setProtectionType("ALL");
$result = $apiInstance->ProtectDocument($name, $encryption, $folder, $storage);