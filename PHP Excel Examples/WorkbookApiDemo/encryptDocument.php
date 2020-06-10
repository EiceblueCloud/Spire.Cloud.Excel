<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\WorkbookEncryptionRequest;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorkbookApi($configuration);

$name = "EncryptDocument.xlsx";
$storage = null;
$folder = "input";
$encryption = new WorkbookEncryptionRequest();
$encryption->setPassword("123");
$result = $apiInstance->EncryptDocument($name, $encryption, $folder, $storage);