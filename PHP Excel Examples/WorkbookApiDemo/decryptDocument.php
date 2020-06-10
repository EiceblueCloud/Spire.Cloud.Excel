<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\WorkbookEncryptionRequest;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorkbookApi($configuration);

$name = "DecryptDocument.xlsx";
$encryption=new WorkbookEncryptionRequest();
$encryption->setPassword( "123");
$storage = null;
$folder = "input";
$result = $apiInstance->DecryptDocument($name, $encryption, $folder, $storage);