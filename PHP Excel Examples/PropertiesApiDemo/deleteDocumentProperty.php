<?php

use Spire\Cloud\Excel\Sdk\Api\PropertiesApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new PropertiesApi($configuration);

$name = "DeleteDocumentProperty_Title.xlsx";
$folder = "input";
$password = null;
$storage = null;
$propertyName = "Title";
$apiInstance->deleteDocumentProperty($name, $propertyName, $password, $folder, $storage);
