<?php

use Spire\Cloud\Excel\Sdk\Api\PropertiesApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new PropertiesApi($configuration);

$name = "GetDocumentProperty_Title.xlsx";
$sourcePath = "input";
$propertyName = "Title";
$password = null;
$storage = null;
$response = $apiInstance->getDocumentProperty($name, $propertyName, $password, $sourcePath, $storage);
