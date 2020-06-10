<?php

use Spire\Cloud\Excel\Sdk\Api\PropertiesApi;
use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new PropertiesApi($configuration);

$name = "DeleteDocumentProperties.xlsx";
$folder = "input";
$password = null;
$storage = null;
$apiInstance->deleteDocumentProperties($name, $password,$folder,$storage);
