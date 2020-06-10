<?php

use Spire\Cloud\Excel\Sdk\Api\PropertiesApi;
use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\DocumentProperty;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new PropertiesApi($configuration);

$name = "SetDocumentProperty.xlsx";
$folder = "input";
$password = null;
$storage = null;
$propertyName = "Keywords";
$property = new DocumentProperty(array('name' => "Keywords", 'value' => "SetDocumentProperty_xlsx", 'built_in' => true));
$apiInstance->setDocumentProperty($name, $propertyName, $property, $password, $folder, $storage);
