<?php

use Spire\Cloud\Excel\Sdk\Api\PropertiesApi;
use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\DocumentProperties;
use Spire\Cloud\Excel\Sdk\Model\DocumentProperty;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new PropertiesApi($configuration);

$name = "SetDocumentProperties.xlsx";
$folder = "input";
$password = null;
$storage = null;
$properties = new DocumentProperties();
$property1 = new DocumentProperty(array('name' => "Keywords", 'value' => "Set document properties.", 'built_in' => false));
$property2 = new DocumentProperty(array('name' => "Author", 'value' => "eiceblue", 'built_in' => false));
$property3 = new DocumentProperty(array('name' => "Company", 'value' => "冰蓝", 'built_in' => true));
$property4 = new DocumentProperty(array('name' => "Last saved by", 'value' => "123456@iCloud.com", 'built_in' => true));
$property5 = new DocumentProperty(array('name' => "Status", 'value' => "true", 'built_in' => false));
$property6 = new DocumentProperty(array('name' => "Company", 'value' => "e-iceblue", 'built_in' => true));
$property7 = new DocumentProperty(array('name' => "作者", 'value' => "冰蓝", 'built_in' => false));
$properties->setList(array($property1, $property2, $property3, $property4, $property5, $property6, $property7));
$apiInstance->setDocumentProperties($name, $properties, $password, $folder, $storage);
