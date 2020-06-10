<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorkbookApi($configuration);

$name = "FromSource.xlsx";
$sourcePath = "input";
$sourcePassword = "Password";
$sourceStorage = null;
$password = "123";
$storage = null;
$folder = "output";
$result = $apiInstance->CreateWorkbookFromSource($name, $sourcePath, $sourceStorage, $sourceStorage, $password, $storage, $folder);