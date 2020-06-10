<?php

use Spire\Cloud\Excel\Sdk\Configuration;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorkbookApi($configuration);

#supports formats: Xlsx, Xlsb, Xls,xlsm,xltm,xltx, Ods
$data = "D:/data/test.xlsx";
$file = new SplFileObject($data);;
#supports formats: Xlsx, Xlsb, Xls,xlsm,xltm,xltx, Ods
$name = "CreateWorkbook.xlsx";
$inputPassword = "Password";
$password = "123";
$storage = null;
$folder = "output";
$result = $apiInstance->CreateWorkbook($name, $file, $inputPassword, $password, $storage, $folder);