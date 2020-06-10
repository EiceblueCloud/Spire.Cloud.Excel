<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\Range;
use Spire\Cloud\Excel\Sdk\Model\RangeCopyRequest;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new RangesApi($configuration);

$name = "CopyRanges.xlsx";
$folder = "input";
$storage =  null;
$sheetName = 'Sheet1';
$rangeOperate=new RangeCopyRequest();
#supports operates:copystyle/copydata/copyvalue/copyto
$rangeOperate->setOperate("copystyle");
$range= new Range();
$range->setFirstColumn(1);
$range->setFirstRow(1);
$range->setColumnCount(10);
$range->setRowCount(10);
$range2= new Range();
$range2->setFirstColumn(1);
$range2->setFirstRow(30);
$range2->setColumnCount(8);
$range2->setRowCount(5);
$rangeOperate->setSource($range);
$rangeOperate->setTarget($range2);
$result=$apiInstance-> copyRanges($name, $sheetName, $rangeOperate, $folder, $storage );