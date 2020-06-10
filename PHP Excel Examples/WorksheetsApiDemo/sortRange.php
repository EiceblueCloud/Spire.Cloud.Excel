<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\DataSorter;
use Spire\Cloud\Excel\Sdk\Model\SortKey;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "SortRange.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet4";
$cellArea = "A1:G10";
$dataSorter = new DataSorter();
$dataSorter->setCaseSensitive(true);
$dataSorter->setHasHeaders(true);
$dataSorter->setSortLeftToRight(false);
$SortKey = new SortKey();
$SortKey->setKey(0);
#supports sort-orders:Ascending/Descending/Top
$SortKey->setSortOrder("Ascending");
#supports sort-types:Values/BackgroundColor/FontColor
$SortKey->setSortType( "Values");
$SortKey1 = new SortKey();
$SortKey1->setKey(1);
$SortKey1->setSortOrder("Descending");
$SortKey1->setSortType( "Values");
$KeyList = array($SortKey,$SortKey1);
$dataSorter->setKeyList($KeyList);
$result = $apiInstance->sortRange($name, $sheetName, $cellArea, $dataSorter, $folder, $storage);