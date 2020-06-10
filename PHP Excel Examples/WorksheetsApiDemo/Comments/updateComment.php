<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\Comment;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "UpdateComment.xlsx";
$storage = null;
$folder = "input";
$sheetName = "Sheet3";
$cellName = "C6";
$comment = new Comment();
$comment->setAuthor("jerry");
$comment->setNote("Hello,Comments");
$comment->setWidth( 160);
$comment->setHeight(30);
$comment->setTextOrientationType("TopToBottom");
$comment->setIsVisible(true);
$comment->setAutoSize(true);
$comment->setTextHorizontalAlignment("center");
$result = $apiInstance->updateComment($name, $sheetName, $cellName,$comment,$folder, $storage);