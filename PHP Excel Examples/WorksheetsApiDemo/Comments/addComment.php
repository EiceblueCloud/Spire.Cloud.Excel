<?php

use Spire\Cloud\Excel\Sdk\Configuration;
use Spire\Cloud\Excel\Sdk\Model\Comment;

$appId = "your id";
$appKey = "your key";
$baseUrl = "https://api.e-iceblue.cn";
$configuration = new Configuration($appId, $appKey, $baseUrl);
$apiInstance = new WorksheetsApi($configuration);

$name = "AddComment.xlsx";
$folder = "input";
$storage = null;
$sheetName = "Sheet3";
$cellName = "C2:D4";
$comment = new Comment();
$comment->setAuthor("jerry");
$comment->setNote("hello");
$comment->setTextVerticalAlignment("Bottom");
$comment->setTextOrientationType("CounterClockwise");
$comment->setAutoSize(true);
$result = $apiInstance->addComment($name, $sheetName, $cellName, $comment,$folder, $storage);