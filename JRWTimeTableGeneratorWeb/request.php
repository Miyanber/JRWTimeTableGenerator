<?php
header('Content-Type: application/json; charset=utf-8');

$raw = file_get_contents('php://input'); // POSTされた生のデータを受け取る
$data = json_decode($raw); // json形式をphp変数に変換

$url = $data;
$conn = curl_init();
curl_setopt($conn, CURLOPT_URL, $url);
curl_setopt($conn, CURLOPT_RETURNTRANSFER, true);
$res = curl_exec($conn);
echo $res;
curl_close($conn);

?>