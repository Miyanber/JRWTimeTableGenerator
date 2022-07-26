<?php
header('Content-Type: application/json; charset=utf-8');

$data = strstr($_SERVER["REQUEST_URI"], '?');

$url = $data;
$conn = curl_init();
curl_setopt($conn, CURLOPT_URL, $url);
curl_setopt($conn, CURLOPT_RETURNTRANSFER, true);
$res = curl_exec($conn);
echo $res;
curl_close($conn);

?>