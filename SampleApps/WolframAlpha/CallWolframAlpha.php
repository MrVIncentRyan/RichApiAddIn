<?php	
	$input = $_GET["input"];
	$format = $_GET["format"];
	$appId = "7QLP9P-UQEKJJGEVP";
	$url = "http://api.wolframalpha.com/v2/query?input=$input&format=$format&appid=$appId";
	
	$ch = curl_init($url);
	curl_setopt_array($ch, array(
		CURLOPT_RETURNTRANSFER => TRUE
	));
	
	$response = curl_exec($ch);
	
	if ($response === FALSE) {
		die(curl_error($ch));
	}
	
	echo $response;
?>