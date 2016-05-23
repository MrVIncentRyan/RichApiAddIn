<?php
	if(isset($_POST["consumer_key"]) && isset($_POST["access_token"])) {
		$consumer_key = $_POST["consumer_key"];
		$access_token = $_POST["access_token"];
		$data = array("consumer_key" => $consumer_key, "access_token" => $access_token);
		
		$ch = curl_init('https://getpocket.com/v3/get');
		curl_setopt_array($ch, array(
			CURLOPT_POST => TRUE,
			CURLOPT_RETURNTRANSFER => TRUE,
			CURLOPT_HTTPHEADER => array(
				'Content-Type: application/json; charset=UTF-8',
				'X-Accept: application/json'
			),
			CURLOPT_POSTFIELDS => json_encode($data),
			CURLOPT_SSL_VERIFYPEER => FALSE
		));
		
		$response = curl_exec($ch);
		
		if ($response === FALSE) {
			die(curl_error($ch));
		}
		
		echo $response;
	}
?>