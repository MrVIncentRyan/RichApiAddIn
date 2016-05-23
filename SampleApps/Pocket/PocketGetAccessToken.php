<?php
	if(isset($_POST["consumer_key"]) && isset($_POST["code"])) {
		$consumer_key = $_POST["consumer_key"];
		$code = $_POST["code"];
		$data = array("consumer_key" => $consumer_key, "code" => $code);
		
		$ch = curl_init('https://getpocket.com/v3/oauth/authorize');
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