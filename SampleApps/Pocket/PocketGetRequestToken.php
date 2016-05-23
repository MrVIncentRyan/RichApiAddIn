<?php
	if(isset($_POST["consumer_key"]) && isset($_POST["redirect_uri"])) {
		$consumer_key = $_POST["consumer_key"];
		$redirect_uri = $_POST["redirect_uri"];
		$data = array("consumer_key" => $consumer_key, "redirect_uri" => $redirect_uri);
		
		$ch = curl_init('https://getpocket.com/v3/oauth/request');
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