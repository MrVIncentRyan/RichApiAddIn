<?php	
	$league = $_GET["league"];
	if ($_GET["type"] == "teams") {
		$url = "http://sports.services.appex.bing.com/LeagueV1.svc/LeagueStandings/$league/?lang=en-us";
	}
	else if ($_GET["type"] == "players" && isset($_GET["teamId"])) {
		$teamId = $_GET["teamId"];
		$url = "http://sports.services.appex.bing.com/TeamV1.svc/Roster/$league/$teamId/?lang=en-us";
	}
	$ch = curl_init($url);
	curl_setopt_array($ch, array(
		CURLOPT_RETURNTRANSFER => TRUE
	));
	
	$response = curl_exec($ch);
	
	if ($response === FALSE) {
		die(curl_error($ch));
	}
	
	if (isset($_GET["teamId"])) {
		$json = json_decode($response, true);
		$json["teamId"] = $_GET["teamId"];
		$response = json_encode($json);
	}
	
	echo $response;
?>