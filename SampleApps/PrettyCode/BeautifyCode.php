<?php
	$code = $_POST["code"];
	$output = array();
	exec("python BeautifyCode.py \"$code\"",  $output);
	echo json_encode($output)
?>