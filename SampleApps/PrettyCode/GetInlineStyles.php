<?php
	include("Emogrifier.php");
	if (isset($_POST["html"]) && isset($_POST["cssFile"])) {
		$html = $_POST["html"];
		$cssFile = $_POST["cssFile"];
		$css = file_get_contents($cssFile);
		$emogrifier = new \Pelago\Emogrifier();
		$emogrifier->setHtml($html);
		$emogrifier->setCss($css);
		$result = $emogrifier->emogrifyBodyContent();
		echo $result;
	}
	else {
		echo "Properties not set";
	}
?>