Office.initialize = function() {
	
};

function prettyCode() {
	Office.context.document.getSelectedDataAsync(
		Office.CoercionType.Text, 
		function (asyncResult) {
			$.ajax({
				type: "POST",
				url: "BeautifyCode.php",
				data: 
				{
					"code": asyncResult.value
				},
				success: function(jsonResponse) 
				{
					var response = $.parseJSON(jsonResponse);
					var code = response.join("<br/>");
					
					var prettyInner = $("<pre class='prettyprint'></pre>").get(0);
					prettyInner.innerHTML = code;
					var pretty = $("<div></div>").get(0);
					pretty.innerHTML = prettyInner.outerHTML;
					PR.prettyPrint(function() {}, pretty);	
					
					$.ajax({
						type: "POST",
						url: "GetInlineStyles.php",
						data: 
						{
							"html": pretty.outerHTML,
							"cssFile": "Styles.css"
						},
						success: function(response) 
						{
							Office.context.document.setSelectedDataAsync(response, {coercionType: Office.CoercionType.Html}, function(result) {});
						}
					});	
				}
			});
		}
	);
}