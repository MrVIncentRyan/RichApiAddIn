Office.initialize = function() {
	getResults();
	Office.context.document.addHandlerAsync(
		Office.EventType.DocumentSelectionChanged,
		function() {
			if ($("#selectionActivatedButton").html() == "Stop") {
				getResults();
			}
		},
		function (asyncResult) {}
	);
};

function selectionActivate() {
	var but = $("#selectionActivatedButton");
	if (but.html() == "Stop") {
		but.html("Start");
	}
	else {
		but.html("Stop");
	}
}

function getResults() {
	$("#results").html("");
	Office.context.document.getSelectedDataAsync(
		Office.CoercionType.Text, 
		function (asyncResult) {
			if (asyncResult.value) {
				$("#selectedText").html(asyncResult.value);
				$.ajax({
					type: "GET",
					url: "CallWolframAlpha.php",
					data: 
					{
						"input": encodeURIComponent(asyncResult.value),
						"format": "image,plaintext"
					},
					success: function(response) 
					{
						console.log(response);
						$xml = $($.parseXML(response));
						$xml.find("pod").each(function() {
							var pod = $("<div/>", {"class": "wolframPod"});
							var title = $(this).attr("title");
							pod.append("<h3>" + title + ":</h3>");
							$("#results").append(pod);
							$(this).find("subpod").each(function() {
								var imgNode = $(this).find("img:first")[0].outerHTML;
								console.log(imgNode);
								var plaintext = $(this).find("plaintext:first").text();
								var img = $(imgNode);
								img.hover(function() {
									$(this).css("cursor", "pointer");
								});
								img.click(function() {
									OneNote.run(function(ctx) {
										var app = ctx.application;
										var outline = app.activeOutline;
										outline.append(plaintext);
										return ctx.sync();
									})
									.catch(function(error) {
										console.log("Error: " + JSON.stringify(error));
									});
								});
								pod.append(img);
							});
						});
					}
				});
			}
			else {
				$("#selectedText").html("Please select some text!");
			}
		}
	);
}

$(document)
.ajaxStart(function () {
	console.log("A");
	$("#loadingImage").show();
})
.ajaxStop(function () {
	console.log("B");
	$("#loadingImage").hide();
})
.ajaxError(function () {
	$("#loadingImage").hide();
});

function doIt() {
	Office.context.document.getSelectedDataAsync(
		Office.CoercionType.Matrix, 
		function (asyncResult) {
			console.log(asyncResult.value);
		}
	);
}