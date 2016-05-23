Office.initialize = function() {
	$(document).ready(function () {
		$("body").hide();
		var cookies = document.cookie.split(";");
		var consumer_key;
		var code;
		for (var i = 0 ; i < cookies.length; i++) {
			if (cookies[i].split("=")[0].trim() == "consumer_key") {
				consumer_key = cookies[i].split("=")[1];
			}
			else if (cookies[i].split("=")[0].trim() == "code") {
				code = cookies[i].split("=")[1];
			}
		}
		if (consumer_key && code) {
			getAccessToken(consumer_key, code);
		}
		else {
			getRequestToken();
		}
	});
};

function getRequestToken() {
	var consumer_key = "52338-681593ae8cb0ecb38c029fb0";
	var redirect_uri = "https://richapiaddin.azurewebsites.net/SampleApps/Pocket/PocketTest.html";
	var postResult = $.ajax({
		type: "POST",
		url: "PocketGetRequestToken.php",
		data: 
		{
			"consumer_key": consumer_key,
			"redirect_uri": redirect_uri
		},
		success: function(jsonResponse) 
		{
			var response = $.parseJSON(jsonResponse);
			var code = response.code;
			var d = new Date();
			d.setMinutes(d.getMinutes() + 1);
			document.cookie = "consumer_key=".concat(consumer_key);
			document.cookie = "code=".concat(code, ";expires=", d.toUTCString());
			authorizePocket(code, redirect_uri);
		}
	});
}

function authorizePocket(code, redirect_uri) {
	var url = 'https://getpocket.com/auth/authorize?request_token='.concat(code, '&redirect_uri=', redirect_uri);
	window.location.replace(url);
}

function getAccessToken(consumer_key, code) {
	var postResult = $.ajax({
		type: "POST",
		url: "PocketGetAccessToken.php",
		data: 
		{
			"consumer_key": consumer_key,
			"code": code
		},
		success: function(jsonResponse) 
		{
			var response = $.parseJSON(jsonResponse);
			var token = response.access_token;
			var username = response.username;
			var d = new Date();
			d.setDate(d.getDate() - 1);
			document.cookie = "code=".concat(code, ";expires=", d.toUTCString());
			document.cookie = "access_token=".concat(token);
			showUI(username);
		}
	});
}

function retrievePocketStuff(action) {
	var cookies = document.cookie.split(";");
	var consumer_key;
	var access_token;
	for (var i = 0 ; i < cookies.length; i++) {
		if (cookies[i].split("=")[0].trim() == "consumer_key") {
			consumer_key = cookies[i].split("=")[1];
		}
		else if (cookies[i].split("=")[0].trim() == "access_token") {
			access_token = cookies[i].split("=")[1];
		}
	}
	
	if (consumer_key && access_token) {
		var postResult = $.ajax({
			type: "POST",
			url: "PocketRetrieveReadings.php",
			data: 
			{
				"consumer_key": consumer_key,
				"access_token": access_token
			},
			success: function(jsonResponse) 
			{
				var response = $.parseJSON(jsonResponse);
				addReadingPages(response.list);
			}
		});
	}
}

function addReadingPages(list) {
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var section = app.activeSection;
		ctx.load(section);
		return ctx.sync()
			.then(function() {
				return recursivelyAddPages(list, section, ctx);
			});
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
}

function recursivelyAddPages(list, section, ctx) {
	for (var key in list) {
		var item = list[key];
		var page = section.addPage(item.given_title);
		ctx.load(page);
		return ctx.sync()
			.then(function() {
				delete list[key];
				page.addOutline(50, 100, item.excerpt + "\n" + item.given_url);
				return recursivelyAddPages(list, section, ctx);
			})
			.catch(function(error) {
				console.log("Error: " + JSON.stringify(error));
			});
	}
}

function addPocketSection() {
	var title = "Pocket Section";
	
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var notebook = app.activeNotebook;
		notebook.addSection(title);
		return ctx.sync();
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
}

function showUI(username) {
	$("body").show();
	$("#welcomeHeader").text("Welcome " + username);
}