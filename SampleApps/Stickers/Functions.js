Office.initialize = function() {};

function goodJobSticker1(args) {
	var html = '<img src="https://richapiaddin.azurewebsites.net/SampleApps/Stickers/Images/GoodJobSticker1.png" >';
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var page = app.activePage;
		page.addOutline(200, 100,  html);
		return ctx.sync();
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
	args.completed();
}

function goodJobSticker2(args) {
	var html = '<img src="https://richapiaddin.azurewebsites.net/SampleApps/Stickers/Images/GoodJobSticker2.gif" >';
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var page = app.activePage;
		page.addOutline(200, 100,  html);
		return ctx.sync();
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
	args.completed();
}

function goodJobSticker3(args) {
	var html = '<img src="https://richapiaddin.azurewebsites.net/SampleApps/Stickers/Images/GoodJobSticker3.gif" >';
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var page = app.activePage;
		page.addOutline(200, 100,  html);
		return ctx.sync();
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
	args.completed();
}

function inspirationalSticker1(args) {
	var html = '<img src="https://richapiaddin.azurewebsites.net/SampleApps/Stickers/Images/InspirationalSticker1.jpg" >';
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var page = app.activePage;
		page.addOutline(200, 100,  html);
		return ctx.sync();
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
	args.completed();
}

function inspirationalSticker2(args) {
	var html = '<img src="https://richapiaddin.azurewebsites.net/SampleApps/Stickers/Images/InspirationalSticker2.jpg" >';
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var page = app.activePage;
		page.addOutline(200, 100,  html);
		return ctx.sync();
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
	args.completed();
}

function inspirationalSticker3(args) {
	var html = '<img src="https://richapiaddin.azurewebsites.net/SampleApps/Stickers/Images/InspirationalSticker3.jpg" >';
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var page = app.activePage;
		page.addOutline(200, 100,  html);
		return ctx.sync();
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
	args.completed();
}