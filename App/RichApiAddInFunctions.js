Office.initialize = function() {};

function addPage(args) {
	var title = "Untitled Page";
	
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var section = app.getActiveSection();
		section.addPage(title);
		return ctx.sync();
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
		
	args.completed();
}