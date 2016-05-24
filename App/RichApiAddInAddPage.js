Office.initialize = function() {};
  
function insertPage() {
	var title = $('#addPageInput').val();
	if (!title) {
		title = "Untitled Page"
	}
	
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var section = app.getActiveSection();
		var page = app.getActivePageOrNull();
		switch ($('input[name=insertPagePosition]:checked').val()) {
			case "before":
				page.insertPageAsSibling(0, title);
				break;
			case "after":
				page.insertPageAsSibling(1, title);
				break;
			case "last":
				section.addPage(title);
				break;
		}
		return ctx.sync();
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
}

function insertSection() {
	var title = $('#addPageInput').val();
	if (!title) {
		return;
	}
	
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var notebook = app.getActiveNotebook();
		var section = app.getActiveSection();
		switch ($('input[name=insertPagePosition]:checked').val()) {
			case "before":
				section.insertSectionAsSibling(0, title);
				break;
			case "after":
				section.insertSectionAsSibling(1, title);
				break;
			case "last":
				notebook.addSection(title);
				break;
		}
		return ctx.sync();
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
}

function insertOutline() {
	var x = parseFloat($('#addPageX').val());
	var y = parseFloat($('#addPageY').val());
	var html = $('#addOutlineInput').val();
	if (isNaN(x) || isNaN(y) || !html) {
		return;
	}
	
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var page = app.getActivePageOrNull();
		page.addOutline(x, y, html);
		return ctx.sync();
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
}

function setHTML() {
	var html = $('#addHTMLInput').val();
	Office.context.document.setSelectedDataAsync(html, {coercionType: Office.CoercionType.Html}, function(result) {});
}