(function () {
	'use strict';

Office.initialize = function () {
	$(document).ready(function () {

		// Connect to server
		var socket = io.connect('https://linswudevbox2:8080', { secure: true });

		// Send notebookId to server
		socket.on('connect', function () {
			sendNotebookId(socket);
		});

		// When teacher presents content, navigate student
		socket.on('presentContent', function (data) {
			parsePresentContentMessage(data);
		});
		
		socket.on('quiz', function(data) {
			parseQuizMessage(data, socket);
		});
		
		socket.on('quizDone', function() {
			quizDone();
		});
	});
};

// Send notebookId to server and join room
function sendNotebookId(socket) {
	OneNote.run(function (ctx) {
		var app = ctx.application;
		var notebook = app.activeNotebook;
		ctx.load(notebook);
		return ctx.sync()
			.then(function () {
				socket.emit('notebookId', notebook.id);
			});
	})
	.catch(function (error) {
		console.log("Error: " + JSON.stringify(error));
	});
}

// Parse the json message from server for presenting content
function parsePresentContentMessage(message) {
	var contentType = message.contentType;
	var url = message.clientUrl;
	var id = message.outlineId;
	if (contentType == "page")
	{
		navigateToPage(url);
	} else
	{
		OneNote.run(function (ctx) {
			//check current active page
			var app = ctx.application;
			//navigate to page
			var page = app.activePage;
			ctx.load(page);
			return ctx.sync()
				.then(function () {
					if (page.clientUrl != url)
					{
						navigateToPage(url);
						setTimeout(function () {
							setActiveOutline(id);
						}, 1000);
					}
					else
					{
						setActiveOutline(id);
					}
				});
		});
	}
}

// Navigate user to specific page
function navigateToPage(pageUrl) {
	pageUrl = decodeURI(pageUrl);
	if (!pageUrl.includes("Class Notes.one"))
	{
		OneNote.run(function (ctx) {
			var app = ctx.application;
			var page = app.navigateToPageWithClientUrl(pageUrl);
			ctx.load(page);
			return ctx.sync().then(function () {
			});
		})
		.catch(function (error) {
			console.log("Error: " + JSON.stringify(error));
		});
	}
	else
	{
		navigateToDistributedPage(pageUrl);
	}
}

// Navigate student to their own distributed page
function navigateToDistributedPage(url) {
	var pageNameBeginIndex = url.indexOf("#");
	var pageNameEndIndex = url.indexOf("&");
	var pageName = decodeURI(url.substring(pageNameBeginIndex + 1, pageNameEndIndex));
	var sectionGroupName;

	OneNote.run(function (ctx) {
		var app = ctx.application;
		var notebook = app.activeNotebook;
		var sectionGroups;
		var sectionGroup;
		var sections;
		var pages;
		//load notebook 
		ctx.load(notebook);
		return ctx.sync().then(function () {
			sectionGroups = notebook.getSectionGroups();
			ctx.load(sectionGroups);
			return ctx.sync();
		})

		//load sections
		.then(function () {
			for (var i = 0; i < sectionGroups.items.length; i++)
			{
				sectionGroup = sectionGroups.items[i];
				var name = sectionGroup.name;
				if (!name.includes("_"))
				{
					sections = sectionGroup.getSections(true);
					ctx.load(sections);
					return ctx.sync();
				}
			}
		})

		//find section and load pages
		.then(function () {
			for (var j = 0; j < sections.items.length; j++)
			{
				var section = sections.items[j];
				var sectionName = section.name;
				if (sectionName == "Class Notes")
				{
					pages = section.getPages();
					ctx.load(pages);
					return ctx.sync();
				}
			}

		})
		.then(function () {
			for (var k = 0; k < pages.items.length; k++)
			{
				var page = pages.items[k];
				if (page.title === pageName)
				{
					var targetPage = app.navigateToPageWithClientUrl(page.clientUrl);
					ctx.load(targetPage);
					return ctx.sync().then(function () {
					});
				}
			}

		})
	})
	.catch(function (error) {
		console.log("Error: " + JSON.stringify(error));
	});
}


// Navigate user to specific outline
function setActiveOutline(Id) {
	OneNote.run(function (ctx) {
		var app = ctx.application;

		var activeOutline = app.activeOutline;
		if (!!activeOutline)
		{
			ctx.load(activeOutline);
			ctx.sync().then(function () {
				var activeOutlineId = activeOutline.id;
				if (activeOutlineId === Id)
				{
					return;
				}
			});
		}

		//set active outline
		var page = ctx.application.activePage;
		var pagecsCollection = [];
		var outlineCollection = [];
		var pagecs = page.getContents();
		ctx.load(pagecs);
		return ctx.sync()
			.then(function () {
				for (var i = 0; i < pagecs.items.length; i++)
				{
					var pagec = pagecs.items[i];
					if (pagec.type === "Outline")
					{
						ctx.load(pagec);
						pagecsCollection.push(pagec);
					}
				}
				return ctx.sync();
			})
			.then(function () {
				for (var j = 0; j < pagecsCollection.length; j++)
				{
					var outline = pagecsCollection[j].outline;
					ctx.load(outline);
					outlineCollection.push(outline);
				}
				return ctx.sync();
			})
			.then(function () {
				for (var k = 0; k < outlineCollection.length; k++)
				{
					var curOutline = outlineCollection[k]
					var outlineId = curOutline.id;
					if (outlineId === Id)
					{
						curOutline.select();
					}
					console.log(outlineId);
				}
				return ctx.sync();
			});
	})
	.catch(function (error) {
		console.log("Error: " + JSON.stringify(error));
	});
}
	
function parseQuizMessage(message, socket)
{
	var answers = message.answers;
	var question = message.question;
	var element = document.getElementById('quiz');
	//var answers = ["yes", "no", "ha", "la", "haha"];
	//var question = "hahaha";

	//Adding question 
	var questionBody = document.createElement("p");
	var questionText = document.createTextNode(question);
	questionBody.appendChild(questionText);
	element.appendChild(questionBody);

	//Adding answer choices
	for (var i = 0; i < answers.length; i++)
	{
		var choice = document.createElement("input");
		choice.type = "radio";
		choice.value = i.toString();
		choice.name = "options";
		var node = document.createTextNode(answers[i]);
		element.appendChild(choice);
		element.appendChild(node);
		var br = document.createElement("br");
		element.appendChild(br);
	}
	var button = document.createElement("input");
	button.type = "button";
	$(button).click(function() {
		submitAnswer(socket);
	});
	button.value = "submit";
	element.appendChild(button);
	return;
}

function submitAnswer(socket)
{
	var answers = document.getElementsByName('options');
	for (var i = 0; i < answers.length; i++)
	{
		if (answers[i].checked)
		{
			socket.emit('quizResponse', answers[i].value);
			$('#quiz').find('input[name=options]').attr('disabled', true);
			$('#quiz').find('input[type=button]').remove();
			return;
		}
	}
}

function quizDone() {
	var div = document.getElementById('quiz');
	while (div.firstChild)
	{
		div.removeChild(div.firstChild);
	}
}

})();