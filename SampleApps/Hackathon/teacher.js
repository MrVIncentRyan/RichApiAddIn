(function(){
  'use strict';

 // The initialize function is run each time the page is loaded.
Office.initialize = function (reason) {
    $(document).ready(function () {
        app.initialize();

		// Connect to server
		var socket = io.connect('https://linswudevbox2:8080', {secure: true});
		
		// Send notebookId to server
		socket.on('connect', function() {
			sendNotebookId(socket);
		});
		
        // Set up event handler for the UI.
		$('#createLessonPlan').click(handleCreateLessonPlan);
		$('#addNewTopic').click(function() {
			handleAddTopic();
		});
		$('#addNewQuiz').click(function() {
			handleAddQuiz();
		});
		$('#addAnswerButton').click(function() {
			handleAddAnswer();
		});
		$('#doneMakingQuizButton').click(function() {
			handleSaveQuiz(socket);
		});
		$('#saveTopic').click(function() {
			handleSaveTopic(socket);
		});
		$('#presentContent').click(function() {
			presentContentToStudents(socket);
		});
		
		socket.on('quizResponse', function(data) {
			handleQuizResponse(data);
		});
		
		updateLessonPlanList();
		
		$('#goback').click(function() {
			saveLessonPlan();
			updateLessonPlanList();
			document.getElementById("lessonplansection").style.display = "none";
			document.getElementById("defaultsection").style.display = "inline";
			document.getElementById("agendaheader").innerText = "";
			document.getElementById("agendaheader").innerText = "";
			document.getElementById("navigationchoice").style.display = "inline";
		});
    });
};

function navigateToLessonPlan(lessonPlanName)
{
	document.getElementById("navigationchoice").style.display = "none";
	var lessonPlanAgendaList = JSON.parse(localStorage.getItem(lessonPlanName));
	document.getElementById("lessonplansection").style.display = "inline";
	document.getElementById("defaultsection").style.display = "none";
	var agendaList = document.getElementById('agenda');
	for(var i = 0; i < lessonPlanAgendaList.length; i++)
	{
		var agendaListItem = document.createElement('li');
		var agendaListItemLink = document.createElement('a');
		agendaListItemLink.textContent = lessonPlanAgendaList[i].topicName;
		agendaListItemLink.setAttribute('data-navigation', lessonPlanAgendaList[i].topicLink);
		agendaListItemLink.onclick = function() { 
			navigateUser(socket, lessonPlanAgendaList[i].topicLink); 
		};
		agendaList.push(agendaListItem);
	}
}

function saveLessonPlan()
{
	var lessonAgendaList = [];
	var newLessonPlanName = document.getElementById('lessonPlanName').value;
	var agendaList = document.getElementById('agenda');
	var agendaListItems = agendaList.getElementsByTagName("li");
	for(var i=0; i < agendaListItems.length; i++)
	{
		var agendaListItemLink = agendaListItems[i].getElementsByTagName('a')[0];
		if (agendaListItemLink) {
			var agendaListItemObject = {topicName: agendaListItemLink.text, 
				topicLink: agendaListItemLink.getAttribute('data-navigation')};
			lessonAgendaList.push(agendaListItemObject);
		}
	}
	localStorage.setItem(newLessonPlanName, JSON.stringify(lessonAgendaList));
}

function updateLessonPlanList()
{
	var lessonPlanList = document.getElementById('lessonplanlist');
	$(lessonPlanList).empty();
	for(var storedLessonPlanName in localStorage)
	{
		if(storedLessonPlanName != "Office API client")
		{
			var lessonPlanListItem = document.createElement('li');
			var lessonPlanLink = document.createElement('a');
			lessonPlanLink.onclick = function(storedLessonPlanName) {
					navigateToLessonPlan(storedLessonPlanName); 
			};
			lessonPlanLink.innerText = storedLessonPlanName;
			lessonPlanListItem.appendChild(lessonPlanLink);
			lessonPlanList.appendChild(lessonPlanListItem);
		}
	}
}

// Present this outline to all students.
function handleCreateLessonPlan() {
	var agendaList = document.getElementById('agenda');
	$(agendaList).empty();
	document.getElementById("lessonplansection").style.display = "inline";
	document.getElementById("defaultsection").style.display = "none";
	document.getElementById("navigationchoice").style.display = "none";
}

function buildNavigationMessage(contentType, clientUrl, outlineId)
{
	var navigationMessage = '{ \"contentType\":\"' + contentType + '\", \"clientUrl\":\"' + clientUrl + '\", \"outlineId\":\"' + outlineId + '\" }';
	var navigationMessageObject = JSON.parse(JSON.stringify(navigationMessage));
	return navigationMessageObject;
}

function navigateUser(socket, navigationMessage)
{
	// Inform the students to navigate to the same location
	socket.emit('presentContent', navigationMessage);
	// Navigate the teacher to the location at this message
	parsePresentContentMessage(navigationMessage);
}

// Parse the json message from server for presenting content
function parsePresentContentMessage(message) {
	var contentType = message.contentType;
	var url = message.clientUrl;
	var id = message.outlineId;
	if (contentType == "page")
	{
		navigateToPage(url);
	} else {
		OneNote.run(function (ctx) {
			//check current active page
			var app = ctx.application;
			//navigate to page
			var page = app.activePage;
			ctx.load(page);
			ctx.sync()
				.then(function() {
					if (page.clientUrl != url) {
						navigateToPage(url);
						setTimeout(function() {
							setActiveOutline(id);
						}, 1000);
					}
					else {
						setActiveOutline(id);
					}
				});
		});
	}
}

// Navigate user to specific page
function navigateToPage(pageUrl){
	OneNote.run(function (ctx) {
		//check current active page
		var app = ctx.application;

		//navigate to page
		var page = app.navigateToPageWithClientUrl(pageUrl);
		ctx.load(page);
		return ctx.sync();
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
		if (!!activeOutline )
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
function handleAddTopic() {
	$('#newtopic').show();
	$('#newQuiz').hide();
}

function handleAddQuiz() {
	$('#newQuiz').show();
	$('#newtopic').hide();
}

function handleAddAnswer() {
	var inputDiv = $('<div></div>', {"class": "answerInput"});
	var radioInput = $('<input />', {"type": "radio", "name": "option"});
	var inputText = $('<input />', {"type": "text"});
	var removeX = $('<span></span>', {"class": "removeX"});
	removeX.html(" x");
	$('#answersDiv').append(inputDiv);
	inputDiv.append(radioInput);
	inputDiv.append(inputText);
	inputDiv.append(removeX);
	inputDiv.append($('</br>'));
	removeX.click(function() {
		inputDiv.remove();
	});
}

function handleSaveTopic(socket) {
    var topicName = document.getElementById("topicName").value;
	
	OneNote.run(function (context) {
	   var selectedContentType = $('input[name="contenttype2"]:checked').val();
        // Get the current page.
        var page = context.application.activePage;
		var outline = context.application.activeOutline;

        // Queue a command to load the page with the title property.             
        context.load(page, 'clientUrl,title'); 
		context.load(outline, 'id'); 
		
        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
            .then(function() {
			   var navigationMessage = {contentType: selectedContentType, clientUrl: page.clientUrl, outlineId: outline.id};
			   var agendaList = document.getElementById("agenda");
			   var topicLink = document.createElement("a");
			   topicLink.textContent = topicName;
			   topicLink.setAttribute('data-navigation', navigationMessage);
			   topicLink.onclick = function() { 
					navigateUser(socket, navigationMessage); 
			   };
			   var topicListItem = document.createElement('li');
			   topicListItem.appendChild(topicLink);
			   agendaList.appendChild(topicListItem);
			   document.getElementById("newtopic").style.display = "none";
            })
            .catch(function(error) {
                app.showNotification("Error: " + error); 
                console.log("Error: " + error); 
                if (error instanceof OfficeExtension.Error) { 
                    console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                } 
            }); 
        });
}

function handleSaveQuiz(socket) {
	var quizDiv = $('<div></div>');
	var quizItem = $('<li></li>');
	var quizLink = $('<a></a>');
	var quizAnswers = $('<div></div>', {"class": "quizAnswers"});
	quizAnswers.hide();
	var question = $('#newQuiz').children('textarea').first().val()
	quizLink.html(question);
	quizLink.click(function (){
		if (quizAnswers.is(":visible")) {
			quizAnswers.hide();
		}
		else {
			$('.quizAnswers').hide();
			quizAnswers.show();
		}
	});
	quizItem.append(quizLink);
	quizDiv.append(quizItem);
	$('#newQuiz').hide();

	var answers = [];
	var checked = 0;
	$.each($('#newQuiz').find('.answerInput'), function(index, input) {
		answers.push($(input).find('input[type=text]').first().val());
		if ($(input).find('input[type=radio]').first().is(":checked")) {
			checked = index;
		}
	});
	var answerList = $('<ul></ul');
	answerList.data("totalAnswers", 0);
	for (var i = 0; i < answers.length; i++) {
		var answerItemDiv = $('<div></div>', {"class": "answerItemDiv"});
		answerItemDiv.data("count", 0);
		var answer = $("<li>" + answers[i] + "</li>");
		var chartDiv = $('<div></div>');
		chartDiv.css({"height": "15px", "width": "250px", "padding": "0px"});
		var barDiv = $('<div></div>');
		barDiv.css({"height": "100%", "width": "0%", "background-color": "yellow"});
		if (i == checked) {
			answer.css({"font-weight": "bold"});
		}
		answerList.append(answerItemDiv);
		answerItemDiv.append(answer);
		answerItemDiv.append(chartDiv);
		chartDiv.append(barDiv);
	}
	var sendButton = $('<button>Send Quiz</button>');
	var endButton = $('<button>End Quiz</button>');
	endButton.hide();
	sendButton.click(function() {
		socket.emit('quiz', {question: question, answers: answers});
		$('#agenda').data("activeAnswerList", answerList);
		sendButton.hide();
		endButton.show();
	});
	endButton.click(function() {
		socket.emit('quizDone');
		$('#agenda').removeData("activeAnswerList");
		endButton.hide();
	});
	quizAnswers.append(answerList);
	quizAnswers.append(sendButton);
	quizAnswers.append(endButton);
	quizDiv.append(quizAnswers);
	
	$('#agenda').append(quizDiv);
}

function handleQuizResponse(data) {
	var answersList = $('#agenda').data("activeAnswerList");
	if (answersList) {
		var totalAnswers = answersList.data("totalAnswers") + 1;
		answersList.data("totalAnswers", totalAnswers);
		var answers = answersList.find('.answerItemDiv');
		var responseAnswer = $(answers[data]);
		responseAnswer.data("count", responseAnswer.data("count") + 1);
		$.each(answers, function(index, answer) {
			var count = $(answer).data("count");
			var chartDiv = $(answer).children('div').first();
			var barDiv = chartDiv.children('div').first();
			barDiv.css({"width": (count/totalAnswers*100) + "%"});
		});
	}
}

// Send notebookId to server and join room
function sendNotebookId(socket) {
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var notebook = app.activeNotebook;
		ctx.load(notebook);
		return ctx.sync()
			.then(function() {
				socket.emit('notebookId', notebook.id);
		});
	})
	.catch(function(error) {
		showError(error);
	});
}

// Present this outline to all students.
function presentContentToStudents(socket) {
    OneNote.run(function (context) {
	   var selectedContentType = $('input[name="contenttype"]:checked').val();
        // Get the current page.
        var page = context.application.activePage;
		var outline = context.application.activeOutline;

        // Queue a command to load the page with the title property.             
        context.load(page, 'clientUrl,title'); 
		context.load(outline, 'id'); 
		
        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
			.then(function() {
			   socket.emit('presentContent', {contentType: selectedContentType, clientUrl: page.clientUrl, outlineId: outline.id})
            })
            .catch(function(error) {
                showError(error);
            }); 
        });
}

function presentOutlineToStudents() {
	var i = 0;
	console.log("context menu clicked");  
	i++;
}

function showError(error) {
	app.showNotification("Error: " + error); 
	console.log("Error: " + error); 
	if (error instanceof OfficeExtension.Error) { 
		console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
	} 
}

})();
