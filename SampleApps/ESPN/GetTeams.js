Office.initialize = function() {};

function getTeams(league) {
	var teamsArray = [];
	$.ajax({
		type: "GET",
		url: "GetTeamsApiCalls.php",
		data: 
		{
			"type": "teams",
			"league": league
		},
		success: function(jsonResponse) 
		{
			var response = $.parseJSON(jsonResponse);
			var teams = response.Results[0].Response.StandingsInfo;
			for (var i = 0; i < teams.length; i++) {
				var team = teams[i];
				teamsArray[team.Localid] = team;
				$.ajax({
					type: "GET",
					url: "GetTeamsApiCalls.php",
					data: 
					{
						"type": "players",
						"teamId": team.Localid,
						"league": league
					},
					success: function(jsonResponse) 
					{
						var response = $.parseJSON(jsonResponse);
						var teamId = response.teamId;
						var players = response.Results[0].Response.TeamPlayers;
						addTeamAndPlayers(teamsArray[teamId], players);
					}
				});	
			}
		}
	});	
}
	
function addTeamAndPlayers(team, players) {
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var notebook = app.activeNotebook;
		var section = notebook.addSection(team.Fullteamname);
		ctx.load(section);
		return ctx.sync().then(function() {
			return addPlayers(section, players, ctx, 0);
		});
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
}

function addPlayers(section, players, ctx, i) {
	if (i < players.length) {
		var player = players[i];
		var page = section.addPage(player.FullName);
		ctx.load(page);
		return ctx.sync().then(function() {
			var age = "Age: " + player.Age + "<br />";
			var height = "Height: " + player.Height + "<br />";
			var weight = "Weight" + player.Weight + "<br />";
			var experience = "Experience: " + player.Experience + " Years<br />";
			var position = "Position: " + player.Position + "<br />";
			var jersey = "Number: " + player.JerseyNumber + "<br />";
			var salary = "Salary: $" + player.Salary + "<br />";
			var imgSrc = player.PlayerImgSmall.SrcUrl;
			var info = "<img src='" + imgSrc + "' /><p>" + age + height + weight + experience + position + jersey + salary + "</p>";
			page.addOutline(50, 100, info);
			delete players[i];
			return addPlayers(section, players, ctx, i+1);
		})
		.catch(function(error) {
			console.log("Error: " + JSON.stringify(error));
		});
	}
	
}

function searchPages() {
	var pageTitle = $("#searchPageInput").val();
	$("#searchItems").html("");
	var teamToPlayers = [];
	
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var notebook = app.activeNotebook;
		var sections = notebook.getSections(false);
		ctx.load(sections);
		return ctx.sync()
			.then(function() {
				for (var i = 0; i < sections.items.length; i++) {
					var section = sections.items[i];
					var pages = section.getPages();
					ctx.load(pages);
					teamToPlayers.push({"section": section, "pages": pages});
				}
				return ctx.sync();
			})
			.then(function() {
				for (var i = 0; i < teamToPlayers.length; i++) {
					var pair = teamToPlayers[i];
					var section = pair.section;
					var pages = pair.pages;
					for (var j = 0; j < pages.items.length; j++) {
						var page = pages.items[j];
						if (page.title.toLowerCase().indexOf(pageTitle.toLowerCase()) >= 0) {
							var playerDiv = $("<p>" + page.title + " - " + section.name + "</p>");
							playerDiv.data("id", page.id);
							var a = playerDiv.data("id");
							playerDiv.css("cursor", "pointer");
							playerDiv.click(goToClickedPage);
							$("#searchItems").append(playerDiv);
						}
					}
				}
				return ctx.sync();
			});
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
}

function goToClickedPage() {
	var clickedPage = $(this);
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var notebook = app.activeNotebook;
		var page = notebook.getPageById(clickedPage.data("id"));
		app.navigateToPage(page);
		return ctx.sync();
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
}