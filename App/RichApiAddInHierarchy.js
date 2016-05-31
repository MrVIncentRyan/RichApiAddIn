Office.initialize = function() {};

function getNotebook() {
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var notebook = app.getActiveNotebook();
		ctx.load(notebook);
		return ctx.sync()
			.then(function() {
				document.getElementById("activeNotebook").innerHTML = notebook.name;
				return ctx.sync();
			});
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
}

function getCurrentSection() {
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var section = app.getActiveSection();
		ctx.load(section);
		return ctx.sync()
			.then(function() {
				document.getElementById("activeSection").innerHTML = section.name;
				return ctx.sync();
			});
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
}

function getCurrentPage() {
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var page = app.getActivePage();
		ctx.load(page);
		return ctx.sync()
			.then(function() {
				document.getElementById("activePage").innerHTML = page.title;
				return ctx.sync();
			});
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
}

function showHierarchy() {
	OneNote.run(function(ctx) {
		var app = ctx.application;
		var notebook = app.getActiveNotebook();
		var sections = notebook.getSections(false);
		var sectionGroups = notebook.getSectionGroups(false);
		var notebookInnerDiv;
		var currentPageId;
		var display = document.getElementById("hierarchyDisplay");
		display.innerHTML = "";
		
		var currentPage = app.getActivePage();
		ctx.load(currentPage);
		ctx.sync()
			.then(function() {
				currentPageId = currentPage.id;
				return ctx.sync();
			})
		
		ctx.load(notebook);
		ctx.load(sections);
		ctx.load(sectionGroups);
		return ctx.sync()
			.then(function() {
				var notebookDiv = document.createElement("div");
				var paragraph = document.createElement("p");
				notebookInnerDiv = document.createElement("div");
				paragraph.className = "hierarchyCollapsible";
				notebookInnerDiv.className = "hierarchyChildDiv";
				paragraph.innerHTML = "+ " + notebook.name;
				paragraph.onclick = hideChildren;
				notebookDiv.appendChild(paragraph);
				notebookDiv.appendChild(notebookInnerDiv);
				display.appendChild(notebookDiv);
				return ctx.sync();
			})
			.then(function() {
				return getSectionsAndPages(ctx, sections, notebookInnerDiv, currentPageId);
			})
			.then(function() {
				return getSectionGroups(ctx, sectionGroups, notebookInnerDiv, currentPageId);
			});
	})
	.catch(function(error) {
		console.log("Error: " + JSON.stringify(error));
	});
}

function getSectionGroups(ctx, sectionGroups, parentDiv, currentPageId) {
	var sectionGroupToChildren = {};
	for (var i = 0; i < sectionGroups.items.length; i++) {
		var sectionGroup = sectionGroups.items[i];
		var div = document.createElement("div");
		var paragraph = document.createElement("p");
		var innerDiv = document.createElement("div");
		paragraph.className = "hierarchyCollapsible";
		innerDiv.className = "hierarchyChildDiv";
		paragraph.innerHTML = "+ " + sectionGroup.name;
		paragraph.onclick = hideChildren;
		div.appendChild(paragraph);
		div.appendChild(innerDiv);
		parentDiv.appendChild(div);
		var childSectionGroups = sectionGroup.getSectionGroups(false);
		var childSections = sectionGroup.getSections(false);
		sectionGroupToChildren[sectionGroup.name] = {sections: childSections, sectionGroups: childSectionGroups, parentDiv: innerDiv};
		ctx.load(childSectionGroups);
		ctx.load(childSections);
	}
	return ctx.sync().then(function() {
		return recursivelyAddSectionsAndSectionGroups(ctx, sectionGroupToChildren, currentPageId);
	});
}

function recursivelyAddSectionsAndSectionGroups(ctx, sectionGroupToChildren, currentPageId) {
	for (var child in sectionGroupToChildren) {
		var sections = sectionGroupToChildren[child].sections;
		var sectionGroups = sectionGroupToChildren[child].sectionGroups;
		var parentDiv = sectionGroupToChildren[child].parentDiv;
		return getSectionsAndPages(ctx, sections, parentDiv, currentPageId)
			.then(function() {
				return getSectionGroups(ctx, sectionGroups, parentDiv, currentPageId);
			})
			.then(function() {
				delete sectionGroupToChildren[child];
				return recursivelyAddSectionsAndSectionGroups(ctx, sectionGroupToChildren, currentPageId);
			})
			.catch(function(error) {
				console.log(JSON.stringify(error));
			});
	}
}

function getSectionsAndPages(ctx, sections, parentDiv, currentPageId) {
	var sectionsToPages = {};
	for (var i = 0; i < sections.items.length; i++) {
		var section = sections.items[i];
		sectionsToPages[section.name] = []
		var pages = section.getPages(false);
		ctx.load(pages);
		sectionsToPages[section.name].push(pages);
	}
	return ctx.sync().then(function() {
		for (var section in sectionsToPages) {
			if (sectionsToPages.hasOwnProperty(section)) {
				var div = document.createElement("div");
				var paragraph = document.createElement("p");
				var innerDiv = document.createElement("div");
				paragraph.className = "hierarchyCollapsible";
				innerDiv.className = "hierarchyChildDiv";
				paragraph.innerHTML = "+ " + section;
				paragraph.onclick = hideChildren;
				div.appendChild(paragraph);
				div.appendChild(innerDiv);
				parentDiv.appendChild(div);
				var pagesArray = sectionsToPages[section];
				for (var i = 0; i < pagesArray.length; i++) {
					var pages = pagesArray[i];
					for (var j = 0; j < pages.items.length; j++) {
						var page = pages.items[j];
						var div2 = document.createElement("div");
						var paragraph2 = document.createElement("p");
						paragraph2.innerHTML = "- " + page.title;
						if (page.id == currentPageId) {
							$(paragraph2).css("font-weight", "bold");
						}
						div2.appendChild(paragraph2);
						innerDiv.appendChild(div2);
					}
				}		
			}
		}
		return ctx.sync();
	});
}

function hideChildren() {
	var paragraph = $(this);
	var parentDiv = $(this).parent();
	parentDiv.children(".hierarchyChildDiv").each(function() {
		if ($(this).is(":visible")) {
			$(this).hide();
			paragraph.html("-" + paragraph.text().substring(1));
		}
		else {
			$(this).show();
			paragraph.html("+" + paragraph.text().substring(1));
		}
	});
}

function getParagraphs() {
	OneNote.run(function (ctx) {
	var application = ctx.application;
	var page = application.getActivePage();
	var pageContents = page.getContents();
	var paragraphs;
	ctx.load(pageContents, "type,outline,outline/paragraphs");
	return ctx.sync()
		.then(function() {
			pageContents.items.forEach(function(pageContent) {
			// Get the Outline objects from each PageContent.   
			if (pageContent.type === "Outline") {
				// Get the ParagraphCollection from each Outline. 
				paragraphs = pageContent.outline.paragraphs;
				ctx.load(paragraphs, "type,image/id,richText/text");
			}
			return ctx.sync()
				.then(function () {
					// Pass access to the loaded para
					for (var i = 0; i < pageContents.items.length; i++) {
						var paragraphList = pageContents.items[i].outline.paragraphs;
						console.log(paragraphList);
					}
				});
			});
		});
					
	});
}
