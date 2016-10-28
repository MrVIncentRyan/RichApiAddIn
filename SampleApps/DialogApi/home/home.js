(function(){
	'use strict';

	var dialogHandler,  // Reference to the dialog box once opened
		pageCollection,  // Reference to the list of pages in the current section
		currentPageIndex = 0,  // Points to the page currently displayed in the dialog box
		NextPage = 1,  // Index offset to use when moving to the next page
		PreviousPage = -1;  // Index offset to use when moving to the previous page

	// Called when the apps for office infrastructure is ready for interaction
	Office.initialize = function (reason) {
		$(document).ready(
			function () {
				$('#openDialog').click(openDialog); 
			}
		);
	};

	function openDialog2() {
		OneNote.run(function (ctx) {		
			var app = ctx.application;
			app.load("_platform");
			return ctx.sync()
				.then(function() {
					console.log(app._platform);
				});
		});
	}
	
	// Launches the dialog
	function openDialog() {
		Office.context.ui.displayDialogAsync(
			'https://localhost:8443/app/home/dialog.html',
			{
				width: 80,
				height: 80,
				displayInIframe: false
			},
			function (dialogProxy) {
				dialogHandler = dialogProxy.value;
				dialogHandler.addEventHandler(
					Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
					dialogMessageHandler
				);
			}
		);
	}

	// Get all of the pages in the current section
	function getPages() {
		return OneNoteRichApi.Application.getActiveSectionAsync().then(
			function(section) {
				if (section != null) {
					return section.getPagesAsync().then(
						function (pages) {
							return pages;
						}
					);
				}
				else {
					return Promise.reject("There is no active page.");
				}
			}
		).catch(
			function(error) {
				console.log(error);
			}
		);
	}

	// Get the currently active page
	function getActivePage() {
		return OneNoteRichApi.Application.getActivePageAsync().then(
			function(page) {
				return page;
			}
		).catch(
			function(error) {
				console.log(error);
			}
		);
	}

	// Navigate to the specified page and obtain its content
	function getPageContent(page, skipNavigationToPage) {
		if (skipNavigationToPage) {
			return page.getStructureAsync();
		}
		else {
			return page.navigateAsync().then(
				function () {
					return page.getStructureAsync();
				}
			);
		}
	}

	// Sends the specified page content to the dialog
	function sendContentToDialog(pageContent) {
		dialogHandler.sendMessage(
			JSON.stringify(
				pageContent,
				function(key, value) {  // Handle circular references in the pageContent JSON
					if (key.indexOf('parent') === 0) {
						return 'parent';
					}
					else {
						return value;
					}
				}
			)
		);
	}

	// Loads the contents for the specified page
	function getContentForDialog(skipNavigationToPage) {
		getPageContent(pageCollection[currentPageIndex], skipNavigationToPage).then(
			function (pageContent) {
				return pageContent;
			}
		).then(sendContentToDialog);
	}

	// Moved forwards or backwards in the pages list and retrieved the content
	// for the new current page.
	function moveToPage(offset) {
		var numPages = pageCollection.length;

		currentPageIndex += offset;

		if (currentPageIndex >= numPages) {
			currentPageIndex = 0;
		}
		else if (currentPageIndex < 0) {
			currentPageIndex = numPages - 1;
		}

		return getContentForDialog();
	}

	// Handles messages received from the dialog (ready, next page, previous page etc)
	function dialogMessageHandler(event) {
		switch (event.message) {
			case Messages.DialogReady:

				getPages().then(
					function (pages) {
						pageCollection = pages;
						
						// Figure out which is the currently active page and set out page index to point to
						// it in the pageCollection.
						return getActivePage().then(
							function (activePage) {
								for (var i = 0; i < pageCollection.length; i++) {
									if (pageCollection[i].id == activePage.id) {
										currentPageIndex = i;
										break;
									}
								}

								// Load the content for the current page
								return getContentForDialog(true /* skipNavigationToPage */);
							}
						);
					}
				).catch(
					function(error) {
						console.log(error);
					}
				);

				break;

			case Messages.DialogMoveToPreviousPage:

				moveToPage(PreviousPage /* -1 */);
				break;

			case Messages.DialogMoveToNextPage:

				moveToPage(NextPage /* +1 */);
				break;
		}
	}
})();