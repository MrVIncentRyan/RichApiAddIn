(function() {
  'use strict';

  	// Handle messages received from the parent agave
	function handleParentMessage(event) {
		$('#content').html(event.message);  // Show the page content sent from the parent agave
	}

	// Send a message to the parent agave
	function sendMessageToParent(messageId) {
		Office.context.ui.messageParent(messageId);
	}

	// Used to generate callback delegates from click events
	function createSendMessageToParentCallback(messageId) {
		return function () {
			sendMessageToParent(messageId);
		}
	}

	Office.initialize = function (reason) {
		// Register for parent agave messages
		Office.context.ui.addHandlerAsync(
			Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived,
			handleParentMessage);

		sendMessageToParent(Messages.DialogReady);  // Notify the parent agave the dialog is ready to receive content
		
		// Hook click events for the next and previous buttons
		$('#previous').click(createSendMessageToParentCallback(Messages.DialogMoveToPreviousPage));
		$('#next').click(createSendMessageToParentCallback(Messages.DialogMoveToNextPage));
	};
})();
