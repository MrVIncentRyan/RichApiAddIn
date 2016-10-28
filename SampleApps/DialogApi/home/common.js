// Messages exchanged between the parent agave and the dialog box
var Messages = {
	DialogReady: 1,  // The dialog is ready to receive content
	DialogMoveToPreviousPage: 2,  // Dialog is asking for the next page's content
	DialogMoveToNextPage: 3  // Dialog is asking for the previous page's content
};